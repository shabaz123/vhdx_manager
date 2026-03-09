import ctypes
import json
import os
import re
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import messagebox
from typing import Any, Dict, List, Optional, Tuple


APP_TITLE = "VHDX Manager"
JSON_FILE = "vhdx_list.json"
POWERSHELL = ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command"]
DEFAULT_VHD_BASE = Path("C:/LOCAL_VHD")
DEFAULT_MOUNT_BASE = Path("C:/DEV/vhd_mounts")

COLORS = {
    "bg": "#0D1117",
    "panel": "#161B22",
    "panel_alt": "#21262D",
    "text": "#E6EDF3",
    "muted": "#8B949E",
    "border": "#30363D",
    "processing": "#000000",
    "Mounted": "#3FB950",
    "Unmounted": "#8B949E",
    "Unavailable": "#D29922",
    "Unhealthy": "#F85149",
}

STATE_ORDER = ["Mounted", "Unmounted", "Unavailable", "Unhealthy"]


@dataclass
class VHDEntry:
    vhd_path: str
    vhd_volume_label: str
    vhd_description: str


class PowerShellError(RuntimeError):
    pass


class VHDManagerApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("980x640")
        self.minsize(760, 420)
        self.configure(bg=COLORS["bg"])

        self.entries: List[VHDEntry] = []
        self.row_widgets: Dict[str, Dict[str, Any]] = {}
        self.is_busy = False
        self.create_dialog: Optional[tk.Toplevel] = None

        self._build_ui()
        self.refresh_all()

    def _build_ui(self) -> None:
        header = tk.Frame(self, bg=COLORS["bg"])
        header.pack(fill="x", padx=18, pady=(16, 10))

        title = tk.Label(
            header,
            text=APP_TITLE,
            bg=COLORS["bg"],
            fg=COLORS["text"],
            font=("Segoe UI", 20, "bold"),
        )
        title.pack(side="left")

        subtitle = tk.Label(
            header,
            text="Click a bullet to mount, unmount, or detach an unhealthy VHDX.",
            bg=COLORS["bg"],
            fg=COLORS["muted"],
            font=("Segoe UI", 10),
        )
        subtitle.pack(side="left", padx=(12, 0), pady=(6, 0))

        controls = tk.Frame(self, bg=COLORS["bg"])
        controls.pack(fill="x", padx=18, pady=(0, 10))

        self.status_var = tk.StringVar(value="Ready")
        self.debug_var = tk.StringVar(value="")

        self.status_label = tk.Label(
            controls,
            textvariable=self.status_var,
            bg=COLORS["bg"],
            fg=COLORS["muted"],
            font=("Segoe UI", 10),
        )
        self.status_label.pack(side="left")

        self.debug_label = tk.Label(
            controls,
            textvariable=self.debug_var,
            bg=COLORS["bg"],
            fg=COLORS["muted"],
            font=("Consolas", 9),
        )
        self.debug_label.pack(side="left", padx=(16, 0))

        button_bar = tk.Frame(controls, bg=COLORS["bg"])
        button_bar.pack(side="right")

        self.create_button = tk.Button(
            button_bar,
            text="Create New VHDX",
            command=self.open_create_dialog,
            bg=COLORS["panel_alt"],
            fg=COLORS["text"],
            activebackground=COLORS["panel_alt"],
            activeforeground=COLORS["text"],
            relief="flat",
            bd=0,
            padx=14,
            pady=8,
            font=("Segoe UI", 10, "bold"),
            cursor="hand2",
        )
        self.create_button.pack(side="right", padx=(0, 10))

        self.refresh_button = tk.Button(
            button_bar,
            text="Refresh",
            command=self.refresh_all,
            bg=COLORS["panel_alt"],
            fg=COLORS["text"],
            activebackground=COLORS["panel_alt"],
            activeforeground=COLORS["text"],
            relief="flat",
            bd=0,
            padx=14,
            pady=8,
            font=("Segoe UI", 10, "bold"),
            cursor="hand2",
        )
        self.refresh_button.pack(side="right")

        legend = tk.Frame(self, bg=COLORS["bg"])
        legend.pack(fill="x", padx=18, pady=(0, 10))
        for state in STATE_ORDER:
            item = tk.Frame(legend, bg=COLORS["bg"])
            item.pack(side="left", padx=(0, 14))
            swatch = tk.Canvas(item, width=14, height=14, bg=COLORS["bg"], highlightthickness=0)
            swatch.create_oval(2, 2, 12, 12, fill=COLORS[state], outline=COLORS[state])
            swatch.pack(side="left")
            tk.Label(
                item,
                text=state,
                bg=COLORS["bg"],
                fg=COLORS["muted"],
                font=("Segoe UI", 9),
            ).pack(side="left", padx=(6, 0))

        list_container = tk.Frame(self, bg=COLORS["bg"])
        list_container.pack(fill="both", expand=True, padx=18, pady=(0, 18))

        self.canvas = tk.Canvas(
            list_container,
            bg=COLORS["bg"],
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            bd=0,
        )
        self.canvas.pack(side="left", fill="both", expand=True)

        scrollbar = tk.Scrollbar(list_container, orient="vertical", command=self.canvas.yview)
        scrollbar.pack(side="right", fill="y")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.scroll_frame = tk.Frame(self.canvas, bg=COLORS["bg"])
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")

        self.scroll_frame.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _on_frame_configure(self, _event: tk.Event) -> None:
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event: tk.Event) -> None:
        self.canvas.itemconfigure(self.canvas_window, width=event.width)

    def _on_mousewheel(self, event: tk.Event) -> None:
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def set_busy(self, busy: bool, message: str) -> None:
        self.is_busy = busy
        state = "disabled" if busy else "normal"
        self.refresh_button.configure(state=state)
        self.create_button.configure(state=state)
        self.status_var.set(message)

    def set_dbgtext(self, message: str) -> None:
        self.debug_var.set(message)

    def load_entries(self) -> List[VHDEntry]:
        json_path = Path(JSON_FILE)
        if not json_path.exists():
            raise FileNotFoundError(f"{JSON_FILE} was not found next to the application.")

        with json_path.open("r", encoding="utf-8") as f:
            raw = json.load(f)

        if not isinstance(raw, list):
            raise ValueError(f"{JSON_FILE} must contain a JSON array.")

        entries: List[VHDEntry] = []
        for i, item in enumerate(raw, start=1):
            if not isinstance(item, dict):
                raise ValueError(f"Entry {i} in {JSON_FILE} is not an object.")

            vhd_path = str(item.get("vhd_path", "")).strip()
            vhd_volume_label = str(item.get("vhd_volume_label", "")).strip()
            vhd_description = str(item.get("vhd_description", "")).strip()

            if not vhd_path or not vhd_volume_label or not vhd_description:
                raise ValueError(
                    f"Entry {i} must include non-empty 'vhd_path', 'vhd_volume_label', and 'vhd_description' fields."
                )

            entries.append(
                VHDEntry(
                    vhd_path=os.path.expandvars(vhd_path),
                    vhd_volume_label=vhd_volume_label,
                    vhd_description=vhd_description,
                )
            )
        return entries

    def refresh_all(self) -> None:
        if self.is_busy:
            return
        self.set_busy(True, "Refreshing VHD states...")
        threading.Thread(target=self._refresh_worker, daemon=True).start()

    def _refresh_worker(self) -> None:
        try:
            entries = self.load_entries()
            volume_index = self.get_volume_index()
            disk_index = self.get_disk_image_index()
            states = [self.compute_state(entry, volume_index, disk_index) for entry in entries]
            self.after(0, lambda: self._render(entries, states))
            self.after(0, lambda: self.set_busy(False, f"Loaded {len(entries)} VHDX item(s)."))
        except Exception as exc:
            self.after(0, lambda: self.set_busy(False, "Failed to refresh."))
            self.after(0, lambda: messagebox.showerror(APP_TITLE, str(exc)))

    def _render(self, entries: List[VHDEntry], states: List[Dict[str, Any]]) -> None:
        self.entries = entries
        self.row_widgets.clear()
        for child in self.scroll_frame.winfo_children():
            child.destroy()

        for idx, (entry, state) in enumerate(zip(entries, states)):
            row_bg = COLORS["panel"] if idx % 2 == 0 else COLORS["panel_alt"]
            row = tk.Frame(
                self.scroll_frame,
                bg=row_bg,
                highlightthickness=1,
                highlightbackground=COLORS["border"],
                padx=14,
                pady=12,
            )
            row.pack(fill="x", pady=(0, 8))

            bullet = tk.Canvas(row, width=26, height=26, bg=row_bg, highlightthickness=0, cursor="hand2")
            bullet.pack(side="left", padx=(0, 12))
            oval = bullet.create_oval(4, 4, 22, 22, fill=COLORS[state["state"]], outline=COLORS[state["state"]])
            bullet.bind("<Button-1>", lambda _e, p=entry.vhd_path: self.on_bullet_click(p))

            body = tk.Frame(row, bg=row_bg)
            body.pack(side="left", fill="x", expand=True)

            top_line = tk.Frame(body, bg=row_bg)
            top_line.pack(fill="x")

            title = tk.Label(
                top_line,
                text=entry.vhd_description,
                bg=row_bg,
                fg=COLORS["text"],
                font=("Segoe UI", 12, "bold"),
                anchor="w",
            )
            title.pack(side="left")

            state_label = tk.Label(
                top_line,
                text=state["state"],
                bg=row_bg,
                fg=COLORS[state["state"]],
                font=("Segoe UI", 10, "bold"),
                anchor="e",
            )
            state_label.pack(side="right")

            desc = tk.Label(
                body,
                text=entry.vhd_volume_label,
                bg=row_bg,
                fg=COLORS["muted"],
                font=("Segoe UI", 10),
                anchor="w",
            )
            desc.pack(fill="x", pady=(3, 0))

            path_label = tk.Label(
                body,
                text=entry.vhd_path,
                bg=row_bg,
                fg=COLORS["muted"],
                font=("Consolas", 10),
                anchor="w",
            )
            path_label.pack(fill="x", pady=(4, 0))

            detail_text = state.get("detail", "")
            if detail_text:
                detail = tk.Label(
                    body,
                    text=detail_text,
                    bg=row_bg,
                    fg=COLORS["muted"],
                    font=("Segoe UI", 9),
                    anchor="w",
                )
                detail.pack(fill="x", pady=(2, 0))

            self.row_widgets[entry.vhd_path.lower()] = {
                "canvas": bullet,
                "oval": oval,
                "state_label": state_label,
                "entry": entry,
            }

    def open_create_dialog(self) -> None:
        if self.create_dialog is not None and self.create_dialog.winfo_exists():
            self.create_dialog.lift()
            self.create_dialog.focus_force()
            return

        dialog = tk.Toplevel(self)
        dialog.title("Create New VHDX")
        dialog.configure(bg=COLORS["bg"])
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()
        self.create_dialog = dialog

        content = tk.Frame(dialog, bg=COLORS["bg"], padx=18, pady=18)
        content.pack(fill="both", expand=True)

        dialog_vars = {
            "title": tk.StringVar(value="Bob Application"),
            "vhd_path": tk.StringVar(value=r"C:\LOCAL_VHD\bob.vhdx"),
            "volume_label": tk.StringVar(value="BOB_DEV"),
            "max_size_gb": tk.StringVar(value="30"),
            "mount_path": tk.StringVar(value=r"C:\DEV\vhd_mounts\bob"),
            "error_text": tk.StringVar(value=""),
        }
        setattr(dialog, "dialog_vars", dialog_vars)

        fields = [
            ("Title", dialog_vars["title"]),
            ("VHDX File Path", dialog_vars["vhd_path"]),
            ("Volume Label", dialog_vars["volume_label"]),
            ("Max Size [GB]", dialog_vars["max_size_gb"]),
            ("Dest Mount Path", dialog_vars["mount_path"]),
        ]

        for row_index, (label_text, var) in enumerate(fields):
            label = tk.Label(
                content,
                text=label_text,
                bg=COLORS["bg"],
                fg=COLORS["text"],
                font=("Segoe UI", 10, "bold"),
                anchor="w",
            )
            label.grid(row=row_index, column=0, sticky="w", pady=(0, 6))

            entry = tk.Entry(
                content,
                textvariable=var,
                bg=COLORS["panel_alt"],
                fg=COLORS["text"],
                insertbackground=COLORS["text"],
                relief="flat",
                width=48,
                font=("Segoe UI", 10),
            )
            entry.grid(row=row_index, column=1, sticky="ew", padx=(12, 0), pady=(0, 6))

        content.columnconfigure(1, weight=1)

        error_label = tk.Label(
            content,
            textvariable=dialog_vars["error_text"],
            bg=COLORS["bg"],
            fg=COLORS["Unhealthy"],
            font=("Segoe UI", 10, "bold"),
            anchor="w",
            justify="left",
            wraplength=520,
        )
        error_label.grid(row=len(fields), column=0, columnspan=2, sticky="w", pady=(6, 10))

        button_row = tk.Frame(content, bg=COLORS["bg"])
        button_row.grid(row=len(fields) + 1, column=0, columnspan=2, sticky="e")

        ok_button = tk.Button(
            button_row,
            text="OK",
            command=lambda: self.submit_create_dialog(dialog),
            bg=COLORS["panel_alt"],
            fg=COLORS["text"],
            activebackground=COLORS["panel_alt"],
            activeforeground=COLORS["text"],
            relief="flat",
            bd=0,
            padx=14,
            pady=8,
            font=("Segoe UI", 10, "bold"),
            cursor="hand2",
        )
        ok_button.pack(side="right", padx=(10, 0))

        cancel_button = tk.Button(
            button_row,
            text="Cancel",
            command=lambda: self.close_create_dialog(dialog),
            bg=COLORS["panel_alt"],
            fg=COLORS["text"],
            activebackground=COLORS["panel_alt"],
            activeforeground=COLORS["text"],
            relief="flat",
            bd=0,
            padx=14,
            pady=8,
            font=("Segoe UI", 10, "bold"),
            cursor="hand2",
        )
        cancel_button.pack(side="right")

        dialog.protocol("WM_DELETE_WINDOW", lambda: self.close_create_dialog(dialog))
        dialog.bind("<Return>", lambda _event: self.submit_create_dialog(dialog))
        dialog.bind("<Escape>", lambda _event: self.close_create_dialog(dialog))

    def close_create_dialog(self, dialog: tk.Toplevel) -> None:
        if dialog.winfo_exists():
            dialog.grab_release()
            dialog.destroy()
        self.create_dialog = None

    def submit_create_dialog(self, dialog: tk.Toplevel) -> None:
        dialog_vars = getattr(dialog, "dialog_vars")
        title_text = dialog_vars["title"].get().strip()
        vhd_path_str = dialog_vars["vhd_path"].get().strip()
        volume_label = dialog_vars["volume_label"].get().strip()
        max_size_str = dialog_vars["max_size_gb"].get().strip()
        mount_path_str = dialog_vars["mount_path"].get().strip()

        if not title_text:
            dialog_vars["error_text"].set("Error, Title must not be empty")
            return

        vhd_path = Path(vhd_path_str)
        mount_path = Path(mount_path_str)

        if not vhd_path.parent.exists() or not mount_path.exists():
            dialog_vars["error_text"].set("Error, folder does not exist")
            return

        try:
            max_size_gb = int(max_size_str)
        except ValueError:
            dialog_vars["error_text"].set("Error, Max Size [GB] must be a whole number")
            return

        if max_size_gb <= 0 or max_size_gb > 200:
            dialog_vars["error_text"].set("Error, Max Size [GB] must be between 1 and 200")
            return

        save_ok, save_message = self.validate_json_append(title_text, vhd_path_str, volume_label)
        if not save_ok:
            dialog_vars["error_text"].set(save_message)
            return

        dialog_vars["error_text"].set("")
        self.set_busy(True, f"Creating {vhd_path.name}...")
        threading.Thread(
            target=self._create_vhd_worker,
            args=(dialog, title_text, vhd_path_str, volume_label, max_size_gb, mount_path_str),
            daemon=True,
        ).start()

    def _create_vhd_worker(
        self,
        dialog: tk.Toplevel,
        title_text: str,
        vhd_path: str,
        volume_label: str,
        max_size_gb: int,
        mount_path: str,
    ) -> None:
        success, message = create_dynamic_vhdx_diskpart_safe(
            vhd_path=vhd_path,
            volume_label=volume_label,
            max_size_gb=max_size_gb,
            mount_folder=mount_path,
            allowed_vhd_base=str(DEFAULT_VHD_BASE),
            allowed_mount_base=str(DEFAULT_MOUNT_BASE),
        )

        if success:
            save_ok, save_message = self.append_entry_to_json(title_text, vhd_path, volume_label)
            if not save_ok:
                self.after(0, lambda: self.set_busy(False, "Create VHDX succeeded, but JSON update failed."))
                self.after(0, lambda: self._set_dialog_error(dialog, save_message))
                return
            self.after(0, lambda: self.close_create_dialog(dialog))
            self.after(0, lambda: self._finish_create_success(message))
        else:
            self.after(0, lambda: self.set_busy(False, "Create VHDX failed."))
            self.after(0, lambda: self._set_dialog_error(dialog, message))

    def _finish_create_success(self, message: str) -> None:
        self.set_busy(False, message)
        self.refresh_all()

    def _set_dialog_error(self, dialog: tk.Toplevel, message: str) -> None:
        if dialog.winfo_exists():
            dialog_vars = getattr(dialog, "dialog_vars")
            dialog_vars["error_text"].set(message)
            dialog.lift()
            dialog.focus_force()

    def validate_json_append(self, title_text: str, vhd_path: str, volume_label: str) -> Tuple[bool, str]:
        if not title_text:
            return False, "Error, Title must not be empty"

        json_path = Path(JSON_FILE)
        if not json_path.exists():
            return False, f"{JSON_FILE} was not found next to the application."

        try:
            with json_path.open("r", encoding="utf-8") as f:
                raw = json.load(f)
        except Exception as exc:
            return False, str(exc)

        if not isinstance(raw, list):
            return False, f"{JSON_FILE} must contain a JSON array."

        normalized_vhd_path = os.path.normcase(vhd_path)
        normalized_volume_label = volume_label.strip().lower()

        for item in raw:
            if not isinstance(item, dict):
                continue
            existing_path = os.path.normcase(str(item.get("vhd_path", "")).strip())
            existing_label = str(item.get("vhd_volume_label", "")).strip().lower()
            if existing_path == normalized_vhd_path:
                return False, f"An entry for this VHDX path already exists in {JSON_FILE}."
            if existing_label == normalized_volume_label:
                return False, f"An entry for this volume label already exists in {JSON_FILE}."

        return True, "OK"

    def append_entry_to_json(self, title_text: str, vhd_path: str, volume_label: str) -> Tuple[bool, str]:
        try:
            json_path = Path(JSON_FILE)
            with json_path.open("r", encoding="utf-8") as f:
                raw = json.load(f)

            if not isinstance(raw, list):
                return False, f"{JSON_FILE} must contain a JSON array."

            raw.append(
                {
                    "vhd_path": vhd_path,
                    "vhd_volume_label": volume_label,
                    "vhd_description": title_text,
                }
            )

            with json_path.open("w", encoding="utf-8") as f:
                json.dump(raw, f, indent=2)
                f.write("\n")

            return True, "OK"
        except Exception as exc:
            return False, str(exc)

    def on_bullet_click(self, vhd_path: str) -> None:
        if self.is_busy:
            return

        row = self.row_widgets.get(vhd_path.lower())
        if not row:
            return

        current_state = row["state_label"].cget("text")
        if current_state == "Unavailable":
            self.status_var.set("That VHDX file is unavailable on disk.")
            return

        row["canvas"].itemconfigure(row["oval"], fill=COLORS["processing"], outline=COLORS["processing"])
        row["state_label"].configure(text="Processing", fg=COLORS["processing"])

        self.set_busy(True, f"Processing {Path(vhd_path).name}...")
        threading.Thread(target=self._toggle_worker, args=(row["entry"], current_state), daemon=True).start()

    def _toggle_worker(self, entry: VHDEntry, current_state: str) -> None:
        try:
            if current_state == "Mounted":
                self.detach_vhd(entry.vhd_path)
            elif current_state == "Unmounted":
                self.attach_vhd(entry.vhd_path)
            elif current_state == "Unhealthy":
                self.detach_vhd(entry.vhd_path)
            self.after(0, self._refresh_after_toggle)
        except Exception as exc:
            self.after(0, lambda: self.set_busy(False, "Operation failed."))
            self.after(0, lambda: messagebox.showerror(APP_TITLE, str(exc)))
            self.after(0, self.refresh_all)

    def _refresh_after_toggle(self) -> None:
        self.set_busy(False, "Refreshing VHD states...")
        self.refresh_all()

    def run_powershell(self, script: str) -> str:
        proc = subprocess.run(
            POWERSHELL + [script],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            shell=False,
        )
        if proc.returncode != 0:
            stderr = (proc.stderr or "").strip()
            stdout = (proc.stdout or "").strip()
            detail = stderr or stdout or f"PowerShell exited with code {proc.returncode}."
            raise PowerShellError(detail)
        return proc.stdout.strip()

    def get_volume_index(self) -> Dict[str, Dict[str, Any]]:
        script = (
            "$data = Get-Volume | Select-Object FileSystemLabel, HealthStatus, OperationalStatus, DriveLetter, Path; "
            "$data | ConvertTo-Json -Depth 4"
        )
        output = self.run_powershell(script)
        if not output:
            return {}

        parsed = json.loads(output)
        if isinstance(parsed, dict):
            parsed = [parsed]

        index: Dict[str, Dict[str, Any]] = {}
        for item in parsed:
            label = str(item.get("FileSystemLabel") or "").strip()
            if label:
                index[label.lower()] = item
        return index

    def get_disk_image_index(self) -> Dict[str, Dict[str, Any]]:
        script = r"""
$virtualDisks = Get-CimInstance Win32_DiskDrive |
    Where-Object {
        $_.Model -match 'Virtual' -or $_.Caption -match 'Virtual' -or $_.PNPDeviceID -match 'VHD'
    } |
    Select-Object -ExpandProperty DeviceID

$result = foreach ($devicePath in $virtualDisks) {
    try {
        Get-DiskImage -DevicePath $devicePath -ErrorAction Stop |
            Where-Object { $_.ImagePath -like '*.vhd' -or $_.ImagePath -like '*.vhdx' } |
            Select-Object ImagePath, Attached, DevicePath
    }
    catch {
    }
}

$result | ConvertTo-Json -Depth 4
"""
        output = self.run_powershell(script)
        if not output:
            return {}

        parsed = json.loads(output)
        if isinstance(parsed, dict):
            parsed = [parsed]

        index: Dict[str, Dict[str, Any]] = {}
        for item in parsed:
            image_path = str(item.get("ImagePath") or "").strip()
            if image_path:
                index[os.path.normcase(image_path)] = item
        return index

    def compute_state(
        self,
        entry: VHDEntry,
        volume_index: Dict[str, Dict[str, Any]],
        disk_index: Dict[str, Dict[str, Any]],
    ) -> Dict[str, Any]:
        path_exists = Path(entry.vhd_path).exists()
        if not path_exists:
            return {"state": "Unavailable", "detail": "VHDX file not found."}

        volume = volume_index.get(entry.vhd_volume_label.lower())
        disk_image = disk_index.get(os.path.normcase(entry.vhd_path))

        if volume:
            health = str(volume.get("HealthStatus") or "").strip()
            operational = volume.get("OperationalStatus")
            if isinstance(operational, list):
                operational_text = ", ".join(str(x) for x in operational)
            else:
                operational_text = str(operational or "").strip()

            detail = f"HealthStatus={health or 'Unknown'}; OperationalStatus={operational_text or 'Unknown'}"
            is_healthy = health.lower() == "healthy" and operational_text.lower() == "ok"
            return {"state": "Mounted" if is_healthy else "Unhealthy", "detail": detail}

        if disk_image and bool(disk_image.get("Attached")):
            return {
                "state": "Unhealthy",
                "detail": "Disk image is attached, but the expected volume label was not found in Get-Volume.",
            }

        return {"state": "Unmounted", "detail": "VHDX file exists and is not currently mounted."}

    def attach_vhd(self, vhd_path: str) -> None:
        escaped = ps_quote(vhd_path)
        script = f"Mount-DiskImage -ImagePath {escaped} -ErrorAction Stop"
        self.run_powershell(script)

    def detach_vhd(self, vhd_path: str) -> None:
        escaped = ps_quote(vhd_path)
        script = f"Dismount-DiskImage -ImagePath {escaped} -ErrorAction Stop"
        self.run_powershell(script)


def create_dynamic_vhdx_diskpart_safe(
    vhd_path: str,
    volume_label: str,
    max_size_gb: int,
    mount_folder: str,
    allowed_vhd_base: str = r"C:\LOCAL_VHD",
    allowed_mount_base: str = r"C:\DEV\vhd_mounts",
) -> Tuple[bool, str]:
    try:
        vhd = Path(vhd_path).resolve()
        mount = Path(mount_folder).resolve()
        vhd_base = Path(allowed_vhd_base).resolve()
        mount_base = Path(allowed_mount_base).resolve()

        if vhd.suffix.lower() != ".vhdx":
            return False, "The target file must have a .vhdx extension."

        if max_size_gb <= 0:
            return False, "max_size_gb must be greater than 0."

        if max_size_gb > 1024:
            return False, "max_size_gb is too large for this safety policy."

        if not vhd_base.exists():
            return False, f"Allowed VHD base folder does not exist: {vhd_base}"

        if not mount_base.exists():
            return False, f"Allowed mount base folder does not exist: {mount_base}"

        if not vhd.parent.exists():
            return False, f"Parent folder does not exist: {vhd.parent}"

        if not mount.exists():
            return False, f"Mount folder does not exist: {mount}"

        try:
            vhd.relative_to(vhd_base)
        except ValueError:
            return False, f"VHD path must be inside {vhd_base}"

        try:
            mount.relative_to(mount_base)
        except ValueError:
            return False, f"Mount folder must be inside {mount_base}"

        if vhd.exists():
            return False, f"Refusing to overwrite existing file: {vhd}"

        if any(mount.iterdir()):
            return False, f"Mount folder must be empty: {mount}"

        safe_label = volume_label.strip()
        if not safe_label:
            return False, "volume_label must not be empty."

        if len(safe_label) > 32:
            return False, "volume_label must be 32 characters or fewer."

        if '"' in safe_label:
            return False, 'volume_label must not contain a double quote (").'

        if not re.fullmatch(r"[A-Za-z0-9 _.-]+", safe_label):
            return False, "volume_label contains unsupported characters."

        max_size_mb = max_size_gb * 1024

        create_script = "\n".join(
            [
                f'create vdisk file="{vhd}" maximum={max_size_mb} type=expandable',
                f'select vdisk file="{vhd}"',
                'attach vdisk',
                'create partition primary',
                f'format fs=ntfs label="{safe_label}" quick',
                f'assign mount="{mount}"',
            ]
        )

        with tempfile.NamedTemporaryFile(
            mode="w",
            suffix=".txt",
            delete=False,
            encoding="utf-8",
            newline="\r\n",
        ) as f:
            script_path = Path(f.name)
            f.write(create_script)
            f.write("\n")

        try:
            result = subprocess.run(
                ["diskpart", "/s", str(script_path)],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                check=False,
            )
        finally:
            script_path.unlink(missing_ok=True)

        output = ((result.stdout or "") + "\n" + (result.stderr or "")).strip()

        if result.returncode != 0:
            cleanup_script = "\n".join(
                [
                    f'select vdisk file="{vhd}"',
                    'detach vdisk',
                ]
            )
            with tempfile.NamedTemporaryFile(
                mode="w",
                suffix=".txt",
                delete=False,
                encoding="utf-8",
                newline="\r\n",
            ) as f:
                cleanup_path = Path(f.name)
                f.write(cleanup_script)
                f.write("\n")
            try:
                subprocess.run(
                    ["diskpart", "/s", str(cleanup_path)],
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    check=False,
                )
            finally:
                cleanup_path.unlink(missing_ok=True)

            return False, output or "DiskPart failed."

        if not vhd.exists():
            return False, "DiskPart reported success, but the VHDX file was not found afterward."

        return True, f"Created VHDX successfully: {vhd}\nMounted at: {mount}"

    except Exception as exc:
        return False, str(exc)


def ps_quote(value: str) -> str:
    return "'" + value.replace("'", "''") + "'"


def ensure_admin() -> None:
    if not ctypes.windll.shell32.IsUserAnAdmin():
        argv = subprocess.list2cmdline(sys.argv)
        ctypes.windll.shell32.ShellExecuteW(
            None,
            "runas",
            sys.executable,
            argv,
            None,
            1,
        )
        sys.exit()


def main() -> None:
    ensure_admin()
    app = VHDManagerApp()
    app.mainloop()


if __name__ == "__main__":
    main()
