import ctypes
import json
import os
import subprocess
import sys
import threading
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import messagebox
from typing import Any, Dict, List


APP_TITLE = "VHDX Manager"
JSON_FILE = "vhdx_list.json"
POWERSHELL = ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command"]

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
            text="Click a bullet to mount/unmount a VHDX file.",
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

        self.refresh_button = tk.Button(
            controls,
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
        self.refresh_button.configure(state=("disabled" if busy else "normal"))
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
