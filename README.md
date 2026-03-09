# VHDX Manager

This Python application for Microsoft Windows can be used to easily attach/detach (i.e. mount/unmount) Virtual Hard Disk (.vhdx) files.

All the .vhdx files need to be specified in the vhdx_list.json file.

To use, simply click the colored bullet alongside the desired VHD, to toggle between mounted and unmounted states. It may take a few seconds to perform the action.

# Usage

From PowerShell, go to the vhdx_manager folder, and type:
```
python .\vhdx_manager.py
```
You may be requested to elevate to administrator privileges.

Alternatively, create a Windows executable by typing the following in a PowerShell window, from the vhdx_manager folder:
```
pip install pyinstaller
.\build.bat
```
The **.\build.bat** command will create a **dist** folder containing **vhdl_manager.exe**. You can create a Windows desktop shortcut to that file if desired, see the screenshot below. You can copy the shortcut to **C:\ProgramData\Microsoft\Windows\Start Menu\Programs** for it to appear in the Start Menu.

![Shortcut properties](shortcut_properties.png)

When run, a graphical window appears. It may take a few seconds to display the list of VHDs. Click on the colored bullets to attach or detach any VHD.


![Example screenshot](example_screenshot.png)
