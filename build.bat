@echo off
cd /d %~dp0

echo ==============================
echo Building vhdx_manager.exe
echo ==============================

REM Clean previous builds
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
del /q *.spec 2>nul

REM Build executable
pyinstaller ^
  --onefile ^
  --windowed ^
  --icon vhdx_manager.ico ^
  --manifest admin.manifest ^
  --name vhdx_manager ^
  vhdx_manager.py

echo.
echo Copying resource files...

copy vhdx_manager_icon.png dist\
copy vhdx_list.json dist\

echo.
echo ==============================
echo Build complete
echo ==============================
echo Executable location:
echo   dist\vhdx_manager.exe
echo.

pause