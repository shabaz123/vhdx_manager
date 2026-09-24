@echo off
cd /d %~dp0

echo ==============================
echo Building vhdx_manager.exe
echo ==============================

REM Clean previous build artifacts
if exist build rmdir /s /q build
del /q *.spec 2>nul
if not exist dist mkdir dist

REM Build executable with UAC elevation requested in manifest
pyinstaller ^
  --onefile ^
  --windowed ^
  --uac-admin ^
  --icon vhdx_manager.ico ^
  --manifest admin.manifest ^
  --name vhdx_manager ^
  vhdx_manager.py

echo.
echo Copying resource files...

copy /y vhdx_manager_icon.png dist\
if not exist dist\vhdx_list.json (
  copy /y vhdx_list.json dist\
)

echo.
echo ==============================
echo Build complete
echo ==============================
echo Executable location:
echo   dist\vhdx_manager.exe
echo.

pause