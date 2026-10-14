@echo off
rem Builds "PAB Checker.exe" (the window version) into the dist folder.
rem Needs:  py -3.10 -m pip install pyinstaller pyside6 pywin32 xlwings
rem pab_checker.spec holds the build settings (icon, branding folder, and the
rem unused libraries left out to keep the .exe small)
cd /d "%~dp0"
py -3.10 -m PyInstaller --noconfirm --clean pab_checker.spec
echo.
echo Done: dist\PAB Checker.exe
pause
