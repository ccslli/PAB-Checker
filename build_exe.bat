@echo off
rem Builds "PAB Checker.exe" (the window version) into the dist folder.
rem Needs:  pip install pyinstaller pyside6 pywin32 xlwings
rem --add-data packs the branding folder (icon) inside the .exe
rem --icon     is the icon shown on the .exe file itself
cd /d "%~dp0"
pyinstaller --noconfirm --onefile --windowed --name "PAB Checker" ^
  --icon "branding\dt_icon.ico" ^
  --add-data "branding;branding" ^
  pab_checker_gui.py
echo.
echo Done: dist\PAB Checker.exe
pause
