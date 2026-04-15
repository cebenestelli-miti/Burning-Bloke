@echo off
setlocal
cd /d "%~dp0"

REM One-file exe: each launch extracts the bundle to %%TEMP%% first — that alone is often 5–20+ s.
REM For much faster cold starts use build-onedir.bat (folder + BBSchedule.exe, no extraction).

python -m pip install --upgrade pip
python -m pip install -r requirements.txt

REM Build in %%TEMP%% so Google Drive / locked dist\BBSchedule.exe does not break PyInstaller
set "PYI_DIST=%TEMP%\BBSchedule_pyinstaller_dist"
set "PYI_WORK=%TEMP%\BBSchedule_pyinstaller_work"
if exist "%PYI_DIST%" rmdir /s /q "%PYI_DIST%"
if exist "%PYI_WORK%" rmdir /s /q "%PYI_WORK%"
if exist build rmdir /s /q build
if exist BBSchedule.spec del /q BBSchedule.spec

python -m PyInstaller --noconfirm --onefile --windowed --name BBSchedule ^
  --icon "app_icon.ico" ^
  --add-data "app_icon.ico;." ^
  --distpath "%PYI_DIST%" ^
  --workpath "%PYI_WORK%" ^
  --hidden-import=tkcalendar ^
  --hidden-import=babel.numbers ^
  --hidden-import=babel.dates ^
  --collect-all tkcalendar ^
  --collect-all babel ^
  --collect-all customtkinter ^
  main.py

if errorlevel 1 (
  echo PyInstaller failed.
  exit /b 1
)

if not exist dist mkdir dist

REM Stage under a new name first, then replace dist\BBSchedule.exe.
REM In-place copy often fails or looks unchanged with Google Drive / read-only / shell locks.
set "DST=dist\BBSchedule.exe"
set "STG=dist\BBSchedule.exe.new"

del /F /Q "%STG%" 2>nul
copy /B /Y "%PYI_DIST%\BBSchedule.exe" "%STG%"
if errorlevel 1 (
  echo.
  echo Failed to stage new exe into dist\
  exit /b 1
)

if exist "%DST%" attrib -R "%DST%" >nul 2>&1
move /Y "%STG%" "%DST%"
if errorlevel 1 (
  echo.
  echo Could not replace %DST% — close BBSchedule.exe and any Explorer windows on dist\, then retry.
  echo Or pause Google Drive sync for this folder.
  echo Fresh exe is still at:
  echo   %PYI_DIST%\BBSchedule.exe
  echo   %STG%
  exit /b 1
)

echo.
for %%F in ("%DST%") do echo Installed: %%~fF  ^| %%~zF bytes ^| %%~tF
echo Build complete.
endlocal
