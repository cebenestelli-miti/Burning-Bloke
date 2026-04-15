@echo off

setlocal

cd /d "%~dp0"



REM Onedir build: no one-file extraction on launch — startup is usually much faster than dist\BBSchedule.exe

REM Output: dist\BBSchedule\BBSchedule.exe (keep the whole BBSchedule folder when distributing)



python -m pip install -q -r requirements.txt



set "PYI_DIST=%TEMP%\BBSchedule_pyinstaller_onedir_dist"

set "PYI_WORK=%TEMP%\BBSchedule_pyinstaller_onedir_work"

if exist "%PYI_DIST%" rmdir /s /q "%PYI_DIST%"

if exist "%PYI_WORK%" rmdir /s /q "%PYI_WORK%"

if exist build-onedir rmdir /s /q build-onedir

if exist BBSchedule.spec del /q BBSchedule.spec



python -m PyInstaller --noconfirm --windowed --name BBSchedule ^

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

if exist "dist\BBSchedule" rmdir /s /q "dist\BBSchedule"

xcopy /E /I /Y "%PYI_DIST%\BBSchedule" "dist\BBSchedule" >nul

if errorlevel 1 (

  echo Could not copy to dist\BBSchedule

  exit /b 1

)



echo.

echo Onedir build complete: dist\BBSchedule\BBSchedule.exe

endlocal

