@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0"

echo ============================================================
echo  Building MXL Column Editor for Windows
echo ============================================================
echo.

set PY=py
where py >nul 2>nul || set PY=python
%PY% --version >nul 2>nul
if errorlevel 1 (
    echo ERROR: Python was not found on this machine.
    echo        Install it from https://www.python.org/downloads/windows/
    echo        and tick "Add python.exe to PATH".
    pause
    exit /b 1
)
for /f "delims=" %%v in ('%PY% --version') do echo Using %%v

echo.
echo Installing PyInstaller...
%PY% -m pip install --upgrade --quiet pyinstaller
if errorlevel 1 goto :fail

set EXTRA=
%PY% -c "import tkinterweb" >nul 2>nul
if errorlevel 1 (
    echo.
    echo NOTE: tkinterweb is not installed, so the built .exe will open 1C
    echo       previews in your browser instead of embedding them.
    echo       To embed them, cancel now, run:
    echo           %PY% -m pip install tkinterweb
    echo       and start this script again.
    echo.
    timeout /t 5 >nul
) else (
    echo tkinterweb found - bundling the embedded HTML view.
    set EXTRA=--collect-all tkinterweb --collect-all tkinterweb_tkhtml
)

echo.
echo Building...
%PY% -m PyInstaller --noconfirm --clean --windowed --onefile ^
    --name "MXL Column Editor" ^
    --add-data "onec;onec" ^
    !EXTRA! ^
    mxl_column_editor.py
if errorlevel 1 goto :fail

rem A copy of onec\ beside the .exe lets anyone drop in a rebuilt .epf
rem without rebuilding the whole program - the editor prefers it over the
rem copy baked inside.
xcopy /E /I /Y onec "dist\onec" >nul

echo.
echo ============================================================
echo  Done.
echo.
echo  Hand over the whole  dist\  folder:
echo      dist\MXL Column Editor.exe
echo      dist\onec\           (rebuildable .epf files)
echo.
echo  The .exe alone also works - it carries its own copy of onec.
echo ============================================================
echo.
pause
exit /b 0

:fail
echo.
echo BUILD FAILED - see the messages above.
pause
exit /b 1
