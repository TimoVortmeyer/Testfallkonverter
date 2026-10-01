@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    where py.exe >nul 2>&1
    if not errorlevel 1 (
        py -3 -m venv .venv
    ) else (
        where python.exe >nul 2>&1
        if errorlevel 1 (
            echo Python 3.11 oder neuer wurde nicht gefunden.
            echo Installiere Python und starte diesen Launcher erneut.
            exit /b 1
        )
        python -m venv .venv
    )
    if errorlevel 1 (
        echo Die virtuelle Umgebung konnte nicht erstellt werden.
        exit /b 1
    )
)

".venv\Scripts\python.exe" -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 11) else 1)"
if errorlevel 1 (
    echo Python 3.11 oder neuer wird benoetigt. Loesche .venv nach einem Python-Upgrade und starte erneut.
    exit /b 1
)

if not exist ".venv\.requirements-installed" (
    echo Installiere Abhaengigkeiten fuer den Testfallkonverter. Beim ersten Start ist Internetzugriff erforderlich.
    ".venv\Scripts\python.exe" -m pip install --disable-pip-version-check -r requirements.txt
    if errorlevel 1 (
        echo Die Abhaengigkeiten konnten nicht installiert werden.
        exit /b 1
    )
    > ".venv\.requirements-installed" echo installed
)

set "PYTHONPATH=%CD%\src"
if "%~1"=="" goto show_help
".venv\Scripts\python.exe" -m lunar_converter %*
exit /b %ERRORLEVEL%

:show_help
".venv\Scripts\python.exe" -m lunar_converter --help
exit /b %ERRORLEVEL%