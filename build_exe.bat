@echo off
REM Builds dist\FPA Automation\ from a clean virtual environment.
REM Run from the project root on a Windows PC with Python 3.13 installed.
setlocal
cd /d "%~dp0"

if not exist .venv-build (
    py -3.13 -m venv .venv-build || goto :error
)
call .venv-build\Scripts\activate.bat || goto :error
python -m pip install --upgrade pip || goto :error
pip install -r requirements-dev.txt || goto :error

echo Running tests...
python -m pytest -q || goto :error

echo Building...
pyinstaller fpa_automation.spec --noconfirm --clean || goto :error

if exist settings.json copy /y settings.json "dist\FPA Automation\settings.json" >nul

echo.
echo Done: dist\FPA Automation\FPA Automation.exe
echo Copy the whole "dist\FPA Automation" folder to the shared drive.
exit /b 0

:error
echo.
echo BUILD FAILED - see the messages above.
exit /b 1
