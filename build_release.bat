@echo off
setlocal EnableExtensions DisableDelayedExpansion
cd /d "%~dp0"

py -3.13 -c "import sys,platform; raise SystemExit(0 if sys.version_info[:2] == (3,13) and platform.architecture()[0] == '64bit' else 1)"
if errorlevel 1 (
 echo [ERROR] CPython 3.13 x64 is required.
 exit /b 1
)
if not exist ".venv-build\Scripts\python.exe" py -3.13 -m venv .venv-build
if errorlevel 1 exit /b 1
".venv-build\Scripts\python.exe" -m pip install -r requirements-build-win11.lock
if errorlevel 1 exit /b 1
".venv-build\Scripts\python.exe" -m pytest -q
if errorlevel 1 exit /b 1
".venv-build\Scripts\python.exe" -m PyInstaller --noconfirm --clean MarkdownToWord.spec
if errorlevel 1 exit /b 1
if not exist "dist\MarkdownToWord\MarkdownToWord.exe" (
 echo [ERROR] EXE not found.
 exit /b 1
)
if not exist release mkdir release
powershell -NoProfile -Command "Compress-Archive -Path 'dist\MarkdownToWord' -DestinationPath 'release\MarkdownToWord-0.1.0-win-x64.zip' -Force"
if errorlevel 1 exit /b 1
echo [OK] release\MarkdownToWord-0.1.0-win-x64.zip
exit /b 0
