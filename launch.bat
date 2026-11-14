@echo off
setlocal EnableExtensions DisableDelayedExpansion
cd /d "%~dp0"

if exist "dist\MarkdownToWord\MarkdownToWord.exe" (
    start "" "dist\MarkdownToWord\MarkdownToWord.exe"
    exit /b 0
)

if exist ".venv\Scripts\pythonw.exe" (
    start "" ".venv\Scripts\pythonw.exe" -m markdown_to_word.gui
    exit /b 0
)

if exist ".venv-build\Scripts\pythonw.exe" (
    start "" ".venv-build\Scripts\pythonw.exe" -m markdown_to_word.gui
    exit /b 0
)

echo [ERROR] No executable or Python virtual environment was found.
echo Run build_release.bat or follow the setup instructions in README.md.
pause
exit /b 1
