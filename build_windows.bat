@echo off
setlocal
cd /d %~dp0

if not exist .venv (
    py -3.7 -m venv .venv
    if errorlevel 1 goto :error
)

call .venv\Scripts\activate.bat
if errorlevel 1 goto :error

python -c "import sys; print(sys.version)"
python -c "import tkinter as tk; print('tkinter ok', tk.TkVersion)"
if errorlevel 1 goto :tkerror

python -m pip install --upgrade pip
if errorlevel 1 echo Warning: pip upgrade failed, continuing...

if exist wheelhouse (
    echo Installing dependencies from local wheelhouse...
    python -m pip install --no-index --find-links=wheelhouse openpyxl==3.0.10 pyinstaller==5.13.2
    if errorlevel 1 goto :error
) else (
    echo Installing dependencies from configured package index...
    python -m pip install -r requirements.txt
    if errorlevel 1 goto :error
)

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

python -m PyInstaller --noconfirm --clean --onedir --windowed --name OfflineDialogueLabelerPro ^
    --paths src ^
    --hidden-import=tkinter ^
    --hidden-import=tkinter.ttk ^
    --add-data "config;config" ^
    --add-data "input;input" ^
    run_app.py
if errorlevel 1 goto :error

echo.
echo Build finished successfully.
echo Open dist\OfflineDialogueLabelerPro\OfflineDialogueLabelerPro.exe
pause
exit /b 0

:tkerror
echo.
echo ERROR: tkinter is not available in this Python installation.
pause
exit /b 1

:error
echo.
echo Build failed.
echo dist folder was not created because one of the steps above failed.
pause
exit /b 1
