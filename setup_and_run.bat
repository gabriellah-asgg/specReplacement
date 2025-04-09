@echo off
setlocal

REM --- Use full path to Python ---
set "PYTHON_PATH=%USERPROFILE%\AppData\Local\Programs\Python\Python312\python.exe"

REM Check if Python is available
if not exist %PYTHON_PATH% (
    echo ERROR: Python is not installed at %PYTHON_PATH%.
    pause
    exit /b
)

REM Check if venv folder already exists
if not exist "venv\" (
    echo Creating virtual environment...
    %PYTHON_PATH% -m venv venv
    call venv\Scripts\activate.bat
    echo Installing dependencies...
    pip install --upgrade pip
    pip install -r requirements.txt
) else (
    echo Virtual environment already exists.
    call venv\Scripts\activate.bat
)

REM Run your script
echo Running script...
python specification_replacement.py

pause