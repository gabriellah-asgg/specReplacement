@echo off
setlocal

REM Check for virtual environment
if not exist "venv" (
    echo Creating virtual environment...
    python -m venv venv

    echo Installing dependencies...
    call venv\Scripts\activate.bat
    pip install --upgrade pip
    pip install -r requirements.txt
) else (
    echo Using existing virtual environment...
    call venv\Scripts\activate.bat
)

REM Run your script
echo Running script...
python specification_replacement.py

pause