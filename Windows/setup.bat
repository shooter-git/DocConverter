@echo off

:: Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python is not installed. Please install Python 3.x from https://www.python.org/downloads/ and try again.
    pause
    exit /b
)

:: Create virtual environment if it doesn't exist
if not exist "env" (
    python -m venv env
    echo Virtual environment created.
) else (
    echo Virtual environment already exists.
)

:: Activate virtual environment
call env\Scripts\activate

:: Ensure pip upgrade happens using the virtual environment's python.exe
env\Scripts\python.exe -m pip install --upgrade pip

:: Install necessary packages from requirements.txt
env\Scripts\python.exe -m pip install -r requirements.txt

:: Run the file_converter.py tool
env\Scripts\python.exe file_converter.py

pause
