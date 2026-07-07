@echo off
setlocal EnableExtensions
cd /d "%~dp0"
if errorlevel 1 (
    echo [ERROR] Cannot enter script folder: %~dp0
    pause
    exit /b 1
)

title XiaoLinAssistant Run
set "LOG_FILE=%~dp0run_log.txt"
set "PYTHON_CMD="

> "%LOG_FILE%" echo XiaoLinAssistant run log
>> "%LOG_FILE%" echo Current folder: %cd%
>> "%LOG_FILE%" echo Time: %date% %time%

if not exist "xiaoxin_assistant.py" (
    echo [ERROR] xiaoxin_assistant.py not found.
    echo Please unzip the whole package first.
    >> "%LOG_FILE%" echo xiaoxin_assistant.py not found
    pause
    exit /b 10
)

call :FIND_PYTHON
if errorlevel 1 (
    echo [ERROR] Python was not found.
    pause
    exit /b 20
)

echo Python command: %PYTHON_CMD%
call %PYTHON_CMD% -m pip install customtkinter pillow pynput pydivert python-docx openpyxl PyPDF2 PyYAML pystray >> "%LOG_FILE%" 2>&1
call %PYTHON_CMD% xiaoxin_assistant.py >> "%LOG_FILE%" 2>&1
if errorlevel 1 (
    echo [ERROR] Program exited with an error. Send run_log.txt to me.
    pause
    exit /b 30
)
exit /b 0

:FIND_PYTHON
py -3.12 --version >nul 2>&1 && set "PYTHON_CMD=py -3.12" && exit /b 0
py -3.11 --version >nul 2>&1 && set "PYTHON_CMD=py -3.11" && exit /b 0
py -3.10 --version >nul 2>&1 && set "PYTHON_CMD=py -3.10" && exit /b 0
py -3.13 --version >nul 2>&1 && set "PYTHON_CMD=py -3.13" && exit /b 0
py -3 --version >nul 2>&1 && set "PYTHON_CMD=py -3" && exit /b 0
py --version >nul 2>&1 && set "PYTHON_CMD=py" && exit /b 0
python --version >nul 2>&1 && set "PYTHON_CMD=python" && exit /b 0
python3 --version >nul 2>&1 && set "PYTHON_CMD=python3" && exit /b 0
exit /b 1
