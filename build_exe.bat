@echo off
setlocal EnableExtensions

rem Always switch to the folder where this bat file is located.
cd /d "%~dp0"
if errorlevel 1 (
    echo [ERROR] Cannot enter script folder: %~dp0
    call :MAYBE_PAUSE
    exit /b 1
)

title XiaoLinAssistant EXE Builder

set "LOG_FILE=%~dp0build_log.txt"
set "VENV_DIR=%~dp0.venv_build"
set "APP_NAME=XiaoLinAssistant"
set "UPDATER_NAME=XiaoLinUpdater"
set "PYTHON_CMD="
set "PYTHON_EXE="
set "ADB_OPT="

cls
echo ============================================
echo XiaoLinAssistant EXE Builder
echo ============================================
echo Current folder: %cd%
echo Log file: %LOG_FILE%
echo.
echo This window will stay open after build.
echo.

> "%LOG_FILE%" echo XiaoLinAssistant build log
>> "%LOG_FILE%" echo Current folder: %cd%
>> "%LOG_FILE%" echo Time: %date% %time%
>> "%LOG_FILE%" echo.

if not exist "xiaoxin_assistant.py" (
    echo [ERROR] xiaoxin_assistant.py not found.
    echo Please unzip the whole package first. Do not run this bat inside the zip preview window.
    >> "%LOG_FILE%" echo xiaoxin_assistant.py not found
    call :MAYBE_PAUSE
    exit /b 10
)

if not exist "assets\app.ico" (
    echo [ERROR] assets\app.ico not found.
    >> "%LOG_FILE%" echo assets app.ico not found
    call :MAYBE_PAUSE
    exit /b 11
)
if not exist "assets\app.png" (
    echo [ERROR] assets\app.png not found.
    >> "%LOG_FILE%" echo assets app.png not found
    call :MAYBE_PAUSE
    exit /b 12
)
if not exist "version.json" (
    echo [ERROR] version.json not found.
    >> "%LOG_FILE%" echo version.json not found
    call :MAYBE_PAUSE
    exit /b 13
)
if not exist "updater_main.py" (
    echo [ERROR] updater_main.py not found.
    >> "%LOG_FILE%" echo updater_main.py not found
    call :MAYBE_PAUSE
    exit /b 14
)

if exist "adb.exe" (
    set "ADB_OPT=--add-binary=adb.exe;."
    echo [INFO] Local adb.exe found. It will be bundled.
    >> "%LOG_FILE%" echo Local adb.exe found
) else (
    echo [INFO] Local adb.exe not found. The app will use adb from system PATH.
    >> "%LOG_FILE%" echo Local adb.exe not found
)

echo.
echo [1/6] Finding Python...
>> "%LOG_FILE%" echo [1/6] Finding Python
call :FIND_PYTHON
if errorlevel 1 goto FAIL_PYTHON

echo Python command: %PYTHON_CMD%
>> "%LOG_FILE%" echo Python command: %PYTHON_CMD%
call %PYTHON_CMD% --version
call %PYTHON_CMD% --version >> "%LOG_FILE%" 2>&1
if errorlevel 1 (
    echo [ERROR] Python command cannot run.
    >> "%LOG_FILE%" echo Python command cannot run
    call :MAYBE_PAUSE
    exit /b 21
)

echo.
echo [2/6] Creating build virtual environment...
>> "%LOG_FILE%" echo [2/6] Creating build virtual environment
if not exist "%VENV_DIR%\Scripts\python.exe" (
    call %PYTHON_CMD% -m venv "%VENV_DIR%" >> "%LOG_FILE%" 2>&1
    if errorlevel 1 (
        echo [ERROR] Failed to create virtual environment.
        echo Send build_log.txt to me.
        call :MAYBE_PAUSE
        exit /b 30
    )
)

set "PYTHON_EXE=%VENV_DIR%\Scripts\python.exe"
if not exist "%PYTHON_EXE%" (
    echo [ERROR] venv python not found: %PYTHON_EXE%
    >> "%LOG_FILE%" echo venv python not found
    call :MAYBE_PAUSE
    exit /b 31
)

echo.
echo [3/6] Installing dependencies. First run may take several minutes...
>> "%LOG_FILE%" echo [3/6] Installing dependencies
"%PYTHON_EXE%" -m pip install --upgrade pip setuptools wheel >> "%LOG_FILE%" 2>&1
if errorlevel 1 (
    echo [ERROR] Failed to update pip tools.
    echo Send build_log.txt to me.
    call :MAYBE_PAUSE
    exit /b 40
)

"%PYTHON_EXE%" -m pip install customtkinter pillow pynput pydivert pyinstaller python-docx openpyxl PyPDF2 PyYAML pystray >> "%LOG_FILE%" 2>&1
if errorlevel 1 (
    echo [ERROR] Failed to install dependencies.
    echo Send build_log.txt to me.
    call :MAYBE_PAUSE
    exit /b 41
)

echo.
echo [4/6] Cleaning old build files...
>> "%LOG_FILE%" echo [4/6] Cleaning old build files
if exist "build" rmdir /s /q "build" >> "%LOG_FILE%" 2>&1
if exist "dist" rmdir /s /q "dist" >> "%LOG_FILE%" 2>&1
if exist "%APP_NAME%.spec" del /f /q "%APP_NAME%.spec" >> "%LOG_FILE%" 2>&1
if exist "%UPDATER_NAME%.spec" del /f /q "%UPDATER_NAME%.spec" >> "%LOG_FILE%" 2>&1

echo.
echo [5/6] Running PyInstaller...
>> "%LOG_FILE%" echo [5/6] Running PyInstaller
"%PYTHON_EXE%" -m PyInstaller --noconfirm --noconsole --onefile --clean --name "%APP_NAME%" --icon "assets\app.ico" --add-data "assets;assets" --add-data "tools;tools" --add-data "version.json;." %ADB_OPT% --collect-all customtkinter --collect-binaries pydivert --collect-all pystray --hidden-import pystray --hidden-import pystray._win32 --hidden-import pydivert --hidden-import pynput.keyboard._win32 --hidden-import pynput.mouse._win32 --hidden-import docx --hidden-import openpyxl --hidden-import PyPDF2 --hidden-import yaml --hidden-import app.adb_tools --hidden-import app.apk_info --hidden-import app.log_monitor --hidden-import app.crash_analyzer --hidden-import app.localization_checker --hidden-import app.localization_sheet_parser --hidden-import app.config_validator --hidden-import app.config_rule_templates --hidden-import app.config_quick_rules --hidden-import app.admin_utils --hidden-import app.ios_log_tools --hidden-import app.time_tools --hidden-import app.value_config_compare --hidden-import app.update_manager "xiaoxin_assistant.py" >> "%LOG_FILE%" 2>&1
if errorlevel 1 (
    echo [ERROR] PyInstaller build failed.
    echo Send build_log.txt to me.
    call :MAYBE_PAUSE
    exit /b 50
)
"%PYTHON_EXE%" -m PyInstaller --noconfirm --noconsole --onefile --clean --name "%UPDATER_NAME%" --icon "assets\app.ico" "updater_main.py" >> "%LOG_FILE%" 2>&1
if errorlevel 1 (
    echo [ERROR] Updater build failed.
    echo Send build_log.txt to me.
    call :MAYBE_PAUSE
    exit /b 51
)

echo.
echo [6/6] Checking output...
>> "%LOG_FILE%" echo [6/6] Checking output
if not exist "dist\%APP_NAME%.exe" (
    echo [ERROR] dist\%APP_NAME%.exe not found.
    >> "%LOG_FILE%" echo output exe not found
    call :MAYBE_PAUSE
    exit /b 60
)
if not exist "dist\%UPDATER_NAME%.exe" (
    echo [ERROR] dist\%UPDATER_NAME%.exe not found.
    >> "%LOG_FILE%" echo updater output not found
    call :MAYBE_PAUSE
    exit /b 61
)

echo.
echo ============================================
echo Build completed.
echo EXE path: %cd%\dist\%APP_NAME%.exe
echo Updater path: %cd%\dist\%UPDATER_NAME%.exe
echo ============================================
>> "%LOG_FILE%" echo Build completed
>> "%LOG_FILE%" echo Output: %cd%\dist\%APP_NAME%.exe
>> "%LOG_FILE%" echo Updater: %cd%\dist\%UPDATER_NAME%.exe
call :MAYBE_PAUSE
exit /b 0

:MAYBE_PAUSE
if /I "%XIAOLIN_SKIP_PAUSE%"=="1" exit /b 0
pause
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

:FAIL_PYTHON
echo [ERROR] Python was not found.
echo Install Python 3.10 / 3.11 / 3.12 / 3.13 and enable Add Python to PATH.
>> "%LOG_FILE%" echo Python was not found
call :MAYBE_PAUSE
exit /b 20
