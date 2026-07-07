@echo off

REM iOS tools installer path note:
REM setup.iss is inside installer\, so Source must use ..\tools\ios\*
REM Correct line:
REM Source: "..\tools\ios\*"; DestDir: "{app}\tools\ios"; Flags: ignoreversion recursesubdirs createallsubdirs


REM ============================================================
REM iOS tools packaging check
REM setup.iss must include:
REM Source: "..\tools\ios\*"; DestDir: "{app}\tools\ios"; Flags: ignoreversion recursesubdirs createallsubdirs
REM The tools\ios folder should contain idevice_id.exe / ideviceinfo.exe / idevicesyslog.exe / idevicecrashreport.exe and required dll files.
REM ============================================================
IF NOT EXIST "%~dp0tools\ios" (
    ECHO [WARN] tools\ios folder not found. iOS log tools will not be packaged into installer.
) ELSE (
    ECHO [OK] Found tools\ios folder. It will be packaged by installer/setup.iss.
)

setlocal EnableExtensions DisableDelayedExpansion
chcp 65001 >nul 2>nul

rem ============================================================
rem QA Test Toolbox Installer Builder - v6 no-log/output fix
rem Fix:
rem - create log immediately at startup
rem - support paths with Chinese characters and spaces
rem - never rely on current command prompt folder
rem - force English safe setup file first
rem - use one ASCII-only setup filename for reliable Windows batch parsing
rem ============================================================

set "ROOT=%~dp0"
set "INSTALLER_LOG=%ROOT%build_installer_log.txt"

> "%INSTALLER_LOG%" echo QA Test Toolbox installer build log
if errorlevel 1 (
    set "INSTALLER_LOG=%TEMP%\xiaolin_build_installer_log.txt"
    > "%INSTALLER_LOG%" echo QA Test Toolbox installer build log
)

>> "%INSTALLER_LOG%" echo Script: %~f0
>> "%INSTALLER_LOG%" echo Root: %ROOT%
>> "%INSTALLER_LOG%" echo Time: %date% %time%
>> "%INSTALLER_LOG%" echo.

cd /d "%ROOT%" >> "%INSTALLER_LOG%" 2>&1
if errorlevel 1 (
    echo [ERROR] Cannot enter script folder: %ROOT%
    echo Log file: %INSTALLER_LOG%
    >> "%INSTALLER_LOG%" echo [ERROR] Cannot enter script folder: %ROOT%
    pause
    exit /b 1
)

cls
title QA Test Toolbox Installer Builder v6

set "APP_VERSION="
for /f "usebackq delims=" %%V in (`powershell.exe -NoProfile -Command "$v=(Get-Content -LiteralPath '.\version.json' -Raw | ConvertFrom-Json).version; Write-Output $v"`) do set "APP_VERSION=%%V"
if not defined APP_VERSION (
    echo [ERROR] Cannot read version from version.json.
    >> "%INSTALLER_LOG%" echo [ERROR] Cannot read version from version.json
    pause
    exit /b 2
)

echo ============================================
echo QA Test Toolbox v%APP_VERSION% Installer Builder v6
echo ============================================
echo Current folder: %cd%
echo Log file: %INSTALLER_LOG%
echo.

set "APP_EXE=%ROOT%dist\XiaoLinAssistant.exe"
set "APP_EXE_REL=dist\XiaoLinAssistant.exe"
set "OUTPUT_DIR=%ROOT%Output"
set "SETUP_BASE_SAFE=XiaoLinAssistant_Setup_v%APP_VERSION%"
set "SETUP_EXE_SAFE=%OUTPUT_DIR%\XiaoLinAssistant_Setup_v%APP_VERSION%.exe"
set "ISCC_EXE="

>> "%INSTALLER_LOG%" echo Current folder: %cd%
>> "%INSTALLER_LOG%" echo APP_EXE=%APP_EXE%
>> "%INSTALLER_LOG%" echo OUTPUT_DIR=%OUTPUT_DIR%
>> "%INSTALLER_LOG%" echo SETUP_EXE_SAFE=%SETUP_EXE_SAFE%
>> "%INSTALLER_LOG%" echo.

if not exist "xiaoxin_assistant.py" (
    echo [ERROR] xiaoxin_assistant.py not found.
    echo Please unzip the whole package first. Do not run this bat inside the zip preview window.
    echo Log file: %INSTALLER_LOG%
    >> "%INSTALLER_LOG%" echo [ERROR] xiaoxin_assistant.py not found
    pause
    exit /b 10
)

if not exist "build_exe.bat" (
    echo [ERROR] build_exe.bat not found.
    echo Log file: %INSTALLER_LOG%
    >> "%INSTALLER_LOG%" echo [ERROR] build_exe.bat not found
    pause
    exit /b 11
)

if not exist "installer\setup.iss" (
    echo [ERROR] installer\setup.iss not found.
    echo Log file: %INSTALLER_LOG%
    >> "%INSTALLER_LOG%" echo [ERROR] installer\setup.iss not found
    pause
    exit /b 12
)

if not exist "assets\app.ico" (
    echo [ERROR] assets\app.ico not found.
    echo Log file: %INSTALLER_LOG%
    >> "%INSTALLER_LOG%" echo [ERROR] assets\app.ico not found
    pause
    exit /b 13
)

if not exist "assets\app.png" (
    echo [ERROR] assets\app.png not found.
    echo Log file: %INSTALLER_LOG%
    >> "%INSTALLER_LOG%" echo [ERROR] assets\app.png not found
    pause
    exit /b 14
)

echo [1/5] Building XiaoLinAssistant.exe...
>> "%INSTALLER_LOG%" echo [1/5] Calling build_exe.bat

set "XIAOLIN_SKIP_PAUSE=1"
call "%ROOT%build_exe.bat"
set "BUILD_EXE_ERR=%ERRORLEVEL%"
set "XIAOLIN_SKIP_PAUSE="

rem build_exe.bat may change caller variables; reset them here.
set "ROOT=%~dp0"
set "INSTALLER_LOG=%ROOT%build_installer_log.txt"
if not exist "%INSTALLER_LOG%" set "INSTALLER_LOG=%TEMP%\xiaolin_build_installer_log.txt"
set "APP_EXE=%ROOT%dist\XiaoLinAssistant.exe"
set "APP_EXE_REL=dist\XiaoLinAssistant.exe"
set "UPDATER_EXE=%ROOT%dist\XiaoLinUpdater.exe"
set "OUTPUT_DIR=%ROOT%Output"
set "SETUP_BASE_SAFE=XiaoLinAssistant_Setup_v%APP_VERSION%"
set "SETUP_EXE_SAFE=%OUTPUT_DIR%\XiaoLinAssistant_Setup_v%APP_VERSION%.exe"
set "ISCC_EXE="

>> "%INSTALLER_LOG%" echo.
>> "%INSTALLER_LOG%" echo Returned from build_exe.bat. ErrorCode=%BUILD_EXE_ERR%
>> "%INSTALLER_LOG%" echo APP_EXE=%APP_EXE%

if not "%BUILD_EXE_ERR%"=="0" (
    echo [ERROR] build_exe.bat failed. ErrorCode=%BUILD_EXE_ERR%
    echo Check build_log.txt and build_installer_log.txt.
    >> "%INSTALLER_LOG%" echo [ERROR] build_exe.bat failed. ErrorCode=%BUILD_EXE_ERR%
    pause
    exit /b 20
)

if not exist "%APP_EXE%" (
    echo [ERROR] %APP_EXE_REL% not found after build.
    echo Actual dist folder content:
    if exist "%ROOT%dist" dir "%ROOT%dist" /b
    if not exist "%ROOT%dist" echo dist folder not found
    echo Log file: %INSTALLER_LOG%
    >> "%INSTALLER_LOG%" echo [ERROR] %APP_EXE_REL% not found after build
    if exist "%ROOT%dist" dir "%ROOT%dist" /b >> "%INSTALLER_LOG%" 2>&1
    if not exist "%ROOT%dist" >> "%INSTALLER_LOG%" echo dist folder not found
    pause
    exit /b 21
)
if not exist "%UPDATER_EXE%" (
    echo [ERROR] dist\XiaoLinUpdater.exe not found after build.
    >> "%INSTALLER_LOG%" echo [ERROR] updater exe not found after build
    pause
    exit /b 22
)

echo.
echo [2/5] Finding Inno Setup Compiler ISCC.exe...
>> "%INSTALLER_LOG%" echo [2/5] Finding ISCC.exe
call :FIND_ISCC
if errorlevel 1 goto FAIL_ISCC

echo ISCC path: %ISCC_EXE%
>> "%INSTALLER_LOG%" echo ISCC path: %ISCC_EXE%

echo.
echo [3/5] Cleaning Output folder...
>> "%INSTALLER_LOG%" echo [3/5] Cleaning Output folder
if not exist "%OUTPUT_DIR%" mkdir "%OUTPUT_DIR%" >> "%INSTALLER_LOG%" 2>&1
del /f /q "%OUTPUT_DIR%\*.exe" >> "%INSTALLER_LOG%" 2>&1
del /f /q "%OUTPUT_DIR%\*.bin" >> "%INSTALLER_LOG%" 2>&1
del /f /q "%OUTPUT_DIR%\*.tmp" >> "%INSTALLER_LOG%" 2>&1

echo.
echo [4/5] Building installer with Inno Setup...
echo Command:
echo "%ISCC_EXE%" /DMyAppVersion=%APP_VERSION% /O"%OUTPUT_DIR%" /F"%SETUP_BASE_SAFE%" "installer\setup.iss"
>> "%INSTALLER_LOG%" echo [4/5] Running ISCC
>> "%INSTALLER_LOG%" echo Command: "%ISCC_EXE%" /DMyAppVersion=%APP_VERSION% /O"%OUTPUT_DIR%" /F"%SETUP_BASE_SAFE%" "installer\setup.iss"

"%ISCC_EXE%" /DMyAppVersion=%APP_VERSION% /O"%OUTPUT_DIR%" /F"%SETUP_BASE_SAFE%" "installer\setup.iss" >> "%INSTALLER_LOG%" 2>&1
set "ISCC_ERR=%ERRORLEVEL%"
>> "%INSTALLER_LOG%" echo ISCC ErrorCode=%ISCC_ERR%

if not "%ISCC_ERR%"=="0" (
    echo [ERROR] Inno Setup build failed. ErrorCode=%ISCC_ERR%
    echo Check build_installer_log.txt.
    echo.
    echo Output folder content:
    if exist "%OUTPUT_DIR%" dir "%OUTPUT_DIR%" /b
    >> "%INSTALLER_LOG%" echo [ERROR] Inno Setup build failed. ErrorCode=%ISCC_ERR%
    if exist "%OUTPUT_DIR%" dir "%OUTPUT_DIR%" /b >> "%INSTALLER_LOG%" 2>&1
    pause
    exit /b 30
)

echo.
echo [5/5] Normalizing setup file names...
>> "%INSTALLER_LOG%" echo [5/5] Normalizing setup file names

if not exist "%SETUP_EXE_SAFE%" (
    for %%F in ("%OUTPUT_DIR%\*.exe") do (
        if exist "%%~fF" (
            echo Inno generated: %%~fF
            echo Normalizing to: %SETUP_EXE_SAFE%
            >> "%INSTALLER_LOG%" echo Inno generated: %%~fF
            >> "%INSTALLER_LOG%" echo Normalizing to: %SETUP_EXE_SAFE%
            copy /y "%%~fF" "%SETUP_EXE_SAFE%" >> "%INSTALLER_LOG%" 2>&1
            goto CHECK_SAFE_EXE
        )
    )
)

:CHECK_SAFE_EXE
if not exist "%SETUP_EXE_SAFE%" (
    echo [ERROR] Installer exe was not generated.
    echo Expected: %SETUP_EXE_SAFE%
    echo Actual Output folder content:
    if exist "%OUTPUT_DIR%" dir "%OUTPUT_DIR%" /b
    if not exist "%OUTPUT_DIR%" echo Output folder not found
    echo Log file: %INSTALLER_LOG%
    >> "%INSTALLER_LOG%" echo [ERROR] Installer exe was not generated
    if exist "%OUTPUT_DIR%" dir "%OUTPUT_DIR%" /b >> "%INSTALLER_LOG%" 2>&1
    if not exist "%OUTPUT_DIR%" >> "%INSTALLER_LOG%" echo Output folder not found
    pause
    exit /b 31
)

for %%F in ("%SETUP_EXE_SAFE%") do set "SAFE_SIZE=%%~zF"
if "%SAFE_SIZE%"=="0" (
    echo [ERROR] English setup exe size is 0. Build output is incomplete.
    >> "%INSTALLER_LOG%" echo [ERROR] English setup exe size is 0
    pause
    exit /b 32
)

echo.
echo Output folder content:
dir "%OUTPUT_DIR%" /b
>> "%INSTALLER_LOG%" echo Output folder content:
dir "%OUTPUT_DIR%" /b >> "%INSTALLER_LOG%" 2>&1

echo.
echo ============================================
echo Installer build completed.
echo English installer path: %SETUP_EXE_SAFE%
echo Log file: %INSTALLER_LOG%
echo ============================================
>> "%INSTALLER_LOG%" echo Installer build completed
>> "%INSTALLER_LOG%" echo English Output: %SETUP_EXE_SAFE%
pause
exit /b 0

:FIND_ISCC
where ISCC.exe >nul 2>&1
if not errorlevel 1 (
    for /f "delims=" %%I in ('where ISCC.exe 2^>nul') do (
        if exist "%%I" (
            set "ISCC_EXE=%%I"
            exit /b 0
        )
    )
)

if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" (
    set "ISCC_EXE=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
    exit /b 0
)

if exist "%ProgramFiles%\Inno Setup 6\ISCC.exe" (
    set "ISCC_EXE=%ProgramFiles%\Inno Setup 6\ISCC.exe"
    exit /b 0
)

if exist "%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe" (
    set "ISCC_EXE=%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe"
    exit /b 0
)

if exist "D:\Program Files (x86)\Inno Setup 6\ISCC.exe" (
    set "ISCC_EXE=D:\Program Files (x86)\Inno Setup 6\ISCC.exe"
    exit /b 0
)

if exist "D:\Program Files\Inno Setup 6\ISCC.exe" (
    set "ISCC_EXE=D:\Program Files\Inno Setup 6\ISCC.exe"
    exit /b 0
)

exit /b 1

:FAIL_ISCC
echo [ERROR] Inno Setup Compiler ISCC.exe was not found.
echo Common path: C:\Program Files (x86)\Inno Setup 6\ISCC.exe
echo If it is installed elsewhere, edit FIND_ISCC in this bat and add the real path.
echo Log file: %INSTALLER_LOG%
>> "%INSTALLER_LOG%" echo [ERROR] ISCC.exe was not found
pause
exit /b 22
