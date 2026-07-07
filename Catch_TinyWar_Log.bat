@echo off
setlocal enabledelayedexpansion
title TinyWar Log Capture

set "PKG_NAME=com.tinywarsurvivalexpress.android"
set "LOG_DIR=App_Logs"

if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"

echo =================================================
echo Package: %PKG_NAME%
echo Please connect the Android device by USB first.
echo Press Ctrl+C to stop log capture.
echo =================================================

:CHECK_DEVICE
adb devices | findstr /r "\<device\>" >nul
if errorlevel 1 (
    echo No Android device detected. Retrying in 3 seconds...
    timeout /t 3 >nul
    goto CHECK_DEVICE
)

set "TIMESTAMP=%date:~0,4%%date:~5,2%%date:~8,2%_%time:~0,2%%time:~3,2%%time:~6,2%"
set "TIMESTAMP=%TIMESTAMP: =0%"
set "FILE_PATH=%LOG_DIR%\TinyWar_%TIMESTAMP%.txt"

echo Step 1: clear old logcat buffer...
adb logcat -c

echo Step 2: detect app process id...
for /f "tokens=1" %%p in ('adb shell pidof -s %PKG_NAME%') do set "APP_PID=%%p"

if "%APP_PID%"=="" (
    echo App process was not found. Capturing logs filtered by package name.
    echo Output: %FILE_PATH%
    adb logcat -v time | findstr "%PKG_NAME%" > "%FILE_PATH%"
) else (
    echo App PID: %APP_PID%
    echo Output: %FILE_PATH%
    adb logcat --pid=%APP_PID% -v time > "%FILE_PATH%"
)

pause
