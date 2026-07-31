@echo off
title WOS Tool
chcp 65001 >nul
echo =======================================================
echo  WOS Item Aggregation Tool
echo =======================================================
python run.py
echo.
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] An error occurred. Please check the message above.
) else (
    echo Completed successfully!
)
echo.
pause
