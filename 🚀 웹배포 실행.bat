@echo off
powershell -ExecutionPolicy Bypass -File "%~dp0deploy.ps1"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo  오류가 발생했습니다. 위 메시지를 확인하세요.
    pause
)
