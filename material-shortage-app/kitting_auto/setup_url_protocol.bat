@echo off
chcp 65001 > nul
title mes-kit:// 프로토콜 등록

:: ── 관리자 권한 확인 ──────────────────────────────────────────────────────
net session >nul 2>&1
if %errorLevel% neq 0 (
    powershell -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

set "RUN_BAT=%~dp0run.bat"
set "ELEVATED_BAT=%~dp0run_elevated.bat"
set "STOP_BAT=%~dp0stop.bat"

:: ── mes-kit://start 등록 ───────────────────────────────────────────────────
reg add "HKEY_CLASSES_ROOT\mes-kit"                         /f /ve /d "URL:MES Kitting Automation"
reg add "HKEY_CLASSES_ROOT\mes-kit"                         /f /v "URL Protocol" /d ""
reg add "HKEY_CLASSES_ROOT\mes-kit\DefaultIcon"             /f /ve /d "%RUN_BAT%,0"
reg add "HKEY_CLASSES_ROOT\mes-kit\shell\open\command"      /f /ve /d "cmd.exe /c \"\"%ELEVATED_BAT%\"\""

:: ── mes-kit-stop:// 등록 (긴급정지) ─────────────────────────────────────
reg add "HKEY_CLASSES_ROOT\mes-kit-stop"                    /f /ve /d "URL:MES Kitting Stop"
reg add "HKEY_CLASSES_ROOT\mes-kit-stop"                    /f /v "URL Protocol" /d ""
reg add "HKEY_CLASSES_ROOT\mes-kit-stop\shell\open\command" /f /ve /d "cmd.exe /c \"\"%STOP_BAT%\"\""

echo.
echo ✅ 프로토콜 등록 완료!
echo    mes-kit://       - 키팅 자동화 시작
echo    mes-kit-stop://  - 긴급정지
echo.
echo    setup_url_protocol.bat 을 다시 관리자로 실행해야 긴급정지가 활성화됩니다.
echo.
pause
