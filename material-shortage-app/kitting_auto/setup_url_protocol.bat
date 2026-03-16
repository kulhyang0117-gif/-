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

:: ── 레지스트리 등록 ───────────────────────────────────────────────────────
reg add "HKEY_CLASSES_ROOT\mes-kit"                         /f /ve /d "URL:MES Kitting Automation"
reg add "HKEY_CLASSES_ROOT\mes-kit"                         /f /v "URL Protocol" /d ""
reg add "HKEY_CLASSES_ROOT\mes-kit\DefaultIcon"             /f /ve /d "%RUN_BAT%,0"
reg add "HKEY_CLASSES_ROOT\mes-kit\shell\open\command"      /f /ve /d "cmd.exe /c \"\"%RUN_BAT%\"\" \"%1\""

echo.
echo ✅ mes-kit:// 프로토콜 등록 완료!
echo.
echo    이제 자재부족현황.html의 [🤖 키팅 자동화] 버튼을 클릭하면
echo    자동으로 sMES 다운로드 및 업로드가 실행됩니다.
echo.
pause
