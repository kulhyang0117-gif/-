@echo off
chcp 65001 > nul
title sMES 키팅 자동화

:: ── 관리자 권한 확인 및 자동 재실행 ─────────────────────────────────────────
net session >nul 2>&1
if %errorLevel% neq 0 (
    powershell -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

cd /d "%~dp0"

:: ── Python 3.12 설치 확인 ────────────────────────────────────────────────
py -3.12 --version > nul 2>&1
if %errorLevel% neq 0 (
    echo.
    echo ❌ Python 3.12가 설치되어 있지 않습니다.
    echo    python.org 에서 Python 3.12 설치 후 다시 실행해주세요.
    echo    설치 시 "Add python.exe to PATH" 반드시 체크!
    echo.
    pause
    exit /b
)

:: ── 패키지 설치 (최초 1회) ───────────────────────────────────────────────
echo 📦 필요 패키지 확인 중...
py -3.12 -m pip install -q -r requirements.txt
py -3.12 -m playwright install chromium --quiet 2>nul

:: ── 자동화 실행 ─────────────────────────────────────────────────────────
echo.
echo 🚀 sMES 키팅 자동화를 시작합니다...
echo.
py -3.12 kitting_automation.py

echo.
echo 📄 로그 파일 열기: kitting_log.txt
start notepad "%~dp0kitting_log.txt"

pause
