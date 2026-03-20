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

:: ── Python 설치 확인 ──────────────────────────────────────────────────────
python --version > nul 2>&1
if %errorLevel% neq 0 (
    echo.
    echo ❌ Python이 설치되어 있지 않습니다.
    echo    https://www.python.org 에서 Python 설치 후 다시 실행해주세요.
    echo.
    pause
    exit /b
)

:: ── 패키지 설치 (최초 1회) ───────────────────────────────────────────────
echo 📦 필요 패키지 확인 중...
pip install -q -r requirements.txt
playwright install chromium --quiet 2>nul

:: ── 자동화 실행 ─────────────────────────────────────────────────────────
echo.
echo 🚀 sMES 키팅 자동화를 시작합니다...
echo.
python kitting_automation.py --step all

echo.
pause
