@echo off
chcp 65001 > nul
title sMES 키팅 자동화 - 권한 상승 중...

:: mes-kit:// 프로토콜에서 호출됨.
:: PowerShell로 run.bat을 관리자 권한으로 실행 (UAC 팝업 → 허용 클릭)
powershell -Command "Start-Process -FilePath '%~dp0run.bat' -WorkingDirectory '%~dp0' -Verb RunAs"
