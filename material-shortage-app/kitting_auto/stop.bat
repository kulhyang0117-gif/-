@echo off
chcp 65001 > nul
echo. > "%~dp0stop_flag.txt"
echo [%time%] 사용자 긴급정지 요청 >> "%~dp0kitting_log.txt"
