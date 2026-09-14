@echo off
chcp 65001 >nul
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 goto nopython

python coupang_eats_csv.py
exit /b 0

:nopython
echo [오류] 파이썬이 설치되어 있지 않습니다.
echo https://python.org 에서 파이썬을 설치한 뒤 다시 실행해주세요.
pause
exit /b 1
