@echo off
chcp 65001 >nul
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 goto nopython

echo 필요한 프로그램을 설치합니다. 잠시만 기다려주세요...
echo.
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo.
echo ========================================================
echo  설치가 완료되었습니다.
echo  SETUP.md 안내에 따라 credentials.json 파일을 이 폴더에
echo  넣은 뒤, run.bat 을 더블클릭해서 실행해주세요.
echo ========================================================
pause
exit /b 0

:nopython
echo [오류] 파이썬이 설치되어 있지 않습니다.
echo https://python.org 에서 파이썬을 설치한 뒤 다시 실행해주세요.
pause
exit /b 1
