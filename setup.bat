@echo off
chcp 65001 > nul
echo.
echo   ============================================================
echo     PPTSlide2Video — 초기 설치
echo   ============================================================
echo.

echo   [1/3] Python 패키지 설치 중...
pip install -r "%~dp0requirements.txt"
if %errorlevel% neq 0 (
    echo.
    echo   [오류] pip 실행 실패. Python이 설치되어 있는지 확인하세요.
    echo          https://www.python.org/downloads/
    pause
    exit /b 1
)

echo.
echo   [2/3] FFmpeg 확인 중...
ffmpeg -version >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo   [경고] FFmpeg가 설치되어 있지 않습니다!
    echo          아래에서 다운로드 후 PATH에 추가하세요:
    echo          https://ffmpeg.org/download.html
    echo.
) else (
    echo          FFmpeg OK
)

echo   [3/3] PowerPoint 확인 중...
where POWERPNT >nul 2>&1
if %errorlevel% neq 0 (
    echo          PowerPoint를 직접 확인하세요 (Office 365 또는 2019+)
) else (
    echo          PowerPoint OK
)

echo.
echo   ============================================================
echo     설치 완료!
echo   ============================================================
echo.
echo     사용법:
echo       1. PPTX 파일과 음성 파일(01.wav, 02.wav...)을 같은 폴더에 준비
echo       2. PPTX 파일을 convert.bat 위에 드래그 앤 드롭
echo.
pause
