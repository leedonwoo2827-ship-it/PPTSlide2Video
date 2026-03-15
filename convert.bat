@echo off
chcp 65001 > nul
echo.
echo   ============================================================
echo     PPTSlide2Video — PPTX → MP4 변환
echo   ============================================================
echo.

if "%~1"=="" (
    echo   사용법:
    echo     convert.bat "발표자료.pptx"
    echo     convert.bat "발표자료.pptx" "음성폴더"
    echo     convert.bat "발표자료.pptx" "음성폴더" "결과.mp4"
    echo.
    echo   PPTX 파일을 이 BAT 파일 위에 드래그 앤 드롭해도 됩니다!
    echo.
    pause
    exit /b 1
)

python "%~dp0run_local.py" %*
pause
