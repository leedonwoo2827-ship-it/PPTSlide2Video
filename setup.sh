#!/bin/bash
echo ""
echo "  ============================================================"
echo "    PPTSlide2Video — 초기 설치"
echo "  ============================================================"
echo ""

echo "  [1/3] Python 패키지 설치 중..."
pip3 install -r "$(dirname "$0")/requirements.txt"
if [ $? -ne 0 ]; then
    echo ""
    echo "  [오류] pip 실행 실패. Python이 설치되어 있는지 확인하세요."
    echo "         https://www.python.org/downloads/"
    exit 1
fi

echo ""
echo "  [2/3] FFmpeg 확인 중..."
if command -v ffmpeg &> /dev/null; then
    echo "         FFmpeg OK"
else
    echo ""
    echo "  [경고] FFmpeg가 설치되어 있지 않습니다!"
    echo "         brew install ffmpeg  (macOS)"
    echo "         sudo apt install ffmpeg  (Ubuntu/Debian)"
    echo ""
fi

echo "  [3/3] PowerPoint 확인 중..."
if [[ "$OSTYPE" == "darwin"* ]]; then
    if [ -d "/Applications/Microsoft PowerPoint.app" ]; then
        echo "         PowerPoint OK"
    else
        echo "         PowerPoint가 설치되어 있는지 확인하세요 (Office 2013 이상)"
    fi
else
    echo "         PowerPoint를 직접 확인하세요 (Office 2013 이상)"
fi

echo ""
echo "  ============================================================"
echo "    설치 완료!"
echo "  ============================================================"
echo ""
echo "    사용법:"
echo "      1. PPTX 파일과 음성 파일(01.wav, 02.wav...)을 같은 폴더에 준비"
echo "      2. python3 run_local.py \"발표자료.pptx\""
echo ""
