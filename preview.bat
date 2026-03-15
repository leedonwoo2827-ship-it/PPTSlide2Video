@echo off
chcp 65001 > nul
echo PPT2SlideDeck Preview Server
echo ==============================
"C:\Users\leedonwoo\AppData\Local\Programs\Python\Python314\python.exe" "%~dp0preview.py" %*
pause
