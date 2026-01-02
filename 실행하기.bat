@echo off
:: 한글 출력을 위해 인코딩을 UTF-8로 변경
chcp 65001
cls

echo 특허 관리 시스템을 실행 중입니다... 잠시만 기다려 주세요.

:: 현재 폴더의 내장 파이썬으로 streamlit 실행
:: --server.headless=true 옵션은 브라우저를 자동으로 띄우기 위함입니다.
start "" "python\python.exe" -m streamlit run app.py --server.headless=true --global.developmentMode=false

:: 또는 브라우저가 안 뜬다면 아래 줄 사용
:: "python\python.exe" -m streamlit run app.py