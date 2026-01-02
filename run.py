import streamlit.web.cli as stcli
import os, sys

def resolve_path(path):
    """
    exe로 변환되었을 때 내부 임시 경로(sys._MEIPASS)를 찾아주는 함수
    """
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, path)
    return os.path.join(os.path.abspath("."), path)

if __name__ == "__main__":
    # exe 실행 시 app.py의 경로를 찾아서 실행 명령어로 전달
    sys.argv = [
        "streamlit",
        "run",
        resolve_path("app.py"),
        "--global.developmentMode=false"
    ]
    sys.exit(stcli.main())