# 📦 무설치(Portable) 배포판 제작 가이드

`PyInstaller`로 만든 EXE 파일이 백신 프로그램에 의해 차단되거나 실행 속도가 느린 경우, **Python Embeddable Package (내장형 파이썬)**를 활용하여 폴더 자체를 배포하는 방법을 권장합니다.

이 방식은 사용자의 PC에 Python이 설치되어 있지 않아도, 폴더 안에 포함된 독립적인 Python 엔진으로 프로그램을 실행합니다.

---

## 1. 준비물 다운로드

1.  **Python Embeddable Package 다운로드**:
    *   [Python 공식 다운로드 페이지](https://www.python.org/downloads/windows/) 접속.
    *   현재 개발 환경과 동일한 버전(예: 3.9, 3.11 등)의 **`Windows embeddable package (64-bit)`** zip 파일을 다운로드합니다.
2.  **배포 폴더 생성**:
    *   바탕화면에 새 폴더를 만들고 이름을 정합니다. (예: `PatentManager_Portable`)

---

## 2. 폴더 구조 구성

다운로드한 zip을 배포 폴더 안에 `python`이라는 이름의 폴더로 압축을 풉니다. 그리고 소스 코드와 데이터를 해당 위치에 복사합니다.

최종 폴더 구조는 아래와 같아야 합니다:

```text
PatentManager_Portable/
├── python/                # (1) 내장형 파이썬 압축 푼 폴더
│   ├── python.exe         #     (확인용: 이 파일이 있어야 함)
│   └── ...
├── app.py                 # (2) 메인 소스 코드 복사
├── app_utils.py           #     (보조 파이썬 파일이 있다면 같이 복사)
├── patent_rawfile.xlsx    # (3) 엑셀 데이터 파일 복사
└── 실행하기.bat            # (4) 실행 스크립트 (직접 생성)
```

---

## 3. 필수 설정 (중요!)

내장형 파이썬이 외부 라이브러리(`pandas`, `streamlit` 등)를 인식할 수 있도록 설정 파일을 수정해야 합니다.

1.  `python` 폴더 안으로 들어갑니다.
2.  `python3xx._pth` 파일을 찾습니다. (예: `python311._pth`)
3.  메모장으로 엽니다.
4.  마지막 줄에 있는 **`#import site`**의 주석을 제거하여 **`import site`**로 바꿉니다.
5.  저장하고 닫습니다.

---

## 4. 라이브러리 설치

사용자의 PC에는 라이브러리가 없으므로, **이 폴더 안에** 필요한 패키지들을 모두 설치해서 넣어줘야 합니다.

1.  명령 프롬프트(CMD)를 엽니다.
2.  배포 폴더(`PatentManager_Portable`) 위치로 이동합니다.
3.  아래 명령어를 입력하여 로컬 폴더에 라이브러리를 설치합니다. (인터넷 연결 필요)

```bash
# 주의: 아래 경로는 실제 python 폴더 위치에 맞게 수정하세요.
# pip 모듈을 사용하여 현재 폴더의 python/Lib/site-packages 에 설치하는 명령입니다.

.\python\python.exe -m pip install streamlit pandas openpyxl altair --target=.\python\Lib\site-packages
```

*   만약 `No module named pip` 에러가 난다면: [get-pip.py](https://bootstrap.pypa.io/get-pip.py)를 다운로드 받아 `python` 폴더에 넣고 `.\python\python.exe get-pip.py`를 먼저 실행하세요.

---

## 5. 실행 스크립트 만들기

사용자가 복잡한 명령어를 칠 수 없으므로, 더블 클릭만 하면 실행되는 `.bat` 파일을 만듭니다.

1.  배포 폴더(`PatentManager_Portable`)에 메모장을 엽니다.
2.  아래 내용을 복사해서 붙여넣습니다.

```batch
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
```

3.  파일 이름을 **`실행하기.bat`** (또는 `특허관리시스템_실행.bat`)로 저장합니다. 반드시 확장자가 `.bat`이어야 합니다.

---

## 6. 배포 하기

이제 `PatentManager_Portable` 폴더 전체를 **압축(.zip)**하여 사내에 공유하면 됩니다.
사용자는 압축을 풀고 **`실행하기.bat`**만 더블 클릭하면 됩니다.

---

### ✅ 장점
*   **백신 충돌 해결**: EXE 패키징 방식이 아니므로 바이러스 오진 가능성이 매우 낮음.
*   **환경 독립성**: 사용자 PC에 파이썬이 깔려있든 없든 무관하게 실행됨.
*   **빠른 실행**: 별도의 압축 해제 과정 없이 즉시 실행됨.

---

## 7. 문제 해결 (Troubleshooting)

### Q. 실행창에 "?뱁뿀 愿由??쒖뒪?쒖쓣..." 처럼 글자가 깨져서 나와요.
**A. 한글 인코딩 문제입니다.**
윈도우 콘솔(CMD)의 기본 인코딩(CP949)과 파이썬 출력(UTF-8)이 맞지 않아 발생합니다.
`실행하기.bat` 파일을 메모장으로 열고, 맨 윗줄 `@echo off` 바로 아래에 다음 명령어를 추가하세요.

```batch
chcp 65001
```

이 명령어는 콘솔 창의 문자셋을 UTF-8로 강제 변경하여 한글을 정상적으로 표시해줍니다. (가이드의 [5. 실행 스크립트 만들기] 예제 코드에는 이미 포함되어 있습니다.)
