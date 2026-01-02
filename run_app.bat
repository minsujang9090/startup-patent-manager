@echo off
REM ====================================================
REM IoT Log Analysis System Run Script
REM Environment: Anaconda 'python_dev'
REM ====================================================

echo [INFO] Initializing Conda from C:\Users\mroon\anaconda3\Scripts\activate.bat...
call C:\Users\mroon\anaconda3\Scripts\activate.bat

echo [INFO] Activating 'python_patent' environment...
call conda activate python_patent

IF %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Failed to activate 'python_patent' environment.
    echo Please check if the environment 'python_patent' exists.
    pause
    exit /b
)

echo [INFO] Environment 'python_patent' activated.
echo [INFO] Checking/Installing dependencies...
pip install -r requirements.txt

echo [INFO] Starting Streamlit App...
streamlit run app.py

pause
