@echo off
echo ===================================================
echo             Starting Labs Tracker...
echo ===================================================
echo.
echo Opening browser at http://127.0.0.1:8000/ ...
start http://127.0.0.1:8000/
echo.
echo Starting FastAPI Web Server...
python -m uvicorn main:app --host 127.0.0.1 --port 8000 --reload
if %ERRORLEVEL% neq 0 (
    echo.
    echo [ERROR] Failed to start uvicorn. Please make sure python and requirements are installed.
    pause
)
