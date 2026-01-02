@echo off
REM Activate virtual environment and run login test
echo ============================================================
echo Running Login and Status Check Test
echo ============================================================
echo.

REM Activate virtual environment if it exists
if exist ".venv\Scripts\activate.bat" (
    echo Activating virtual environment...
    call .venv\Scripts\activate.bat
) else (
    echo Virtual environment not found, using system Python...
)

echo.
echo Running test script...
echo.

python test_login_complete.py

echo.
echo ============================================================
echo Test completed!
echo ============================================================

pause
