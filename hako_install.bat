@echo off
title Installer

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [X] Python is not installed or not added to your system PATH
    echo Please install Python from https://www.python.org/ and check the
    echo "Add Python to PATH" box during installation
    echo.
    pause
    exit /b
)

echo [✓] Python detected, upgrading pip
python -m pip install --upgrade pip

echo.
echo [?] Installing required external libraries
echo.
python -m pip install pandas numpy matplotlib scipy beautifulsoup4 curl_cffi python-dateutil gspread lxml adjustText html2image Pillow

if %errorlevel% equ 0 (
    echo.
    echo [✓] All dependencies installed successfully!
    echo.
    echo You can now run your analyzer pipeline by running run.py
    echo or executing it via your terminal
) else (
    echo [X] Something went wrong during the installation
    echo Please check the error messages above
)

echo.
pause