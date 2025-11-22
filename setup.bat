@echo off
echo ========================================
echo   Company Contact Scraper Setup
echo ========================================
echo.

echo [1/3] Installing dependencies...
pip install -r requirements.txt
echo.

echo [2/3] Checking for Gemini API key...
if "%GEMINI_API_KEY%"=="" (
    echo WARNING: GEMINI_API_KEY not set!
    echo.
    echo Please set your API key:
    echo   1. Get API key from: https://aistudio.google.com/app/apikey
    echo   2. Run: set GEMINI_API_KEY=your-api-key-here
    echo.
) else (
    echo API key found!
)
echo.

echo [3/3] Setup complete!
echo.
echo To run the scraper:
echo   python scraper.py
echo.
echo ========================================
pause
