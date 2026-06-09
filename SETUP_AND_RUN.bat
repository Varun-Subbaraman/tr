@echo off
echo ================================================
echo   SMC Trading Agent - Setup and Run
echo ================================================
echo.
echo Step 1: Installing required packages...
echo.
pip install yfinance pandas colorama openpyxl
echo.
echo ================================================
echo   Setup complete! Starting agent...
echo ================================================
echo.
python trading_agent.py
pause
