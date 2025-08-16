@echo off
REM Build CommunityCrawler executable using PyInstaller

python -m venv .venv
call .venv\Scripts\activate

pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller

pyinstaller ^
  --onefile --noconsole ^
  --name CommunityCrawler ^
  --add-binary "chromedriver.exe;." ^
  --icon icon.ico ^
  --collect-all pandas ^
  --collect-all openpyxl ^
  community_crawler_gui_hours.py

echo.
echo Build complete. The executable is in the dist directory.
