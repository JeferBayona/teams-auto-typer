$ErrorActionPreference = 'Stop'

Write-Host "Creating Python Virtual Environment..."
python -m venv venv

Write-Host "Activating Virtual Environment..."
.\venv\Scripts\Activate.ps1

Write-Host "Installing dependencies..."
pip install pandas openpyxl xlrd pyautogui pyinstaller html5lib beautifulsoup4

Write-Host "Building executable..."
pyinstaller --noconfirm --onefile --windowed --name "TeamsAutoTyper"  "app_brute.py"

Write-Host "Build complete! The executable is located in the 'dist' folder."
