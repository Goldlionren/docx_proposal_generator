@echo off
if not exist ".venv\Scripts\python.exe" (
    python -m venv .venv
)
".venv\Scripts\python.exe" -m pip install -r requirements.txt
".venv\Scripts\python.exe" -m pip install pyinstaller
".venv\Scripts\python.exe" -m PyInstaller --onefile --windowed --name "DOCXProposalGenerator" --collect-all pypandoc --hidden-import pythoncom --hidden-import win32com --hidden-import win32com.client main.py
pause
