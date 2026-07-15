@echo off
cd /d "%~dp0"
python -m pip install -r requirements.txt
python -m pip install pyinstaller
python -m PyInstaller --noconfirm --clean --onefile --windowed --name UnitCostLookup unit_cost_lookup.py
echo Built: dist\UnitCostLookup.exe
pause
