$ErrorActionPreference = "Stop"

python -m PyInstaller --noconfirm --clean --onefile --name Sales_Data_Collector main.py
python -m PyInstaller --noconfirm --clean --onefile --windowed --name Sales_Data_Collector_GUI gui.py

Write-Host "Built executables:"
Write-Host "  dist\Sales_Data_Collector.exe"
Write-Host "  dist\Sales_Data_Collector_GUI.exe"
