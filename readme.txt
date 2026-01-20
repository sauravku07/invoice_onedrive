1. Install Tesseract OCR
2. Login to OneDrive
3. Create folder: OneDrive/Invoices
4. Double-click invoice_watcher.exe
5. Put files in Invoices/Input
Win + R → shell:startup
Paste invoice_watcher.exe shortcut

python -m pip install --upgrade pip
python -m pip install -r requirements.txt

python -m pip install --upgrade pip
python -m pip install pyinstaller
python -m PyInstaller --version
pip install pyinstaller watchdog
convert to exe--python -m PyInstaller --onefile --noconsole invoice_watcher.py
install tesseract  -https://github.com/UB-Mannheim/tesseract/wiki
