@echo off
setlocal
REM 需要 Windows + Python（若你電腦可以裝 Python，這是最快方式）

python -m pip install --upgrade pip
python -m pip install --upgrade pyinstaller
if exist requirements.txt python -m pip install -r requirements.txt

pyinstaller --noconfirm --clean --onefile --windowed --name "程式內容檢查工具" ^
  --hidden-import=tkinterdnd2 ^
  --hidden-import=tkinterdnd2.TkinterDnD ^
  main.py

echo Done. See dist\程式內容檢查工具.exe
pause
endlocal
