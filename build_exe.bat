@echo off
setlocal
cd /d "%~dp0"

echo ============================================
echo  Build WeChatAccountExporter (single-file EXE)
echo ============================================
echo.

echo [1/3] Installing runtime dependencies...
python -m pip install -r requirements.txt
if errorlevel 1 goto :failed
python -m pip install pyinstaller
if errorlevel 1 goto :failed

echo.
echo [2/3] Building with PyInstaller...
python -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name WeChatAccountExporter ^
  --hidden-import Crypto.Cipher.AES ^
  app.py
if errorlevel 1 goto :failed

echo.
echo [3/3] Collecting artifact...
if not exist "dist\WeChatAccountExporter.exe" goto :failed
copy /Y "dist\WeChatAccountExporter.exe" ".\WeChatAccountExporter.exe" >nul
echo.
echo Done. EXE is at:
echo   %cd%\WeChatAccountExporter.exe
echo.
pause
exit /b 0

:failed
echo.
echo Build FAILED. Read the messages above.
pause
exit /b 1
