@echo off
cd /d "%~dp0"
echo ============================================
echo   Збірка VK_Kartky.exe
echo ============================================

echo [1/3] Встановлення залежностей...
pip install flask pyinstaller -q

echo [2/3] Збірка .exe файлу...
pyinstaller app.spec --clean --noconfirm

echo [3/3] Готово!
echo.
echo Файл знаходиться в папці: dist\VK_Kartky.exe
echo.
pause
