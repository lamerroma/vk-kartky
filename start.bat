@echo off
cd /d "%~dp0"
echo Запуск програми "Відділ кадрів"...
python -m pip install flask -q
python app.py
pause
