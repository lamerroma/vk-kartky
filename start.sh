#!/bin/bash
cd "$(dirname "$0")"

# Створити venv якщо не існує
if [ ! -d "venv" ]; then
    echo "Створення віртуального середовища..."
    python3 -m venv venv
fi

# Встановити flask через pip всередині venv (не системний pip!)
venv/bin/pip install flask -q

# Запуск через python всередині venv
echo "Запуск програми 'Відділ кадрів'..."
venv/bin/python app.py
