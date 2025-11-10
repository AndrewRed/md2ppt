#!/bin/bash
# Скрипт запуска конвертера Markdown в PowerPoint

# Получаем директорию скрипта
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR"

# Проверяем наличие venv
if [ ! -d "venv" ]; then
    echo "Создание виртуального окружения..."
    python3 -m venv venv
fi

# Активируем venv
echo "Активация виртуального окружения..."
source venv/bin/activate

# Устанавливаем зависимости
echo "Установка зависимостей..."
pip install -q --upgrade pip
pip install -q -r requirements.txt

# Запускаем GUI приложение
echo "Запуск приложения..."
python3 md_to_pptx_gui.py

# Деактивируем venv при выходе
deactivate

