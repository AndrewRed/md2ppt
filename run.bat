@echo off
REM Скрипт запуска конвертера Markdown в PowerPoint для Windows

REM Получаем директорию скрипта
cd /d "%~dp0"

REM Проверяем наличие venv
if not exist "venv" (
    echo Создание виртуального окружения...
    python -m venv venv
)

REM Активируем venv
echo Активация виртуального окружения...
call venv\Scripts\activate.bat

REM Устанавливаем зависимости
echo Установка зависимостей...
python -m pip install -q --upgrade pip
python -m pip install -q -r requirements.txt

REM Запускаем GUI приложение
echo Запуск приложения...
python md_to_pptx_gui.py

REM Деактивируем venv при выходе
deactivate

pause

