#!/usr/bin/env python3
"""
GUI приложение для конвертации Markdown в PowerPoint
"""
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import sys
from md_to_pptx import convert_markdown_to_pptx

# Промпт для языковой модели
PROMPT_TEMPLATE = """## Системный промпт для создания презентаций из Markdown

**Ты — ассистент, который создает Markdown файлы для конвертации в PowerPoint презентации.**

### Твоя задача

1. Запомни правила форматирования ниже.
2. Дождись, когда пользователь пришлет данные (текст, тезисы, информацию).
3. Создай Markdown файл по этим правилам на основе полученных данных.

### Правила форматирования (запомни их)

#### Структура документа

1. Основные разделы: используй `##` (два хеша) для каждого основного раздела
   - Первый раздел станет заголовком титульного слайда
   - Названия разделов должны быть краткими и понятными

2. Подразделы: используй `###` (три хеша) для подразделов внутри основного раздела

3. Специальные разделы:
   - Раздел "Введение" обрабатывается отдельно (если используется)
   - Разделы с "Кейс" в названии группируются по 2 на слайд

#### Форматирование контента

1. Списки:
   - Маркированные: `-`, `*` или `+`
   - Нумерованные: `1.`, `2.`, и т.д.
   - Каждый пункт — одна строка
   - Оптимально: 5–7 пунктов на раздел

2. Заголовки подразделов: используй формат `**Заголовок:**` для выделения секций внутри раздела

3. Таблицы: используй стандартный Markdown формат
   ```
   | Заголовок 1 | Заголовок 2 | Заголовок 3 |
   |-------------|-------------|-------------|
   | Данные 1    | Данные 2    | Данные 3    |
   ```
   - Максимум 8 строк и 5 столбцов для читаемости

4. Текст:
   - Обычный текст длиннее 15 символов будет включен
   - Можно использовать `**жирный текст**` для выделения
   - Избегай блоков кода (```) — они будут удалены

#### Ограничения и рекомендации

- На слайд помещается до 7 пунктов (для введения — до 10)
- Если пунктов больше, они разбиваются на несколько слайдов
- Таблицы автоматически ограничиваются 8 строками и 5 столбцами
- Эмодзи в тексте поддерживаются
- Избегай горизонтальных разделителей `---` — они игнорируются
- Не используй код-блоки — они удаляются при обработке

### Пример структуры

```markdown
## Название презентации

Это первый раздел — он станет заголовком титульного слайда.

## Введение

- Первый пункт введения
- Второй пункт введения
- Третий пункт введения

## Основной раздел

- Пункт 1
- Пункт 2
- Пункт 3

### Подраздел
- Детали подраздела 1
- Детали подраздела 2

## Кейсы использования

### Кейс 1: Название
- Описание кейса 1
- Преимущества 1
- Результаты 1

### Кейс 2: Название
- Описание кейса 2
- Преимущества 2
- Результаты 2

## Сравнительная таблица

| Параметр | Вариант A | Вариант B |
|----------|-----------|-----------|
| Цена     | 100       | 150       |
| Качество | Высокое   | Среднее   |
```

### Важные замечания

- Первый раздел (`##`) станет заголовком титульного слайда
- Разделы с подразделами создают отдельные слайды для каждого подраздела
- Разделы "Кейс" автоматически группируются по 2 на слайд
- Таблицы создают отдельные слайды с табличным форматом
- Все Markdown разметка (жирный, курсив, ссылки) будет очищена, но текст сохранится

---

### Инструкция для работы

1. Подтверди, что ты запомнил правила выше.
2. Сообщи, что готов принять данные от пользователя.
3. Когда пользователь пришлет данные, создай Markdown файл по этим правилам.
4. Выведи только готовый Markdown код, без дополнительных комментариев (если не попросят).

Готов к работе. Жду данные для создания презентации."""

def get_font(family, size, weight="normal"):
    """Возвращает кортеж шрифта с fallback для кроссплатформенности"""
    # Попытка использовать современные шрифты с fallback
    preferred_fonts = {
        'default': ('Segoe UI', 'Helvetica Neue', 'Arial', 'sans-serif'),
        'mono': ('Consolas', 'Monaco', 'Courier New', 'monospace')
    }
    
    # Выбираем первый доступный шрифт из списка
    font_family = family
    if family in preferred_fonts:
        # Для tkinter используем первый шрифт, система сама выберет доступный
        font_family = preferred_fonts[family][0]
    
    if weight == "bold":
        return (font_family, size, "bold")
    return (font_family, size)

class MarkdownToPPTXApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Конвертер Markdown → PowerPoint")
        self.root.geometry("720x680")
        self.root.resizable(False, False)
        
        # Современная цветовая схема
        self.colors = {
            'bg_primary': '#f8f9fa',
            'bg_secondary': '#ffffff',
            'bg_accent': '#e9ecef',
            'primary': '#0066cc',
            'primary_hover': '#0052a3',
            'success': '#28a745',
            'text_primary': '#212529',
            'text_secondary': '#6c757d',
            'border': '#dee2e6',
            'shadow': '#adb5bd'
        }
        
        # Настройка стиля окна
        self.root.configure(bg=self.colors['bg_primary'])
        
        # Переменные
        self.input_file = tk.StringVar()
        self.output_file = tk.StringVar()
        
        # Создаем интерфейс
        self.create_widgets()
        
        # Центрируем окно
        self.center_window()
    
    def center_window(self):
        """Центрирует окно на экране"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def create_widgets(self):
        """Создает виджеты интерфейса"""
        # Главный контейнер с отступами
        main_container = tk.Frame(self.root, bg=self.colors['bg_primary'], padx=40, pady=25)
        main_container.pack(fill=tk.BOTH, expand=True)
        
        # Заголовок с иконкой
        header_frame = tk.Frame(main_container, bg=self.colors['bg_primary'])
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        title_label = tk.Label(
            header_frame,
            text="📄 Конвертер Markdown в PowerPoint",
            font=get_font('default', 24, 'bold'),
            bg=self.colors['bg_primary'],
            fg=self.colors['text_primary'],
            pady=0
        )
        title_label.pack()
        
        subtitle_label = tk.Label(
            header_frame,
            text="Преобразуйте ваши Markdown файлы в профессиональные презентации",
            font=get_font('default', 13),
            bg=self.colors['bg_primary'],
            fg=self.colors['text_secondary'],
            pady=8
        )
        subtitle_label.pack()
        
        # Кнопка копирования промпта
        prompt_button_frame = tk.Frame(main_container, bg=self.colors['bg_primary'])
        prompt_button_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.copy_prompt_button = tk.Button(
            prompt_button_frame,
            text="📋 Скопировать промпт для языковой модели",
            command=self.copy_prompt_to_clipboard,
            font=get_font('default', 11),
            bg=self.colors['bg_accent'],
            fg=self.colors['text_primary'],
            relief=tk.FLAT,
            bd=0,
            padx=15,
            pady=8,
            cursor="hand2",
            takefocus=False,
            activebackground='#d0d3d6',
            activeforeground=self.colors['text_primary']
        )
        self.copy_prompt_button.pack()
        
        # Фрейм для выбора входного файла
        input_frame = tk.Frame(main_container, bg=self.colors['bg_secondary'], relief=tk.FLAT, bd=0)
        input_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Внутренний фрейм с отступами
        input_inner = tk.Frame(input_frame, bg=self.colors['bg_secondary'], padx=20, pady=18)
        input_inner.pack(fill=tk.BOTH, expand=True)
        
        input_label = tk.Label(
            input_inner,
            text="📥 Входной файл (Markdown)",
            font=get_font('default', 14, 'bold'),
            bg=self.colors['bg_secondary'],
            fg=self.colors['text_primary'],
            anchor=tk.W
        )
        input_label.pack(anchor=tk.W, pady=10)
        
        input_file_frame = tk.Frame(input_inner, bg=self.colors['bg_secondary'])
        input_file_frame.pack(fill=tk.X)
        
        self.input_entry = tk.Entry(
            input_file_frame,
            textvariable=self.input_file,
            font=get_font('default', 13),
            state="readonly",
            relief=tk.SOLID,
            bd=1,
            bg='#ffffff',
            fg=self.colors['text_primary'],
            readonlybackground='#ffffff',
            insertbackground=self.colors['text_primary'],
            highlightthickness=1,
            highlightcolor=self.colors['primary'],
            highlightbackground=self.colors['border']
        )
        self.input_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10), ipady=10)
        
        input_button = tk.Button(
            input_file_frame,
            text="Выбрать файл",
            command=self.browse_input_file,
            font=get_font('default', 12, 'bold'),
            bg=self.colors['bg_accent'],
            fg=self.colors['text_primary'],
            relief=tk.FLAT,
            bd=0,
            padx=20,
            pady=10,
            cursor="hand2",
            takefocus=False,
            activebackground='#d0d3d6',
            activeforeground=self.colors['text_primary']
        )
        input_button.pack(side=tk.RIGHT)
        
        # Фрейм для выбора выходного файла
        output_frame = tk.Frame(main_container, bg=self.colors['bg_secondary'], relief=tk.FLAT, bd=0)
        output_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Внутренний фрейм с отступами
        output_inner = tk.Frame(output_frame, bg=self.colors['bg_secondary'], padx=20, pady=18)
        output_inner.pack(fill=tk.BOTH, expand=True)
        
        output_label = tk.Label(
            output_inner,
            text="📤 Выходной файл (PowerPoint)",
            font=get_font('default', 14, 'bold'),
            bg=self.colors['bg_secondary'],
            fg=self.colors['text_primary'],
            anchor=tk.W
        )
        output_label.pack(anchor=tk.W, pady=10)
        
        output_file_frame = tk.Frame(output_inner, bg=self.colors['bg_secondary'])
        output_file_frame.pack(fill=tk.X)
        
        self.output_entry = tk.Entry(
            output_file_frame,
            textvariable=self.output_file,
            font=get_font('default', 13),
            relief=tk.SOLID,
            bd=1,
            bg='#ffffff',
            fg=self.colors['text_primary'],
            insertbackground=self.colors['text_primary'],
            highlightthickness=1,
            highlightcolor=self.colors['primary'],
            highlightbackground=self.colors['border']
        )
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10), ipady=10)
        
        output_button = tk.Button(
            output_file_frame,
            text="Выбрать файл",
            command=self.browse_output_file,
            font=get_font('default', 12, 'bold'),
            bg=self.colors['bg_accent'],
            fg=self.colors['text_primary'],
            relief=tk.FLAT,
            bd=0,
            padx=20,
            pady=10,
            cursor="hand2",
            takefocus=False,
            activebackground='#d0d3d6',
            activeforeground=self.colors['text_primary']
        )
        output_button.pack(side=tk.RIGHT)
        
        # Кнопка конвертации
        button_frame = tk.Frame(main_container, bg=self.colors['bg_primary'])
        button_frame.pack(fill=tk.X, pady=(10, 15))
        
        convert_button = tk.Button(
            button_frame,
            text="🚀 Конвертировать",
            command=self.convert,
            font=get_font('default', 18, 'bold'),
            bg=self.colors['primary'],
            fg=self.colors['text_primary'],
            relief=tk.FLAT,
            bd=0,
            padx=50,
            pady=18,
            cursor="hand2",
            takefocus=False,
            activebackground=self.colors['primary_hover'],
            activeforeground=self.colors['text_primary']
        )
        convert_button.pack()
        
        # Статус бар
        status_frame = tk.Frame(main_container, bg=self.colors['bg_primary'])
        status_frame.pack(fill=tk.X)
        
        self.status_label = tk.Label(
            status_frame,
            text="✨ Готов к работе",
            font=get_font('default', 12),
            fg=self.colors['text_secondary'],
            bg=self.colors['bg_primary'],
            pady=5
        )
        self.status_label.pack()
    
    def browse_input_file(self):
        """Открывает диалог выбора входного файла"""
        filename = filedialog.askopenfilename(
            title="Выберите Markdown файл",
            filetypes=[("Markdown files", "*.md"), ("All files", "*.*")]
        )
        if filename:
            self.input_file.set(filename)
            # Автоматически генерируем имя выходного файла
            base_name = os.path.splitext(os.path.basename(filename))[0]
            directory = os.path.dirname(filename)
            output_path = os.path.join(directory, f"{base_name}.pptx")
            self.output_file.set(output_path)
    
    def browse_output_file(self):
        """Открывает диалог выбора выходного файла"""
        filename = filedialog.asksaveasfilename(
            title="Сохранить PowerPoint файл",
            defaultextension=".pptx",
            filetypes=[("PowerPoint files", "*.pptx"), ("All files", "*.*")]
        )
        if filename:
            self.output_file.set(filename)
    
    def copy_prompt_to_clipboard(self):
        """Копирует промпт в буфер обмена"""
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(PROMPT_TEMPLATE)
            self.root.update()  # Обновляем буфер обмена
            
            # Показываем сообщение об успехе
            messagebox.showinfo(
                "✅ Скопировано",
                "Промпт успешно скопирован в буфер обмена!\n\n"
                "Теперь вы можете вставить его в языковую модель."
            )
            # Убираем фокус с кнопок после закрытия messagebox
            self._reset_button_focus()
        except Exception as e:
            messagebox.showerror(
                "❌ Ошибка",
                f"Не удалось скопировать промпт:\n{str(e)}"
            )
            # Убираем фокус с кнопок после закрытия messagebox
            self._reset_button_focus()
    
    def _reset_button_focus(self):
        """Сбрасывает фокус со всех кнопок"""
        # Используем after_idle для сброса состояния после обработки всех событий
        def reset_focus():
            # Убираем фокус со всех кнопок
            if hasattr(self, 'copy_prompt_button'):
                self.copy_prompt_button.config(relief=tk.FLAT, state=tk.NORMAL)
            # Устанавливаем фокус на root окно
            self.root.focus_force()
            # Убираем выделение с любого виджета
            self.root.focus_set()
        
        self.root.after_idle(reset_focus)
    
    def convert(self):
        """Выполняет конвертацию"""
        input_path = self.input_file.get()
        output_path = self.output_file.get()
        
        # Валидация
        if not input_path:
            messagebox.showerror("Ошибка", "Пожалуйста, выберите входной файл")
            return
        
        if not os.path.exists(input_path):
            messagebox.showerror("Ошибка", f"Файл не найден: {input_path}")
            return
        
        if not output_path:
            messagebox.showerror("Ошибка", "Пожалуйста, укажите выходной файл")
            return
        
        # Обновляем статус
        self.status_label.config(
            text="⏳ Конвертация в процессе...",
            fg=self.colors['primary'],
            font=get_font('default', 12, 'bold')
        )
        self.root.update()
        
        try:
            # Выполняем конвертацию
            output_file, slide_count = convert_markdown_to_pptx(input_path, output_path)
            
            # Показываем успешное сообщение
            messagebox.showinfo(
                "✅ Успех",
                f"Презентация успешно создана!\n\n"
                f"📄 Файл: {os.path.basename(output_file)}\n"
                f"📊 Всего слайдов: {slide_count}\n\n"
                f"📁 Путь: {output_file}"
            )
            
            self.status_label.config(
                text=f"✅ Готово! Создано {slide_count} слайдов",
                fg=self.colors['success'],
                font=get_font('default', 12, 'bold')
            )
            
        except Exception as e:
            error_msg = str(e)
            messagebox.showerror(
                "❌ Ошибка",
                f"Ошибка при конвертации:\n\n{error_msg}"
            )
            self.status_label.config(
                text="❌ Ошибка при конвертации",
                fg="#dc3545",
                font=get_font('default', 12, 'bold')
            )

def main():
    """Запускает GUI приложение"""
    root = tk.Tk()
    app = MarkdownToPPTXApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()

