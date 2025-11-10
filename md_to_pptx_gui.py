#!/usr/bin/env python3
"""
GUI приложение для конвертации Markdown в PowerPoint
"""
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import sys
from md_to_pptx import convert_markdown_to_pptx

class MarkdownToPPTXApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Конвертер Markdown → PowerPoint")
        self.root.geometry("600x300")
        self.root.resizable(False, False)
        
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
        # Заголовок
        title_label = tk.Label(
            self.root,
            text="Конвертер Markdown в PowerPoint",
            font=("Arial", 16, "bold"),
            pady=20
        )
        title_label.pack()
        
        # Фрейм для выбора входного файла
        input_frame = tk.Frame(self.root, pady=10)
        input_frame.pack(fill=tk.X, padx=20)
        
        tk.Label(
            input_frame,
            text="Входной файл (Markdown):",
            font=("Arial", 10)
        ).pack(anchor=tk.W)
        
        input_file_frame = tk.Frame(input_frame)
        input_file_frame.pack(fill=tk.X, pady=5)
        
        self.input_entry = tk.Entry(
            input_file_frame,
            textvariable=self.input_file,
            font=("Arial", 10),
            state="readonly"
        )
        self.input_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        tk.Button(
            input_file_frame,
            text="Выбрать...",
            command=self.browse_input_file,
            width=12
        ).pack(side=tk.RIGHT)
        
        # Фрейм для выбора выходного файла
        output_frame = tk.Frame(self.root, pady=10)
        output_frame.pack(fill=tk.X, padx=20)
        
        tk.Label(
            output_frame,
            text="Выходной файл (PowerPoint):",
            font=("Arial", 10)
        ).pack(anchor=tk.W)
        
        output_file_frame = tk.Frame(output_frame)
        output_file_frame.pack(fill=tk.X, pady=5)
        
        self.output_entry = tk.Entry(
            output_file_frame,
            textvariable=self.output_file,
            font=("Arial", 10)
        )
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        tk.Button(
            output_file_frame,
            text="Выбрать...",
            command=self.browse_output_file,
            width=12
        ).pack(side=tk.RIGHT)
        
        # Кнопка конвертации
        convert_button = tk.Button(
            self.root,
            text="Конвертировать",
            command=self.convert,
            font=("Arial", 12, "bold"),
            bg="#0066cc",
            fg="white",
            padx=20,
            pady=10,
            cursor="hand2"
        )
        convert_button.pack(pady=20)
        
        # Статус бар
        self.status_label = tk.Label(
            self.root,
            text="Готов к работе",
            font=("Arial", 9),
            fg="gray",
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
            if not self.output_file.get():
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
        self.status_label.config(text="Конвертация...", fg="blue")
        self.root.update()
        
        try:
            # Выполняем конвертацию
            output_file, slide_count = convert_markdown_to_pptx(input_path, output_path)
            
            # Показываем успешное сообщение
            messagebox.showinfo(
                "Успех",
                f"Презентация успешно создана!\n\n"
                f"Файл: {output_file}\n"
                f"Всего слайдов: {slide_count}"
            )
            
            self.status_label.config(
                text=f"Готово! Создано {slide_count} слайдов",
                fg="green"
            )
            
        except Exception as e:
            error_msg = str(e)
            messagebox.showerror("Ошибка", f"Ошибка при конвертации:\n{error_msg}")
            self.status_label.config(text="Ошибка при конвертации", fg="red")

def main():
    """Запускает GUI приложение"""
    root = tk.Tk()
    app = MarkdownToPPTXApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()

