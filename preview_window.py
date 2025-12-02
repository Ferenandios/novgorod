"""
Окно предпросмотра документа
"""
import tkinter as tk
from tkinter import ttk, messagebox
import tempfile
import os
import subprocess

class PreviewWindow:
    def __init__(self, parent, data, doc_generator):
        self.data = data
        self.doc_generator = doc_generator
        
        self.window = tk.Toplevel(parent)
        self.window.title("Предпросмотр маршрутной карты")
        self.window.geometry("1000x700")
        
        # Центрирование окна
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (1000 // 2)
        y = (self.window.winfo_screenheight() // 2) - (700 // 2)
        self.window.geometry(f"+{x}+{y}")
        
        self.setup_ui()
        self.generate_preview()
    
    def setup_ui(self):
        """Создание интерфейса"""
        # Верхняя панель
        top_frame = ttk.Frame(self.window, padding="10")
        top_frame.pack(fill=tk.X)
        
        ttk.Label(top_frame, text="Предпросмотр маршрутной карты", 
                 font=('Arial', 14, 'bold')).pack(side=tk.LEFT)
        
        ttk.Button(top_frame, text="🔄 Обновить", 
                  command=self.generate_preview).pack(side=tk.RIGHT, padx=5)
        ttk.Button(top_frame, text="📄 Открыть в Word", 
                  command=self.open_in_word).pack(side=tk.RIGHT, padx=5)
        
        # Область предпросмотра
        preview_frame = ttk.Frame(self.window)
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Текстовое поле с прокруткой
        scroll_y = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL)
        scroll_x = ttk.Scrollbar(preview_frame, orient=tk.HORIZONTAL)
        
        self.preview_text = tk.Text(preview_frame, 
                                    wrap=tk.NONE,
                                    yscrollcommand=scroll_y.set,
                                    xscrollcommand=scroll_x.set,
                                    font=('Courier', 10))
        
        scroll_y.config(command=self.preview_text.yview)
        scroll_x.config(command=self.preview_text.xview)
        
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.preview_text.pack(fill=tk.BOTH, expand=True)
        
        # Статистика
        stats_frame = ttk.Frame(self.window, padding="10")
        stats_frame.pack(fill=tk.X)
        
        self.stats_label = ttk.Label(stats_frame, text="")
        self.stats_label.pack(side=tk.LEFT)
        
        # Кнопки действий
        button_frame = ttk.Frame(self.window, padding="10")
        button_frame.pack(fill=tk.X)
        
        ttk.Button(button_frame, text="💾 Сохранить DOCX", 
                  command=self.save_docx).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="📑 Сохранить PDF", 
                  command=self.save_pdf).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Закрыть", 
                  command=self.window.destroy).pack(side=tk.RIGHT, padx=5)
    
    def generate_preview(self):
        """Генерация предпросмотра"""
        try:
            self.preview_text.delete('1.0', tk.END)
            
            # Заголовок
            self.preview_text.insert(tk.END, "=" * 80 + "\n")
            self.preview_text.insert(tk.END, "МАРШРУТНАЯ КАРТА".center(80) + "\n")
            self.preview_text.insert(tk.END, "=" * 80 + "\n\n")
            
            # Информация
            self.preview_text.insert(tk.END, "Наименование изделия: Печатный узел\n")
            self.preview_text.insert(tk.END, "Обозначение: \n")
            self.preview_text.insert(tk.END, "Дата: \n\n")
            
            self.preview_text.insert(tk.END, "-" * 80 + "\n\n")
            
            # Данные
            if self.data is not None and not self.data.empty:
                # Определение колонок для отображения
                display_columns = []
                for col in self.data.columns:
                    if not self.data[col].isna().all():
                        display_columns.append(col)
                
                # Заголовки
                header = " | ".join([f"{col[:15]:15}" for col in display_columns])
                self.preview_text.insert(tk.END, header + "\n")
                self.preview_text.insert(tk.END, "-" * len(header) + "\n")
                
                # Строки данных
                for idx, row in self.data.iterrows():
                    row_text = " | ".join([
                        f"{str(row[col])[:15]:15}" 
                        for col in display_columns
                    ])
                    self.preview_text.insert(tk.END, row_text + "\n")
                
                # Статистика
                stats = f"Всего строк: {len(self.data)} | Колонок: {len(display_columns)}"
                self.stats_label.config(text=stats)
            else:
                self.preview_text.insert(tk.END, "Нет данных для отображения\n")
            
            self.preview_text.config(state=tk.DISABLED)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать предпросмотр:\n{e}")
    
    def open_in_word(self):
        """Открытие документа в Word"""
        try:
            # Создаем временный файл
            with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp:
                tmp_path = tmp.name
            
            self.doc_generator.create_route_card(self.data, tmp_path)
            
            # Открываем в Word
            if os.name == 'nt':  # Windows
                os.startfile(tmp_path)
            else:  # Linux/Mac
                subprocess.call(['xdg-open', tmp_path])
            
            messagebox.showinfo("Информация", 
                              "Документ открыт в Word.\nВременный файл будет удален при закрытии.")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось открыть документ:\n{e}")
    
    def save_docx(self):
        """Сохранение в DOCX"""
        from tkinter import filedialog
        
        output_path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word documents", "*.docx"), ("All files", "*.*")]
        )
        
        if output_path:
            try:
                self.doc_generator.create_route_card(self.data, output_path)
                messagebox.showinfo("Успех", f"Документ сохранен:\n{output_path}")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось сохранить:\n{e}")
    
    def save_pdf(self):
        """Сохранение в PDF"""
        from tkinter import filedialog
        import tempfile
        
        output_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        
        if output_path:
            try:
                # Создаем временный DOCX
                with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp:
                    tmp_docx = tmp.name
                
                self.doc_generator.create_route_card(self.data, tmp_docx)
                self.doc_generator.convert_to_pdf(tmp_docx, output_path)
                
                os.unlink(tmp_docx)
                
                messagebox.showinfo("Успех", f"PDF сохранен:\n{output_path}")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось сохранить PDF:\n{e}")
