"""
Диалоговое окно для редактирования данных
"""
import tkinter as tk
from tkinter import ttk

class EditDialog:
    def __init__(self, parent, columns, values):
        self.result = None
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Редактирование данных")
        self.dialog.geometry("600x400")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Центрирование окна
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (600 // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (400 // 2)
        self.dialog.geometry(f"+{x}+{y}")
        
        self.columns = columns
        self.entries = {}
        
        self.setup_ui(values)
        
        # Ожидание закрытия окна
        self.dialog.wait_window()
    
    def setup_ui(self, values):
        """Создание интерфейса"""
        # Основной фрейм с прокруткой
        canvas = tk.Canvas(self.dialog)
        scrollbar = ttk.Scrollbar(self.dialog, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Создание полей ввода
        for i, col in enumerate(self.columns):
            label = ttk.Label(scrollable_frame, text=f"{col}:", width=20)
            label.grid(row=i, column=0, sticky=tk.W, padx=10, pady=5)
            
            entry = ttk.Entry(scrollable_frame, width=50)
            entry.grid(row=i, column=1, sticky=tk.EW, padx=10, pady=5)
            
            # Заполнение значениями
            if i < len(values):
                entry.insert(0, str(values[i]))
            
            self.entries[col] = entry
        
        scrollable_frame.columnconfigure(1, weight=1)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Кнопки
        button_frame = ttk.Frame(self.dialog)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(button_frame, text="Сохранить", 
                  command=self.save).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Отмена", 
                  command=self.cancel).pack(side=tk.RIGHT, padx=5)
    
    def save(self):
        """Сохранение изменений"""
        self.result = [self.entries[col].get() for col in self.columns]
        self.dialog.destroy()
    
    def cancel(self):
        """Отмена"""
        self.result = None
        self.dialog.destroy()
