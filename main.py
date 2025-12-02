"""
Главный модуль приложения для формирования маршрутных карт
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pandas as pd
from pathlib import Path
from datetime import datetime
from data_processor import DataProcessor
from document_generator import DocumentGenerator
from preview_window import PreviewWindow
from edit_dialog import EditDialog

class RouteCardApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Генератор маршрутных карт v2.0")
        self.root.geometry("1400x900")
        
        self.data_processor = DataProcessor()
        self.doc_generator = DocumentGenerator()
        
        self.elements_data = None
        self.proc_data = None
        self.merged_data = None
        
        self.setup_ui()
    
    def setup_ui(self):
        """Создание интерфейса"""
        # Верхняя панель с кнопками
        top_frame = ttk.Frame(self.root, padding="10")
        top_frame.pack(fill=tk.X)
        
        # Кнопки загрузки
        ttk.Button(top_frame, text="📂 Загрузить Elements.xlsx", 
                  command=self.load_elements).pack(side=tk.LEFT, padx=5)
        ttk.Button(top_frame, text="📂 Загрузить Proc.txt", 
                  command=self.load_proc).pack(side=tk.LEFT, padx=5)
        
        # Разделитель
        ttk.Separator(top_frame, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)
        
        # Кнопки редактирования
        ttk.Button(top_frame, text="✏️ Редактировать", 
                  command=self.edit_selected).pack(side=tk.LEFT, padx=5)
        ttk.Button(top_frame, text="➕ Добавить строку", 
                  command=self.add_row).pack(side=tk.LEFT, padx=5)
        ttk.Button(top_frame, text="🗑️ Удалить строку", 
                  command=self.delete_row).pack(side=tk.LEFT, padx=5)
        
        # Разделитель
        ttk.Separator(top_frame, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)
        
        # Кнопки генерации
        ttk.Button(top_frame, text="👁️ Предпросмотр", 
                  command=self.preview_document).pack(side=tk.LEFT, padx=5)
        ttk.Button(top_frame, text="📄 Сохранить DOCX", 
                  command=self.generate_document).pack(side=tk.LEFT, padx=5)
        ttk.Button(top_frame, text="📑 Экспорт в PDF", 
                  command=self.export_to_pdf).pack(side=tk.LEFT, padx=5)
        
        # Область для отображения данных
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Вкладка Elements
        self.elements_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.elements_frame, text="Элементы")
        self.setup_elements_table()
        
        # Вкладка Proc
        self.proc_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.proc_frame, text="Процессы")
        self.setup_proc_table()
        
        # Статус бар
        self.status_var = tk.StringVar(value="Готов к работе")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, 
                              relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)
    
    def setup_elements_table(self):
        """Создание таблицы для элементов"""
        # Scrollbar
        scroll_y = ttk.Scrollbar(self.elements_frame, orient=tk.VERTICAL)
        scroll_x = ttk.Scrollbar(self.elements_frame, orient=tk.HORIZONTAL)
        
        self.elements_tree = ttk.Treeview(self.elements_frame,
                                         yscrollcommand=scroll_y.set,
                                         xscrollcommand=scroll_x.set)
        
        scroll_y.config(command=self.elements_tree.yview)
        scroll_x.config(command=self.elements_tree.xview)
        
        # Двойной клик для редактирования
        self.elements_tree.bind('<Double-Button-1>', lambda e: self.edit_selected())
        
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.elements_tree.pack(fill=tk.BOTH, expand=True)
    
    def setup_proc_table(self):
        """Создание таблицы для процессов"""
        scroll_y = ttk.Scrollbar(self.proc_frame, orient=tk.VERTICAL)
        scroll_x = ttk.Scrollbar(self.proc_frame, orient=tk.HORIZONTAL)
        
        self.proc_tree = ttk.Treeview(self.proc_frame,
                                     yscrollcommand=scroll_y.set,
                                     xscrollcommand=scroll_x.set)
        
        scroll_y.config(command=self.proc_tree.yview)
        scroll_x.config(command=self.proc_tree.xview)
        
        # Двойной клик для редактирования
        self.proc_tree.bind('<Double-Button-1>', lambda e: self.edit_selected())
        
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.proc_tree.pack(fill=tk.BOTH, expand=True)
    
    def load_elements(self):
        """Загрузка файла Elements.xlsx"""
        filename = filedialog.askopenfilename(
            title="Выберите файл Elements",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                self.elements_data = self.data_processor.load_excel(filename)
                self.display_elements()
                self.status_var.set(f"Загружено элементов: {len(self.elements_data)}")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось загрузить файл:\n{e}")
    
    def load_proc(self):
        """Загрузка файла Proc.txt"""
        filename = filedialog.askopenfilename(
            title="Выберите файл Proc",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                self.proc_data = self.data_processor.load_proc_txt(filename)
                self.display_proc()
                self.status_var.set(f"Загружено процессов: {len(self.proc_data)}")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось загрузить файл:\n{e}")
    
    def display_elements(self):
        """Отображение данных элементов"""
        # Очистка
        for item in self.elements_tree.get_children():
            self.elements_tree.delete(item)
        
        if self.elements_data is not None and not self.elements_data.empty:
            # Настройка колонок
            columns = list(self.elements_data.columns)
            self.elements_tree['columns'] = columns
            self.elements_tree['show'] = 'headings'
            
            for col in columns:
                self.elements_tree.heading(col, text=col)
                self.elements_tree.column(col, width=100)
            
            # Добавление данных
            for idx, row in self.elements_data.iterrows():
                values = [str(row[col]) for col in columns]
                self.elements_tree.insert('', tk.END, values=values)
    
    def display_proc(self):
        """Отображение данных процессов"""
        for item in self.proc_tree.get_children():
            self.proc_tree.delete(item)
        
        if self.proc_data is not None and not self.proc_data.empty:
            columns = list(self.proc_data.columns)
            self.proc_tree['columns'] = columns
            self.proc_tree['show'] = 'headings'
            
            for col in columns:
                self.proc_tree.heading(col, text=col)
                self.proc_tree.column(col, width=100)
            
            for idx, row in self.proc_data.iterrows():
                values = [str(row[col]) for col in columns]
                self.proc_tree.insert('', tk.END, values=values)
    
    def get_merged_data(self):
        """Получение объединенных данных"""
        if self.merged_data is not None:
            return self.merged_data
        
        if self.elements_data is None or self.proc_data is None:
            return None
        
        self.merged_data = self.data_processor.merge_data(
            self.elements_data, self.proc_data
        )
        return self.merged_data
    
    def edit_selected(self):
        """Редактирование выбранной строки"""
        # Определяем активную вкладку
        current_tab_index = self.notebook.index(self.notebook.select())
        
        if current_tab_index == 0:  # Вкладка "Элементы"
            tree = self.elements_tree
            data = self.elements_data
            data_name = "elements"
        elif current_tab_index == 1:  # Вкладка "Процессы"
            tree = self.proc_tree
            data = self.proc_data
            data_name = "proc"
        else:
            messagebox.showinfo("Информация", "Выберите строку для редактирования")
            return
        
        if data is None or data.empty:
            messagebox.showinfo("Информация", "Нет данных для редактирования")
            return
        
        selected = tree.selection()
        if not selected:
            messagebox.showinfo("Информация", "Выберите строку для редактирования")
            return
        
        # Получаем индекс выбранной строки
        item = selected[0]
        values = tree.item(item)['values']
        
        # Открываем диалог редактирования
        dialog = EditDialog(self.root, data.columns.tolist(), values)
        if dialog.result:
            # Обновляем данные
            row_idx = tree.index(item)
            for i, col in enumerate(data.columns):
                data.at[row_idx, col] = dialog.result[i]
            
            # Обновляем отображение
            if data_name == "elements":
                self.display_elements()
            else:
                self.display_proc()
            
            self.merged_data = None  # Сбрасываем кэш
            self.status_var.set("Данные обновлены")
    
    def add_row(self):
        """Добавление новой строки"""
        # Определяем активную вкладку
        current_tab_index = self.notebook.index(self.notebook.select())
        
        if current_tab_index == 0:  # Вкладка "Элементы"
            data = self.elements_data
            data_name = "elements"
        elif current_tab_index == 1:  # Вкладка "Процессы"
            data = self.proc_data
            data_name = "proc"
        else:
            messagebox.showinfo("Информация", "Выберите вкладку для добавления строки")
            return
        
        if data is None or data.empty:
            messagebox.showinfo("Информация", "Сначала загрузите данные")
            return
        
        # Создаем пустую строку
        new_row = {col: '' for col in data.columns}
        
        # Открываем диалог редактирования
        dialog = EditDialog(self.root, data.columns.tolist(), list(new_row.values()))
        if dialog.result:
            # Добавляем новую строку
            new_df = pd.DataFrame([dialog.result], columns=data.columns)
            if data_name == "elements":
                self.elements_data = pd.concat([self.elements_data, new_df], ignore_index=True)
                self.display_elements()
            else:
                self.proc_data = pd.concat([self.proc_data, new_df], ignore_index=True)
                self.display_proc()
            
            self.merged_data = None
            self.status_var.set("Строка добавлена")
    
    def delete_row(self):
        """Удаление выбранной строки"""
        # Определяем активную вкладку
        current_tab_index = self.notebook.index(self.notebook.select())
        
        if current_tab_index == 0:  # Вкладка "Элементы"
            tree = self.elements_tree
            data_name = "elements"
            data = self.elements_data
        elif current_tab_index == 1:  # Вкладка "Процессы"
            tree = self.proc_tree
            data_name = "proc"
            data = self.proc_data
        else:
            messagebox.showinfo("Информация", "Выберите строку для удаления")
            return
        
        if data is None or data.empty:
            messagebox.showinfo("Информация", "Нет данных для удаления")
            return
        
        selected = tree.selection()
        if not selected:
            messagebox.showinfo("Информация", "Выберите строку для удаления")
            return
        
        if messagebox.askyesno("Подтверждение", "Удалить выбранную строку?"):
            item = selected[0]
            row_idx = tree.index(item)
            
            if data_name == "elements":
                self.elements_data = self.elements_data.drop(row_idx).reset_index(drop=True)
                self.display_elements()
            else:
                self.proc_data = self.proc_data.drop(row_idx).reset_index(drop=True)
                self.display_proc()
            
            self.merged_data = None
            self.status_var.set("Строка удалена")
    
    def preview_document(self):
        """Предпросмотр документа"""
        merged_data = self.get_merged_data()
        if merged_data is None:
            messagebox.showwarning("Предупреждение", 
                                 "Загрузите оба файла перед предпросмотром")
            return
        
        try:
            preview = PreviewWindow(self.root, merged_data, self.doc_generator)
            self.status_var.set("Предпросмотр открыт")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось открыть предпросмотр:\n{e}")
    
    def generate_document(self):
        """Генерация маршрутной карты в DOCX"""
        merged_data = self.get_merged_data()
        if merged_data is None:
            messagebox.showwarning("Предупреждение", 
                                 "Загрузите оба файла перед генерацией")
            return
        
        # Запрос информации о документе
        doc_info = self.get_document_info()
        if doc_info is None:
            return
        
        try:
            output_path = filedialog.asksaveasfilename(
                defaultextension=".docx",
                filetypes=[("Word documents", "*.docx"), ("All files", "*.*")]
            )
            
            if output_path:
                self.doc_generator.create_route_card(merged_data, output_path, doc_info)
                self.status_var.set(f"Документ сохранен: {output_path}")
                messagebox.showinfo("Успех", "Маршрутная карта создана по ГОСТ 3.1118!")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать документ:\n{e}")
    
    def get_document_info(self):
        """Диалог для ввода информации о документе"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Информация о документе")
        dialog.geometry("500x300")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Поля ввода
        fields = {}
        labels = [
            ('product_name', 'Наименование изделия:', 'Печатный узел'),
            ('designation', 'Обозначение:', ''),
            ('developer', 'Разработал:', ''),
            ('date', 'Дата:', datetime.now().strftime('%d.%m.%Y'))
        ]
        
        for i, (key, label, default) in enumerate(labels):
            ttk.Label(dialog, text=label).grid(row=i, column=0, sticky=tk.W, padx=10, pady=5)
            entry = ttk.Entry(dialog, width=40)
            entry.insert(0, default)
            entry.grid(row=i, column=1, padx=10, pady=5)
            fields[key] = entry
        
        result = {}
        
        def on_ok():
            for key, entry in fields.items():
                result[key] = entry.get()
            dialog.destroy()
        
        def on_cancel():
            dialog.destroy()
        
        # Кнопки
        btn_frame = ttk.Frame(dialog)
        btn_frame.grid(row=len(labels), column=0, columnspan=2, pady=20)
        
        ttk.Button(btn_frame, text="OK", command=on_ok).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Отмена", command=on_cancel).pack(side=tk.LEFT, padx=5)
        
        # Ожидание закрытия диалога
        self.root.wait_window(dialog)
        
        return result if result else None
    
    def export_to_pdf(self):
        """Экспорт в PDF"""
        merged_data = self.get_merged_data()
        if merged_data is None:
            messagebox.showwarning("Предупреждение", 
                                 "Загрузите оба файла перед экспортом")
            return
        
        # Запрос информации о документе
        doc_info = self.get_document_info()
        if doc_info is None:
            return
        
        try:
            output_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
            )
            
            if output_path:
                # Сначала создаем DOCX во временном файле
                import tempfile
                with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp:
                    tmp_docx = tmp.name
                
                self.doc_generator.create_route_card(merged_data, tmp_docx, doc_info)
                
                # Конвертируем в PDF
                self.doc_generator.convert_to_pdf(tmp_docx, output_path)
                
                # Удаляем временный файл
                import os
                os.unlink(tmp_docx)
                
                self.status_var.set(f"PDF сохранен: {output_path}")
                messagebox.showinfo("Успех", "PDF файл создан!")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать PDF:\n{e}")

def main():
    root = tk.Tk()
    app = RouteCardApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
