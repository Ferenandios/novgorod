"""
Модуль для обработки данных из Excel и текстовых файлов
"""
import pandas as pd
import re

class DataProcessor:
    def __init__(self):
        pass
    
    def load_excel(self, filepath):
        """Загрузка данных из Excel файла"""
        try:
            df = pd.read_excel(filepath)
            # Очистка данных
            df = df.dropna(how='all')  # Удаление пустых строк
            return df
        except Exception as e:
            raise Exception(f"Ошибка при чтении Excel: {e}")
    
    def load_proc_txt(self, filepath):
        """Загрузка данных из текстового файла Proc.txt"""
        try:
            # Читаем файл с правильной кодировкой
            with open(filepath, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # Заменяем точки-разделители на табуляцию
            # Символ · (греческая точка, Unicode U+0387)
            content = content.replace('\u0387\u0387', '\t')  # Две точки -> табуляция
            content = content.replace('\u0387', '')  # Удаляем одиночные точки
            
            # Разделяем на строки
            lines = content.split('\n')
            
            # Парсинг данных
            data = []
            headers = None
            
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                
                # Разделение по табуляции
                parts = [p.strip() for p in line.split('\t') if p.strip()]
                
                if not parts:
                    continue
                
                # Первая строка - заголовки
                if headers is None:
                    headers = parts
                    continue
                
                # Дополнение или обрезка до нужной длины
                while len(parts) < len(headers):
                    parts.append('')
                if len(parts) > len(headers):
                    parts = parts[:len(headers)]
                
                data.append(parts)
            
            if headers and data:
                df = pd.DataFrame(data, columns=headers)
                return df
            else:
                raise Exception("Не удалось распознать структуру файла")
        except Exception as e:
            raise Exception(f"Ошибка при чтении текстового файла: {e}")
    
    def merge_data(self, elements_df, proc_df):
        """Объединение данных из двух источников"""
        try:
            # Попытка объединить по общему полю (например, Designator)
            if 'Designator' in elements_df.columns and 'Designator' in proc_df.columns:
                merged = pd.merge(elements_df, proc_df, 
                                on='Designator', 
                                how='outer',
                                suffixes=('_elem', '_proc'))
            else:
                # Если нет общего поля, просто объединяем по индексу
                merged = pd.concat([elements_df, proc_df], axis=1)
            
            return merged
        except Exception as e:
            raise Exception(f"Ошибка при объединении данных: {e}")
    
    def validate_data(self, df):
        """Валидация данных перед генерацией документа"""
        errors = []
        
        # Проверка на пустые обязательные поля
        required_fields = ['Designator', 'Footprint']
        for field in required_fields:
            if field in df.columns:
                if df[field].isna().any():
                    errors.append(f"Найдены пустые значения в поле {field}")
        
        return errors
