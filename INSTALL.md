# Инструкция по установке

## Быстрый старт

### 1. Проверка Python
```bash
python --version
```
Должна быть версия 3.8 или выше.

### 2. Установка зависимостей
```bash
pip install -r requirements.txt
```

### 3. Запуск приложения
```bash
python main.py
```

## Подробная установка

### Windows

1. **Установка Python**
   - Скачайте установщик с https://www.python.org/downloads/
   - Запустите установщик
   - ✓ Отметьте "Add Python to PATH"
   - Нажмите "Install Now"

2. **Установка зависимостей**
   - Откройте командную строку (Win+R, введите `cmd`)
   - Перейдите в папку с проектом:
     ```bash
     cd путь\к\проекту
     ```
   - Установите библиотеки:
     ```bash
     pip install -r requirements.txt
     ```

3. **Запуск**
   ```bash
   python main.py
   ```

### Альтернативный способ (создание исполняемого файла)

Для создания .exe файла используйте PyInstaller:

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name="RouteCardGenerator" main.py
```

Исполняемый файл будет создан в папке `dist/`.

## Проверка установки

Запустите тестовый скрипт:

```python
python -c "from data_processor import DataProcessor; print('OK')"
```

Если выводится "OK", установка прошла успешно.

## Возможные проблемы

### Ошибка: "pip не является внутренней командой"
**Решение**: Python не добавлен в PATH. Переустановите Python с опцией "Add to PATH".

### Ошибка: "ModuleNotFoundError: No module named 'pandas'"
**Решение**: Установите зависимости:
```bash
pip install pandas openpyxl python-docx
```

### Ошибка: "Permission denied"
**Решение**: Запустите командную строку от имени администратора.

## Обновление

Для обновления зависимостей:
```bash
pip install --upgrade -r requirements.txt
```

## Удаление

Для удаления программы:
1. Удалите папку с проектом
2. (Опционально) Удалите установленные библиотеки:
   ```bash
   pip uninstall pandas openpyxl python-docx pillow
   ```
