import os
from tkinter import Tk, Label, Entry, Button, Text, filedialog, messagebox, END
from docx import Document

def log_message(message):
    """Функция для записи сообщений в лог."""
    log_area.insert(END, message + '\n')
    log_area.see(END)  # Прокрутка вниз, чтобы видеть последние сообщения

def replace_text_in_docx(file_path, search_text, replace_text):
    doc = Document(file_path)
    text_replaced = False  # Флаг для отслеживания замен
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            if search_text in run.text:
                run.text = run.text.replace(search_text, replace_text)
                text_replaced = True  # Установить флаг, если текст заменен
    if text_replaced:
        doc.save(file_path)
        log_message(f"Заменен текст в файле: {file_path}")  # Логирование
    else:
        log_message(f"Текст не найден в файле: {file_path}")  # Логирование

def replace_text_in_multiple_docs(directory, search_text, replace_text):
    for filename in os.listdir(directory):
        if filename.endswith('.docx'):
            file_path = os.path.join(directory, filename)
            replace_text_in_docx(file_path, search_text, replace_text)

def browse_directory():
    directory = filedialog.askdirectory()
    if directory:
        dir_entry.delete(0, 'end')
        dir_entry.insert(0, directory)

def start_replacement():
    directory = dir_entry.get()
    search_text = search_text_area.get("1.0", 'end-1c').strip()  # Удаляем лишние пробелы
    replace_text = replace_text_area.get("1.0", 'end-1c')

    if not directory or not search_text or not replace_text:
        messagebox.showerror("Ошибка", "Пожалуйста, заполните все поля.")
        return

    log_message("Начинаем замену текста...")
    replace_text_in_multiple_docs(directory, search_text, replace_text)
    log_message("Замена текста завершена!")

def copy_to_clipboard(text_widget):
    text = text_widget.get("1.0", 'end-1c')
    root.clipboard_clear()  # Очистить буфер обмена
    root.clipboard_append(text)  # Добавить текст в буфер обмена

def paste_from_clipboard(text_widget):
    try:
        text = root.clipboard_get()  # Получить текст из буфера обмена
        text_widget.delete("1.0", END)  # Очистить поле перед вставкой
        text_widget.insert("1.0", text)  # Вставить текст в виджет
    except Exception as e:
        log_message("Ошибка при вставке текста из буфера обмена.")

# Создание графического интерфейса
root = Tk()
root.title("Замена текста в документах Word")

# Разрешаем растягивание окна
root.columnconfigure(0, weight=1)
root.columnconfigure(1, weight=1)
root.rowconfigure(5, weight=1)

Label(root, text="Папка с документами:").grid(row=0, column=0, padx=5, pady=5, sticky='ew')
dir_entry = Entry(root, width=40)
dir_entry.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
Button(root, text="Обзор", command=browse_directory).grid(row=0, column=2, padx=5, pady=5)

Label(root, text="Текст для замены:").grid(row=1, column=0, padx=5, pady=5, sticky='ew')
search_text_area = Text(root, width=40, height=5)
search_text_area.grid(row=1, column=1, padx=5, pady=5, sticky='ew')

Button(root, text="Копировать", command=lambda: copy_to_clipboard(search_text_area)).grid(row=1, column=2, padx=5, pady=5)
Button(root, text="Вставить", command=lambda: paste_from_clipboard(search_text_area)).grid(row=1, column=3, padx=5, pady=5)

Label(root, text="Новый текст:").grid(row=2, column=0, padx=5, pady=5, sticky='ew')
replace_text_area = Text(root, width=40, height=5)  # Многострочное поле
replace_text_area.grid(row=2, column=1, padx=5, pady=5, sticky='ew')

Button(root, text="Копировать", command=lambda: copy_to_clipboard(replace_text_area)).grid(row=2, column=2, padx=5, pady=5)
Button(root, text="Вставить", command=lambda: paste_from_clipboard(replace_text_area)).grid(row=2, column=3, padx=5, pady=5)

Button(root, text="Начать замену", command=start_replacement).grid(row=3, column=1, padx=5, pady=10)

# Логирование
Label(root, text="Лог:").grid(row=4, column=0, padx=5, pady=5, sticky='ew')
log_area = Text(root, width=80, height=10)  # Поле для логов
log_area.grid(row=5, column=0, columnspan=4, padx=5, pady=5, sticky='nsew')

root.mainloop()
