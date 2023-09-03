import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import PyPDF2
import os
import re
import shutil
import time
import threading

def merge_pdfs(pdf_files, output_folder, done_folder, trebles_folder):
    # Создаем словарь для хранения пар файлов "face" и "back" по номеру
    pdf_dict = {}

    try:
        for pdf_file in pdf_files:
            # Используем регулярное выражение для извлечения номера и части (face/back) из имени файла
            match = re.search(r"\((\d+)-(\d+)\)-(\w+)\.pdf", pdf_file)
            if match:
                start_num, end_num, part = match.groups()
                key = f"{start_num}-{end_num}"
                if key not in pdf_dict:
                    pdf_dict[key] = {}
                pdf_dict[key][part] = pdf_file

        # Объединяем пары файлов "face" и "back"
        for key, pdf_data in pdf_dict.items():
            if 'face' in pdf_data and 'back' in pdf_data:
                # Извлекаем имя первого исходного файла для использования в имени объединенного файла
                first_source_file = pdf_data['face']
                output_filename = os.path.join(output_folder, os.path.basename(first_source_file))

                pdf_merger = PyPDF2.PdfMerger()

                try:
                    pdf_merger.append(pdf_data['face'])
                    pdf_merger.append(pdf_data['back'])

                    with open(output_filename, 'wb') as output_file:
                        pdf_merger.write(output_file)

                    pdf_merger.close()

                    # Копируем исходные файлы в папку "done"
                    shutil.copy(pdf_data['face'], done_folder)
                    shutil.copy(pdf_data['back'], done_folder)

                    # Удаляем исходные файлы
                    os.remove(pdf_data['face'])
                    os.remove(pdf_data['back'])
                except PyPDF2.errors.PdfReadError as e:
                    raise Exception(f"Ошибка при объединении PDF-файлов: {str(e)}")
            else:
                # Если нет пары, перемещаем файл в папку "trebles" с ограничением на количество попыток
                for pdf_file in pdf_data.values():
                    max_attempts = 3  # Максимальное количество попыток перемещения
                    for attempt in range(max_attempts):
                        try:
                            # Копируем файл в папку "trebles"
                            shutil.copy(pdf_file, os.path.join(trebles_folder, os.path.basename(pdf_file)))

                            # Если копирование успешно, удаляем исходный файл
                            os.remove(pdf_file)
                            break
                        except PermissionError:
                            print(f"Файл {pdf_file} занят другим процессом. Повторная попытка через 1 секунду...")
                            time.sleep(1)
                            if attempt == max_attempts - 1:
                                print(f"Не удалось переместить файл {pdf_file} после {max_attempts} попыток.")
    except Exception as e:
        messagebox.showerror("Ошибка", str(e))
        status_label.config(text=f"Ошибка: {str(e)}")
    else:
        # После успешного завершения объединения, показать информационное сообщение
        messagebox.showinfo("Готово", "Объединение PDF-файлов выполнено без ошибок.")
        # После успешного завершения объединения, обновить статус
        status_label.config(text="Статус: завершено без ошибок")

def browse_folder():
    folder_selected = filedialog.askdirectory()
    work_dir_entry.delete(0, tk.END)
    work_dir_entry.insert(0, folder_selected)

def start_processing():
    work_dir = work_dir_entry.get()  # Получаем путь из поля ввода

    # Проверяем, что путь к рабочей директории был выбран
    if not work_dir:
        messagebox.showerror("Ошибка", "Пожалуйста, выберите рабочую директорию.")
        return

    pdf_files = [os.path.join(work_dir, filename) for filename in os.listdir(work_dir) if filename.endswith('.pdf')]

    if not pdf_files:
        messagebox.showinfo("Информация", "В выбранной директории нет PDF-файлов для объединения.")
        return

    # Создаем папки "ready," "trebles," и "originals" внутри рабочей директории, если они не существуют
    yes_folder = os.path.join(work_dir, "ready")
    if not os.path.exists(yes_folder):
        os.makedirs(yes_folder)

    trebles_folder = os.path.join(work_dir, "trebles")
    if not os.path.exists(trebles_folder):
        os.makedirs(trebles_folder)

    done_folder = os.path.join(work_dir, "originals")
    if not os.path.exists(done_folder):
        os.makedirs(done_folder)

    # Запускаем обработку в отдельном потоке, чтобы не блокировать GUI
    processing_thread = threading.Thread(target=merge_pdfs, args=(pdf_files, yes_folder, done_folder, trebles_folder))
    processing_thread.start()

    status_label.config(text="Процесс объединения PDF-файлов запущен. Пожалуйста, подождите...")
# Создаем главное окно
root = tk.Tk()
# root.iconbitmap(r'pdf.ico')
root.title("Объединение PDF")

# Устанавливаем размеры окна
root.geometry("400x200")

# Создаем метку для отображения статуса работы программы
status_label = tk.Label(root, text="Статус: ожидание", wraplength=350)
status_label.pack(pady=10)

# Создаем поле для ввода пути к рабочей директории
work_dir_label = tk.Label(root, text="Выберите рабочую директорию:")
work_dir_label.pack()
work_dir_entry = tk.Entry(root, width=30)
work_dir_entry.pack()

# Создаем кнопку для выбора пути к рабочей директории
browse_button = tk.Button(root, text="Обзор", command=browse_folder)
browse_button.pack()

# Создаем кнопку "Старт" для запуска обработки
start_button = tk.Button(root, text="Старт", command=start_processing)
start_button.pack(pady=10)

# Получаем или создаем папки "yes," "trebles," и "done"______________________________________________________________
# input_folder = filedialog.askdirectory(initialdir="/path/to/starting/directory", title="Выберите каталог для сохранения файла")
input_folder = r"C:"

# Создаем папки "ready," "trebles," и "originals" внутри папки input_folder, если они не существуют
yes_folder = os.path.join(input_folder, "ready")
if not os.path.exists(yes_folder):
    os.makedirs(yes_folder)

trebles_folder = os.path.join(input_folder, "trebles")
if not os.path.exists(trebles_folder):
    os.makedirs(trebles_folder)

done_folder = os.path.join(input_folder, "originals")
if not os.path.exists(done_folder):
    os.makedirs(done_folder)

# Запускаем главный цикл обработки событий
root.mainloop()
