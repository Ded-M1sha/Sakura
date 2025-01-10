import tkinter as tk
from tkinter import simpledialog, filedialog, messagebox
from form1 import process_form1
from form2 import process_form2
from form3 import process_form3
from form4 import process_form4

# Создание главного окна приложения
root = tk.Tk()
root.title("Sakura")

# Переменные для хранения путей к файлам
form1_filepath = None


# Функция для запроса верхней границы объема
def get_upper_limit():
    upper_limit = simpledialog.askfloat("Ввод верхней границы объема",
                                        "Введите верхнюю границу объема (введите 0 для автоматического расчета):")
    if upper_limit is None:
        return None  # Если пользователь отменил ввод
    return upper_limit


# Обработчик для выбора и обработки файла "Форма 1"
def select_form1_file(progress_var):
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not filepath:
        return None  # Если файл не выбран

    # Запрашиваем верхнюю границу объема у пользователя
    upper_limit = get_upper_limit()
    if upper_limit is None:  # Пользователь отменил ввод
        messagebox.showinfo("Отмена", "Обработка формы 1 была отменена.")
        return None

    try:
        # Пытаемся обработать файл "Форма 1"
        process_form1(filepath, progress_var, root, lambda: on_form_done(1), upper_limit)
        return filepath  # Возвращаем путь к файлу формы 1 для последующей обработки формы 2
    except Exception as e:
        messagebox.showerror("Ошибка обработки формы 1", f"Не удалось обработать форму 1.\nОшибка: {e}")
        return None





# Обработчик для выбора уже обработанного файла "Форма 1"
def load_existing_form1_file():
    global form1_filepath
    form1_filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not form1_filepath:
        messagebox.showinfo("Отмена", "Файл формы 1 не был загружен.")
    else:
        messagebox.showinfo("Форма 1 загружена", "Обработанный файл формы 1 успешно загружен.")


# Обработчик для выбора и обработки файла "Форма 2"
def select_form2_file(progress_var):
    global form1_filepath
    if not form1_filepath:
        messagebox.showinfo("Требуется форма 1", "Для обработки формы 2 необходимо сначала обработать или загрузить форму 1.")
        return

    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not filepath:
        messagebox.showinfo("Отмена", "Обработка формы 2 была отменена.")
        return None

    try:
        process_form2(filepath, form1_filepath, progress_var, root, lambda: on_form_done(2))
        return filepath  # Возвращаем путь к файлу формы 2 для последующей обработки формы 3
    except Exception as e:
        messagebox.showerror("Ошибка обработки формы 2", f"Не удалось обработать форму 2.\nОшибка: {e}")
        return None


# Обработчик для выбора и обработки файла "Форма 3"
def select_form3_file(progress_var):
    global form1_filepath
    if not form1_filepath:
        messagebox.showinfo("Требуется форма 1", "Для обработки формы 3 необходимо сначала обработать или загрузить форму 1.")
        return

    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not filepath:
        messagebox.showinfo("Отмена", "Обработка формы 3 была отменена.")
        return

    try:
        process_form3(filepath, form1_filepath, progress_var, root, lambda: on_form_done(3))
    except Exception as e:
        messagebox.showerror("Ошибка обработки формы 3", f"Не удалось обработать форму 3.\nОшибка: {e}")

# Обработчик для выбора и обработки файла "Форма 4"
def select_form4_file(progress_var):
    global form1_filepath
    if not form1_filepath:
        messagebox.showinfo("Требуется форма 1", "Для обработки формы 3 необходимо сначала обработать или загрузить форму 1.")
        return

    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not filepath:
        messagebox.showinfo("Отмена", "Обработка формы 3 была отменена.")
        return

    try:
        process_form4(filepath, form1_filepath, progress_var, root, lambda: on_form_done(4))
    except Exception as e:
        messagebox.showerror("Ошибка обработки формы 4", f"Не удалось обработать форму 4.\nОшибка: {e}")

# Завершающий обработчик для каждой формы
def on_form_done(form_number):
    messagebox.showinfo(f"Форма {form_number} обработана", f"Файл формы {form_number} успешно обработан.")


# Функция для запуска обработки форм
def start_processing():
    selected_forms = [form_var1.get(), form_var2.get(), form_var3.get(), form_var4.get()]

    # Проверка, что выбрана хотя бы одна форма
    if not any(selected_forms):
        messagebox.showinfo("Выбор форм", "Выберите хотя бы одну форму для обработки.")
        return

    # Переменные для хранения путей к файлам
    global form1_filepath

    # Обработка формы 1, если выбрана
    if selected_forms[0] and not form1_filepath:
        form1_filepath = select_form1_file(progress_var)
        if not form1_filepath:
            return  # Останавливаем обработку, если форма 1 не была успешно обработана

    # Обработка формы 2, если выбрана
    if selected_forms[1]:
        select_form2_file(progress_var)

    # Обработка формы 3, если выбрана
    if selected_forms[2]:
        select_form3_file(progress_var)

    if selected_forms[3]:
        select_form4_file(progress_var)



    # Завершаем работу программы
    messagebox.showinfo("Завершение работы", "Обработка выбранных форм завершена.")
    root.quit()


# Создаем метку и кнопки для выбора форм
instruction_label = tk.Label(root, text="Какие формы вы хотите обработать?")
instruction_label.pack()

# Переменные для хранения выбора пользователя
form_var1 = tk.BooleanVar()
form_var2 = tk.BooleanVar()
form_var3 = tk.BooleanVar()
form_var4 = tk.BooleanVar()

# Чекбоксы для выбора форм
checkbox1 = tk.Checkbutton(root, text="Форма 1", variable=form_var1)
checkbox1.pack()
checkbox2 = tk.Checkbutton(root, text="Форма 2", variable=form_var2)
checkbox2.pack()
checkbox3 = tk.Checkbutton(root, text="Форма 3", variable=form_var3)
checkbox3.pack()
checkbox4 = tk.Checkbutton(root, text="Форма 4", variable=form_var4)
checkbox4.pack()

# Кнопка для загрузки уже обработанного файла формы 1
load_form1_button = tk.Button(root, text="Загрузить обработанную форму 1", command=load_existing_form1_file)
load_form1_button.pack()


# Кнопка для начала обработки
start_button = tk.Button(root, text="Начать обработку", command=start_processing)
start_button.pack()

# Переменная и метка для отображения состояния обработки
progress_var = tk.StringVar()
progress_var.set("Выберите формы для обработки.")
progress_label = tk.Label(root, textvariable=progress_var)
progress_label.pack()

root.mainloop()
