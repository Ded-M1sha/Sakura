import tkinter as tk
from tkinter import filedialog, messagebox
from form1PD import process_form1
from form2PD import process_form2
from form3PD import process_form3
from form4PD import process_form4
from itog import create_summary_from_memory

def main():
    def choose_file(var_label):
        """Функция для выбора файла и сохранения его пути в переменную."""
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            var_label.set(file_path)

    def on_form_done(form_number, new_file_path):
        """Функция вызывается при завершении обработки формы."""
        # Обновляем путь к обработанному файлу
        if form_number == 1:
            filepath_var1.set(new_file_path)
        elif form_number == 2:
            filepath_var2.set(new_file_path)
        elif form_number == 3:
            filepath_var3.set(new_file_path)
        elif form_number == 4:
            filepath_var4.set(new_file_path)

        progress_var.set(f"Форма {form_number} обработана успешно.")
        root.update_idletasks()  # Обновляем интерфейс

    def process_form(form_number):
        """Функция для обработки формы по номеру."""
        try:
            if form_number == 1:
                process_form1(
                    filepath_var1.get(),
                    progress_var,
                    root,
                    lambda new_fp: on_form_done(1, new_fp)
                )
            elif form_number == 2:
                process_form2(
                    filepath_var2.get(),
                    filepath_var1.get(),
                    progress_var,
                    root,
                    lambda new_fp: on_form_done(2, new_fp)
                )
            elif form_number == 3:
                process_form3(
                    filepath_var3.get(),
                    filepath_var1.get(),
                    progress_var,
                    root,
                    lambda new_fp: on_form_done(3, new_fp)
                )
            elif form_number == 4:
                process_form4(
                    filepath_var4.get(),
                    filepath_var1.get(),
                    progress_var,
                    root,
                    lambda new_fp: on_form_done(4, new_fp)
                )
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось обработать форму {form_number}. Ошибка: {e}")
            print(e)
            progress_var.set(f"Ошибка при обработке формы {form_number}.")
            root.update_idletasks()
    def create_summary():
        """Создание итогового сводного файла из обработанных форм."""
        try:
            forms_data = {
                2: filepath_var2.get(),
                3: filepath_var3.get(),
                4: filepath_var4.get()
            }
            output_file = filedialog.asksaveasfilename(filetypes=[("Excel Files", "*.xlsx")], defaultextension=".xlsx")
            if output_file:
                create_summary_from_memory(forms_data, output_file)
                messagebox.showinfo("Успешно", f"Итоговый файл сохранен по пути: {output_file}")
                progress_var.set("Итоговый файл создан успешно.")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать итоговый файл. Ошибка: {e}")
            progress_var.set("Ошибка при создании итогового файла.")
    # Инициализация GUI
    root = tk.Tk()
    root.title("Обработка форм")
    progress_var = tk.StringVar()
    tk.Label(root, text="Выберите файлы:").grid(row=0, column=0, columnspan=2)
    # Переменные для хранения путей к файлам
    filepath_var1 = tk.StringVar()
    filepath_var2 = tk.StringVar()
    filepath_var3 = tk.StringVar()
    filepath_var4 = tk.StringVar()
    # Кнопки для выбора файлов
    tk.Button(root, text="Форма 1", command=lambda: choose_file(filepath_var1)).grid(row=1, column=0)
    tk.Entry(root, textvariable=filepath_var1, width=50).grid(row=1, column=1)
    tk.Button(root, text="Форма 2", command=lambda: choose_file(filepath_var2)).grid(row=2, column=0)
    tk.Entry(root, textvariable=filepath_var2, width=50).grid(row=2, column=1)
    tk.Button(root, text="Форма 3", command=lambda: choose_file(filepath_var3)).grid(row=3, column=0)
    tk.Entry(root, textvariable=filepath_var3, width=50).grid(row=3, column=1)
    tk.Button(root, text="Форма 4", command=lambda: choose_file(filepath_var4)).grid(row=4, column=0)
    tk.Entry(root, textvariable=filepath_var4, width=50).grid(row=4, column=1)
    # Кнопки обработки форм
    tk.Button(root, text="Обработать форму 1", command=lambda: process_form(1)).grid(row=5, column=0)
    tk.Button(root, text="Обработать форму 2", command=lambda: process_form(2)).grid(row=6, column=0)
    tk.Button(root, text="Обработать форму 3", command=lambda: process_form(3)).grid(row=7, column=0)
    tk.Button(root, text="Обработать форму 4", command=lambda: process_form(4)).grid(row=8, column=0)
    # Кнопка для создания итогового файла
    tk.Button(root, text="Создать итоговый файл", command=create_summary).grid(row=9, column=0, columnspan=2)
    # Прогресс-бар
    tk.Label(root, textvariable=progress_var).grid(row=10, column=0, columnspan=2)
    progress_var.set("Готово к работе.")
    root.mainloop()

if __name__ == "__main__":
    main()