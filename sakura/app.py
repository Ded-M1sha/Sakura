import customtkinter as ctk
from tkinter import filedialog, messagebox
from sakura.forms import form1, form2, form3, form4
from sakura.processing import summary
# from sakura.utils import lines
# from sakura.processing.summary import create_summary_from_memory


def main():
    ctk.set_appearance_mode("ddark")  # Темная тема
    ctk.set_default_color_theme("green")  # Синяя цветовая схема

    def choose_file(var_label):
        """Функция для выбора файла и сохранения его пути в переменную."""
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            var_label.set(file_path)

    def on_form_done(form_number, new_file_path):
        """Функция вызывается при завершении обработки формы."""
        if form_number == 1:
            filepath_var1.set(new_file_path)
        elif form_number == 2:
            filepath_var2.set(new_file_path)
        elif form_number == 3:
            filepath_var3.set(new_file_path)
        elif form_number == 4:
            filepath_var4.set(new_file_path)

        progress_var.set(f"Форма {form_number} обработана успешно")
        progress_bar.set(1)  # Обновляем прогресс-бар
        root.update_idletasks()

    def process_form(form_number):
        """Функция для обработки формы по номеру."""
        try:
            progress_bar.set(0.5)  # Устанавливаем промежуточный прогресс
            if form_number == 1:
                form1.process_form1(filepath_var1.get(), progress_var, root, lambda new_fp: on_form_done(1, new_fp))
            elif form_number == 2:

                form2.process_form2(filepath_var2.get(), filepath_var1.get(), progress_var, root, lambda new_fp: on_form_done(2, new_fp))
            elif form_number == 3:
                form3.process_form3(filepath_var3.get(), filepath_var1.get(), progress_var, root, lambda new_fp: on_form_done(3, new_fp))
            elif form_number == 4:
                form4.process_form4(filepath_var4.get(), filepath_var1.get(), progress_var, root, lambda new_fp: on_form_done(4, new_fp))
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось обработать форму {form_number}. Ошибка: {e}")
            progress_var.set(f"Ошибка при обработке формы {form_number}.")
            progress_bar.set(0)
            root.update_idletasks()

    def create_summary():
        """Создание итогового сводного файла из обработанных форм."""
        try:
            forms_data = {2: filepath_var2.get(), 3: filepath_var3.get(), 4: filepath_var4.get()}
            output_file = filedialog.asksaveasfilename(filetypes=[("Excel Files", "*.xlsx")], defaultextension=".xlsx")
            if output_file:
                create_summary_from_memory(forms_data, output_file, filepath_var5, progress_var)
                messagebox.showinfo("Успешно", f"Итоговый файл сохранен: {output_file}")
                progress_bar.set(1)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при создании итогового файла: {e}")
            progress_var.set("Ошибка при создании итогового файла")
            progress_bar.set(0)

    # Создание основного окна
    root = ctk.CTk()
    root.title("Обработка форм")
    root.geometry("700x400")

    # Переменные для хранения путей к файлам
    filepath_var1 = ctk.StringVar()
    filepath_var2 = ctk.StringVar()
    filepath_var3 = ctk.StringVar()
    filepath_var4 = ctk.StringVar()
    filepath_var5 = ctk.StringVar()
    progress_var = ctk.StringVar(value="Готово к работе")

    # Метка заголовка
    title_label = ctk.CTkLabel(root, text="Выберите файлы:", font=("Arial", 16, "bold"))
    title_label.pack(pady=10)

    # Функция для добавления кнопки и поля ввода
    def create_file_input(label_text, var):
        frame = ctk.CTkFrame(root)
        frame.pack(pady=5, fill="x", padx=20)
        button = ctk.CTkButton(frame, text=label_text, command=lambda: choose_file(var), width=120)
        button.pack(side="left", padx=5)
        entry = ctk.CTkEntry(frame, textvariable=var, width=500)
        entry.pack(side="right", padx=5, fill="x", expand=True)

    create_file_input("Форма 1", filepath_var1)
    create_file_input("Форма 2", filepath_var2)
    create_file_input("Форма 3", filepath_var3)
    create_file_input("Форма 4", filepath_var4)
    create_file_input("Форма 5", filepath_var5)

    # Кнопки обработки форм
    button_frame = ctk.CTkFrame(root)
    button_frame.pack(pady=10, fill="x", padx=20)

    ctk.CTkButton(button_frame, text="Обработать форму 1", command=lambda: process_form(1)).pack(side="left", padx=5)
    ctk.CTkButton(button_frame, text="Обработать форму 2", command=lambda: process_form(2)).pack(side="left", padx=5)
    ctk.CTkButton(button_frame, text="Обработать форму 3", command=lambda: process_form(3)).pack(side="left", padx=5)
    ctk.CTkButton(button_frame, text="Обработать форму 4", command=lambda: process_form(4)).pack(side="left", padx=5)

    # Кнопка создания итогового файла
    summary_button = ctk.CTkButton(root, text="Создать итоговый файл", command=create_summary, fg_color="green")
    summary_button.pack(pady=10)

    # Прогресс-бар
    progress_bar = ctk.CTkProgressBar(root)
    progress_bar.pack(pady=10, fill="x", padx=20)
    progress_bar.set(0)  # Начальное значение

    # Метка состояния
    status_label = ctk.CTkLabel(root, textvariable=progress_var, font=("roboto", 12))
    status_label.pack(pady=5)


    root.mainloop()

if __name__ == "__main__":
    main()
