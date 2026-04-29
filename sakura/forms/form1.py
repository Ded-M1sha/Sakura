import pandas as pd
import numpy as np
import os
from openpyxl import load_workbook
from tkinter import messagebox, Toplevel, Button, Label, simpledialog, ttk, BooleanVar, Checkbutton
import customtkinter as ctk

def choose_column(df, root):
    """
    Функция для выбора пользователем столбца аппроксимации.
    Открывает окно с кнопками, соответствующими именам столбцов.
    """
    selected_column = {"name": None}

    def select_column(col_name):
        selected_column["name"] = col_name
        choose_window.destroy()

    # Определение столбцов для выбора (исключая некоторые)
    columns = [col for col in df.columns if col not in ["Код товара", "Длина, см", "Ширина, см", "Высота, см"]]

    # Создаем новое окно для выбора столбца
    choose_window = ctk.CTkToplevel(root)
    choose_window.title("Выбор столбца для аппроксимации")
    ctk.CTkLabel(choose_window, text="Выберите столбец для аппроксимации:").pack(pady=10)

    # Добавляем кнопки для каждого столбца
    for col in columns:
        ctk.CTkButton(choose_window, text=col, command=lambda c=col: select_column(c)).pack(pady=10)

    # Ожидаем закрытия окна
    root.wait_window(choose_window)
    return selected_column["name"]

def get_upper_limit(root):
    """
    Функция для запроса верхней границы объема у пользователя.
    Возвращает введенное значение или None, если пользователь отменил ввод.
    """
    choose_window = ctk.CTkToplevel(root)
    choose_window.title("Ввод верхней границы объема")
    choose_window.geometry("600x200")

    label = ctk.CTkLabel(choose_window, text="Введите верхнюю границу объема (введите 0 для автоматического расчета):")
    label.pack(pady=10)

    entry = ctk.CTkEntry(choose_window, width=200)
    entry.pack(pady=5)

    upper_limit = {"value": None}

    def on_submit():
        """Сохранение введенного значения и закрытие окна"""
        try:
            upper_limit["value"] = float(entry.get())
        except ValueError:
            upper_limit["value"] = None  # Если введены некорректные данные
        choose_window.destroy()

    def on_cancel():
        """Закрытие окна без сохранения"""
        upper_limit["value"] = None
        choose_window.destroy()

    button_frame = ctk.CTkFrame(choose_window)
    button_frame.pack(pady=10)

    button_ok = ctk.CTkButton(button_frame, text="ОК", command=on_submit)
    button_ok.pack(side="left", padx=5)

    button_cancel = ctk.CTkButton(button_frame, text="Отмена", command=on_cancel)
    button_cancel.pack(side="left", padx=5)

    root.wait_window(choose_window)

    return upper_limit["value"]



def process_form1(filepath, progress_var, root, on_form1_done):
    new_filepath = os.path.join(os.path.dirname(filepath), "Форма 1_обработанная.xlsx")
    try:
        # Загрузка данных из файла
        df = pd.read_excel(filepath)
        df['Длина, см'] = pd.to_numeric(df['Длина, см'])
        df['Ширина, см'] = pd.to_numeric(df['Ширина, см'])
        df['Высота, см'] = pd.to_numeric(df['Высота, см'])

        # Вызываем окно "Качество данных" до аппроксимации данных
        def calculate_quality_metrics(df):

            quality_data = []
            total_rows = len(df) - 1  # Исключаем строку заголовка

            for col in df.columns:
                # Пустые значения
                empty_count = df[col].isnull().sum()
                empty_percentage = (empty_count / total_rows) * 100

                # Нули
                zero_count = (df[col] == 0).sum() if df[col].dtype in ['int64', 'float64'] else 0
                zero_percentage = (zero_count / total_rows) * 100

                # Константа
                unique_values = df[col].dropna().unique()
                is_constant = "Да" if len(unique_values) == 1 else "Нет"

                # Уникальный
                is_unique = "Да" if df[col].dropna().is_unique else "Нет"

                # Выбросы
                if df[col].dtype in ['int64', 'float64']:
                    Q1 = df[col].quantile(0.25)
                    Q3 = df[col].quantile(0.75)
                    IQR = Q3 - Q1
                    outliers = df[(df[col] > Q3 + 1.5 * IQR) | (df[col] < Q1 - 1.5 * IQR)]
                    outlier_count = len(outliers)
                else:
                    outlier_count = 0
                outlier_percentage = (outlier_count / total_rows) * 100

                # Отрицательные
                negative_count = (df[col] < 0).sum() if df[col].dtype in ['int64', 'float64'] else 0
                negative_percentage = (negative_count / total_rows) * 100

                # Пробелы в конце
                trailing_spaces_count = df[col].apply(
                    lambda x: isinstance(x, str) and x.rstrip() != x
                ).sum()
                trailing_spaces_percentage = (trailing_spaces_count / total_rows) * 100

                # Добавляем метрики в итоговый список
                quality_data.append([
                    col,
                    f"{empty_count} ({empty_percentage:.1f}%)",
                    f"{zero_count} ({zero_percentage:.1f}%)",
                    is_constant,
                    is_unique,
                    f"{outlier_count} ({outlier_percentage:.1f}%)",
                    f"{negative_count} ({negative_percentage:.1f}%)",
                    f"{trailing_spaces_count} ({trailing_spaces_percentage:.1f}%)",
                ])

            return quality_data, total_rows

        def countinue_without_changes():
            quality_window.destroy()
            # Добавляем расчет объема в м3 для единицы товара
            df['Объем единицы итоговый, м3'] = df['Длина, см'] * df['Ширина, см'] * df['Высота, см'] * 0.000001


            # Сохраняем результаты вычислений и исходные данные в новый Excel файл
            output_path = new_filepath
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                # Сохраняем основной обработанный лист
                df.to_excel(writer, index=False, sheet_name="Обработанные данные")

            # Завершение обработки формы
            on_form1_done(output_path)

        def improve_data_quality():
            """
            Улучшает качество данных. Открывает окно для выбора
            пользователем столбцов и проблем для обработки.
            """
            # Запрашиваем верхнюю границу объема у пользователя
            quality_window.destroy()
            upper_limit = get_upper_limit(root)
            if upper_limit is None:  # Если пользователь отменил ввод
                progress_var.set("Обработка формы 1 отменена.")
                root.quit()
                return

            # Создаем окно для выбора столбцов и проблем
            def show_selection_window():
                """Открывает окно для выбора столбцов и проблем."""
                selected_columns = []
                selected_problems = {
                    "remove_trailing_spaces": False,
                    "convert_negatives": False,
                    "handle_outliers": False,
                    "null_replace": False
                }

                def submit_selection():
                    """Фиксируем выбор пользователя."""
                    for idx, col_var in enumerate(column_vars):
                        if col_var.get():
                            selected_columns.append(df.columns[idx])

                    for problem, var in problem_vars.items():
                        selected_problems[problem] = var.get()

                    selection_window.destroy()

                # Создаем окно
                selection_window = ctk.CTkToplevel(root)
                selection_window.title("Выбор обработки качества данных")

                # Инструкции
                ctk.CTkLabel(selection_window, text="Выберите столбцы для обработки:", font=("Arial", 12)).pack(pady=5)

                # Чекбоксы для выбора столбцов
                column_vars = []
                for col in df.columns:
                    var = BooleanVar()
                    column_vars.append(var)
                    ctk.CTkCheckBox(selection_window, text=col, variable=var).pack(anchor="w", pady = 10)

                # Чекбоксы для выбора проблем
                ctk.CTkLabel(selection_window, text="Выберите проблемы для обработки:", font=("Arial", 12)).pack(pady = 10)

                problem_vars = {
                    "remove_trailing_spaces": ctk.BooleanVar(),
                    "convert_negatives": ctk.BooleanVar(),
                    "handle_outliers": ctk.BooleanVar(),
                    "null_replace": ctk.BooleanVar(),
                }

                ctk.CTkCheckBox(selection_window, text="Убрать пробелы в конце значений",
                            variable=problem_vars["remove_trailing_spaces"]).pack(anchor="w")
                ctk.CTkCheckBox(selection_window, text="Заменить отрицательные значения на модули",
                            variable=problem_vars["convert_negatives"]).pack(anchor="w")
                ctk.CTkCheckBox(selection_window, text="Обработать выбросы",
                            variable=problem_vars["handle_outliers"]).pack(anchor="w")
                ctk.CTkCheckBox(selection_window, text="Аппроксимировать нули",
                            variable=problem_vars["null_replace"]).pack(anchor="w")

                ctk.CTkButton(selection_window, text="Подтвердить выбор", command=submit_selection).pack(pady=10)

                root.wait_window(selection_window)
                return selected_columns, selected_problems

            # Запрашиваем у пользователя столбцы и проблемы для обработки
            columns_to_improve, problems_to_handle = show_selection_window()

            if not columns_to_improve:
                raise ValueError("Столбцы для обработки не выбраны.")
            df['Объем единицы, м3'] = df['Длина, см'] * df['Ширина, см'] * df['Высота, см'] * 0.000001
            # Очистка пробелов в конце значений
            if problems_to_handle["remove_trailing_spaces"]:

                for col in columns_to_improve:
                    if df[col].dtype == 'object':
                        df[col] = df[col].str.strip()

            # Замена отрицательных значений на модули
            if problems_to_handle["convert_negatives"]:


                for col in columns_to_improve:
                    if df[col].dtype in ['int64', 'float64']:
                        df[col] = df[col].apply(lambda x: abs(x) if x < 0 else x)

            # Обработка выбросов
            if problems_to_handle["handle_outliers"]:

                Q1 = df['Объем единицы, м3'].quantile(0.25)
                Q3 = df['Объем единицы, м3'].quantile(0.75)
                IQR = Q3 - Q1
                mean_approximated = df['Объем единицы, м3'].mean()
                median_approximated = df['Объем единицы, м3'].median()

                if upper_limit == 0:
                    Q6 = median_approximated if abs(
                        median_approximated - mean_approximated) / mean_approximated > 0.1 else mean_approximated
                else:
                    Q6 = upper_limit

                df['Является ли выбросом?'] = np.where(
                    (df['Объем единицы, м3'] < (Q1 - 1.5 * IQR)) |
                    (df['Объем единицы, м3'] > (Q3 + 1.5 * IQR)),
                    'Да', 'Нет'
                )

                df['Объем единицы после обработки выбросов, м3'] = np.where(
                    df['Является ли выбросом?'] == 'Да', Q6, df['Объем единицы, м3']
                )

            # Заполнение нулей
            if problems_to_handle["null_replace"]:

                column_to_approximate = choose_column(df, root)
                if not column_to_approximate:
                    raise ValueError("Столбец для аппроксимации не выбран.")

                df['Объем единицы, м3'] = df['Длина, см'] * df['Ширина, см'] * df['Высота, см'] * 0.000001

                unique_values = df[column_to_approximate].unique()
                avg_volume_by_group = {}
                global_mean = df['Объем единицы, м3'].mean(skipna=True)

                for value in unique_values:
                    group = df[df[column_to_approximate] == value]
                    avg_volume = group['Объем единицы, м3'].mean(skipna=True)
                    avg_volume_by_group[value] = avg_volume

                for value, avg_volume in avg_volume_by_group.items():
                    if np.isnan(avg_volume) or avg_volume == 0:
                        avg_volume_by_group[value] = global_mean

                df['Объем единицы после аппроксимации, м3'] = df.apply(
                    lambda row: row['Объем единицы, м3'] if pd.notnull(row['Объем единицы, м3']) and row[
                        'Объем единицы, м3'] != 0
                    else avg_volume_by_group.get(row[column_to_approximate], global_mean), axis=1
                )

                Q1 = df['Объем единицы после аппроксимации, м3'].quantile(0.25)
                Q3 = df['Объем единицы после аппроксимации, м3'].quantile(0.75)
                IQR = Q3 - Q1
                mean_approximated = df['Объем единицы после аппроксимации, м3'].mean()
                median_approximated = df['Объем единицы после аппроксимации, м3'].median()

                if upper_limit == 0:
                    Q6 = median_approximated if abs(
                        median_approximated - mean_approximated) / mean_approximated > 0.1 else mean_approximated
                else:
                    Q6 = upper_limit

                df['Является ли выбросом?'] = np.where(
                    (df['Объем единицы после аппроксимации, м3'] < (Q1 - 1.5 * IQR)) |
                    (df['Объем единицы после аппроксимации, м3'] > (Q3 + 1.5 * IQR)),
                    'Да', 'Нет'
                )

                df['Объем единицы после обработки выбросов, м3'] = np.where(
                    df['Является ли выбросом?'] == 'Да', Q6, df['Объем единицы после аппроксимации, м3']
                )
            if 'Объем единицы после аппроксимации, м3' in df.columns: df['Объем единицы итоговый, м3'] = df['Объем единицы после аппроксимации, м3']
            else:
                if 'Объем единицы после обработки выбросов, м3' in df.columns: df['Объем единицы итоговый, м3'] =  df['Объем единицы после обработки выбросов, м3']
                else: df['Объем единицы итоговый, м3'] = df['Объем единицы, м3']
            # Сохраняем результаты вычислений и исходные данные в новый Excel файл
            output_path = new_filepath
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                # Сохраняем основной обработанный лист
                df.to_excel(writer, index=False, sheet_name="Обработанные данные")

                # Создаем лист для данных аппроксимации
                df_approximation = pd.DataFrame({
                    "Уникальное значение": avg_volume_by_group.keys(),
                    "Средний объем, м3": avg_volume_by_group.values()
                })
                df_approximation.to_excel(writer, index=False, sheet_name="Данные аппроксимации")

                # Получаем доступ к книге и листу для добавления метрик
                workbook = writer.book
                sheet = workbook["Обработанные данные"]

                # Добавляем текстовые метки в ячейки P1-P6
                sheet["U1"] = "Q1"
                sheet["U2"] = "Q3"
                sheet["U3"] = "IQR"
                sheet["U4"] = "Среднее арифметическое после аппроксимации"
                sheet["U5"] = "Медиана после аппроксимации"
                sheet["U6"] = "Значение для замены"

                # Вписываем значения в ячейки Q1-Q6
                sheet["V1"] = Q1  # Q1
                sheet["V2"] = Q3  # Q3
                sheet["V3"] = IQR  # IQR
                sheet["V4"] = mean_approximated  # Среднее арифметическое
                sheet["V5"] = median_approximated  # Медиана
                sheet["V6"] = Q6  # Значение для замены
            on_form1_done(output_path)

        def show_quality_window():
            quality_data, total_rows = calculate_quality_metrics(df)

            # Создаем окно
            global quality_window
            quality_window = ctk.CTkToplevel(root)
            quality_window.title("Качество данных")

            ctk.CTkLabel(quality_window, text="Оценка качества данных", font=("Arial", 14, "bold")).pack(pady=10)

            # Создаем таблицу
            frame = ctk.CTkFrame(quality_window)
            frame.pack(fill="both", expand=True)

            # Заголовки столбцов
            headers = [
                "Столбец", "Пустые значения", "Нули", "Константа",
                "Уникальный", "Выбросы", "Отрицательные", "Пробелы в конце"
            ]

            # Размещение заголовков столбцов
            for col_num, header in enumerate(headers):
                label = ctk.CTkLabel(frame, text=header)
                label.grid(row=0, column=col_num, padx=10, pady=5)

            # Заполнение таблицы данными
            for row_num, row_data in enumerate(quality_data, start=1):
                for col_num, value in enumerate(row_data):
                    label = ctk.CTkLabel(frame, text=value)
                    label.grid(row=row_num, column=col_num, padx=10, pady=5)

            # Оценка качества данных
            total_issues = sum(int(row[1].split()[0]) for row in quality_data)
            data_quality_score = 100 - (total_issues / (total_rows * len(df.columns)) * 100)
            ctk.CTkLabel(
                quality_window,
                text=f"Оценка качества данных: {data_quality_score:.1f}/100",
                font=("Arial", 12, "bold")
            ).pack(pady=5)

            # Пояснение под таблицей
            explanation_text = (
                "Пустые значения: Количество строк с пустыми значениями и их доля в процентах.\n"
                "Нули: Количество строк с нулевыми значениями и их доля в процентах.\n"
                "Константа: 'Да', если все значения столбца одинаковы, 'Нет' — если есть хотя бы два уникальных значения.\n"
                "Уникальный: 'Да', если все значения уникальны, 'Нет' — если есть повторения.\n"
                "Выбросы: Количество строк с выбросами и их доля в процентах.\n"
                "Отрицательные: Количество строк с отрицательными значениями и их доля в процентах.\n"
                "Пробелы в конце: Количество строк с лишними отступами в конце значений и их доля в процентах."
            )
            ctk.CTkLabel(quality_window, text=explanation_text, justify="left", wraplength=600).pack(pady=10)

            # Кнопка закрытия окна
            ctk.CTkButton(quality_window, text="Обработать без изменения данных", command=countinue_without_changes).pack(pady=5)
            ctk.CTkButton(quality_window, text="Улучшить качество данных и обработать", command=improve_data_quality).pack(pady=5)


        # Показать окно качества данных
        show_quality_window()



    except Exception as e:
    # Вывод ошибки при возникновении исключения
        messge = "Произошла ошибка при обработке файла формы 1: " + str(e)
        show_error("Ошибка", messge)
        progress_var.set("Ошибка обработки.")


def show_error(title, message):
    error_window = ctk.CTkToplevel()
    error_window.title(title)
    error_window.geometry("400x200")

    label = ctk.CTkLabel(error_window, text=message, wraplength=350)
    label.pack(pady=10, padx=10)

    button = ctk.CTkButton(error_window, text="OK", command=error_window.destroy)
    button.pack(pady=10)

    error_window.grab_set()