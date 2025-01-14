import pandas as pd
import numpy as np
from openpyxl import load_workbook
from tkinter import messagebox, Toplevel, Button, Label, simpledialog, ttk, BooleanVar, Checkbutton


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
    choose_window = Toplevel(root)
    choose_window.title("Выбор столбца для аппроксимации")
    Label(choose_window, text="Выберите столбец для аппроксимации:").pack()

    # Добавляем кнопки для каждого столбца
    for col in columns:
        Button(choose_window, text=col, command=lambda c=col: select_column(c)).pack()

    # Ожидаем закрытия окна
    root.wait_window(choose_window)
    return selected_column["name"]


def get_upper_limit():
    """
    Функция для запроса верхней границы объема у пользователя.
    Возвращает введенное значение или None, если пользователь отменил ввод.
    """
    upper_limit = simpledialog.askfloat("Ввод верхней границы объема",
                                        "Введите верхнюю границу объема (введите 0 для автоматического расчета):")
    if upper_limit is None:
        return None  # Если пользователь отменил ввод
    return upper_limit


def process_form1(filepath, progress_var, root, on_form1_done):
    try:
        # Загружаем данные из файла
        df = pd.read_excel(filepath)

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
            # Добавляем расчет объема в м3 для единицы товара
            df['Объем единицы, м3'] = df['Длина, см'] * df['Ширина, см'] * df['Высота, см'] * 0.000001


            # Сохраняем результаты вычислений и исходные данные в новый Excel файл
            output_path = "Форма 1 обработанная.xlsx"
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
            upper_limit = get_upper_limit()
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
                selection_window = Toplevel(root)
                selection_window.title("Выбор обработки качества данных")

                # Инструкции
                Label(selection_window, text="Выберите столбцы для обработки:", font=("Arial", 12)).pack(pady=5)

                # Чекбоксы для выбора столбцов
                column_vars = []
                for col in df.columns:
                    var = BooleanVar()
                    column_vars.append(var)
                    Checkbutton(selection_window, text=col, variable=var).pack(anchor="w")

                # Чекбоксы для выбора проблем
                Label(selection_window, text="Выберите проблемы для обработки:", font=("Arial", 12)).pack(pady=5)

                problem_vars = {
                    "remove_trailing_spaces": BooleanVar(),
                    "convert_negatives": BooleanVar(),
                    "handle_outliers": BooleanVar(),
                    "null_replace": BooleanVar(),
                }

                Checkbutton(selection_window, text="Убрать пробелы в конце значений",
                            variable=problem_vars["remove_trailing_spaces"]).pack(anchor="w")
                Checkbutton(selection_window, text="Заменить отрицательные значения на модули",
                            variable=problem_vars["convert_negatives"]).pack(anchor="w")
                Checkbutton(selection_window, text="Обработать выбросы",
                            variable=problem_vars["handle_outliers"]).pack(anchor="w")
                Checkbutton(selection_window, text="Аппроксимировать нули",
                            variable=problem_vars["null_replace"]).pack(anchor="w")

                Button(selection_window, text="Подтвердить выбор", command=submit_selection).pack(pady=10)

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
                    if np.isnan(avg_volume):
                        avg_volume_by_group[value] = global_mean

                df['Объем единицы после аппроксимации, м3'] = df.apply(
                    lambda row: row['Объем единицы, м3'] if pd.notnull(row['Объем единицы, м3'])
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
                #Удаление дубликатов
                df.drop_duplicates(keep='first')
                # Сохраняем результаты вычислений и исходные данные в новый Excel файл
                output_path = "Форма 1 обработанная.xlsx"
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
                    sheet["P1"] = "Q1"
                    sheet["P2"] = "Q3"
                    sheet["P3"] = "IQR"
                    sheet["P4"] = "Среднее арифметическое после аппроксимации"
                    sheet["P5"] = "Медиана после аппроксимации"
                    sheet["P6"] = "Значение для замены"

                    # Вписываем значения в ячейки Q1-Q6
                    sheet["Q1"] = Q1  # Q1
                    sheet["Q2"] = Q3  # Q3
                    sheet["Q3"] = IQR  # IQR
                    sheet["Q4"] = mean_approximated  # Среднее арифметическое
                    sheet["Q5"] = median_approximated  # Медиана
                    sheet["Q6"] = Q6  # Значение для замены
            on_form1_done(output_path)
            show_quality_window()

        def show_quality_window():
            quality_data, total_rows = calculate_quality_metrics(df)

            # Создаем окно
            quality_window = Toplevel(root)
            quality_window.title("Качество данных")

            Label(quality_window, text="Оценка качества данных", font=("Arial", 14, "bold")).pack(pady=10)

            # Создаем таблицу
            frame = ttk.Frame(quality_window)
            frame.pack(fill="both", expand=True)
            tree = ttk.Treeview(frame, columns=[
                "Столбец", "Пустые значения", "Нули", "Константа",
                "Уникальный", "Выбросы", "Отрицательные", "Пробелы в конце"
            ], show="headings")

            # Определяем заголовки
            headers = [
                "Столбец", "Пустые значения", "Нули", "Константа",
                "Уникальный", "Выбросы", "Отрицательные", "Пробелы в конце"
            ]
            for col in headers:
                tree.heading(col, text=col)
                tree.column(col, anchor="center", width=150)

            # Заполняем таблицу
            for row in quality_data:
                tree.insert("", "end", values=row)

            tree.pack(fill="both", expand=True)

            # Оценка качества данных
            duplicate_count = df.duplicated(keep=False).sum()
            total_issues = sum(int(row[1].split()[0]) for row in quality_data)
            data_quality_score = 100 - ((total_issues / (total_rows * len(df.columns)) * 100)*(duplicate_count/total_rows))
            Label(
                quality_window,

                text=f"Количество строк-дубликатов: {duplicate_count}\n\nОценка качества данных: {data_quality_score:.1f}/100",
                font=("Arial", 12, "bold")
            ).pack(pady=5)

            # Пояснение под таблицей
            explanation_text = (
                " Пустые значения: Количество строк с пустыми значениями и их доля в процентах.\n"
                "Нули: Количество строк с нулевыми значениями и их доля в процентах.\n"
                "Константа: 'Да', если все значения столбца одинаковы, 'Нет' — если есть хотя бы два уникальных значения.\n"
                "Уникальный: 'Да', если все значения уникальны, 'Нет' — если есть повторения.\n"
                "Выбросы: Количество строк с выбросами и их доля в процентах.\n"
                "Отрицательные: Количество строк с отрицательными значениями и их доля в процентах.\n"
                "Пробелы в конце: Количество строк с лишними отступами в конце значений и их доля в процентах.\n"
                "Количество строк-дубликатов: количество полностью одинаковых строк"
            )
            Label(quality_window, text=explanation_text, justify="left", wraplength=600).pack(pady=10)
            print(df.duplicated(keep=False).sum())
            # Кнопка закрытия окна
            ttk.Button(quality_window, text="Обработать без изменения данных", command=countinue_without_changes).pack(pady=5)
            ttk.Button(quality_window, text="Улучшить качество данных и обработать", command=improve_data_quality).pack(pady=5)

        # Показать окно качества данных
        show_quality_window()



    except Exception as e:
        # Вывод ошибки при возникновении исключения
        messagebox.showerror("Ошибка", f"Произошла ошибка при обработке файла формы 1: {e}")
        progress_var.set("Ошибка обработки.")