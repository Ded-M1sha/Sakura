import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import load_workbook
import re  # Для извлечения числа из строки


def read_data_from_excel(file_path, sheet_name, date_col, start_row, end_row):
    wb = load_workbook(file_path)
    sheet = wb[sheet_name]

    dates = []
    values = []
    table_titles = []  # Список для хранения названий таблиц
    table_start_row = None  # Строка с "Приведенная дата"
    columns_names = []  # Список названий параметров

    # Чтение данных и поиск блоков с датами
    for row in range(start_row, sheet.max_row + 1):
        cell_value = sheet.cell(row=row, column=date_col).value

        # Ищем строку с "Приведенная дата"
        if cell_value == "Приведенная дата":
            if table_start_row is not None:
                # Прочитаем все значения между двумя блоками с датами
                table_data = process_table(sheet, table_start_row + 1, row - 1, date_col, columns_names)
                dates.append(table_data['Date'])
                values.append(table_data.drop(columns=["Date"]))

            # Извлечем название таблицы из строки перед "Приведенная дата"
            title_row = row - 1
            title_text = sheet.cell(row=title_row, column=date_col).value
            table_titles.append(extract_table_title(title_text))

            table_start_row = row  # Начало нового блока с датами

            # Запоминаем названия параметров (игнорируем пустые столбцы)
            columns_names = []
            for col in range(date_col + 1, sheet.max_column + 1):
                param_name = sheet.cell(row=row, column=col).value
                if param_name:  # Если значение не пустое
                    columns_names.append(param_name)

    # Если после последней строки с датами есть данные
    if table_start_row is not None:
        table_data = process_table(sheet, table_start_row + 1, sheet.max_row, date_col, columns_names)
        dates.append(table_data['Date'])
        values.append(table_data.drop(columns=["Date"]))

    return dates, values, table_titles


def extract_table_title(title_text):
    """
    Извлекает название таблицы на основе строки перед "Приведенная дата".
    """
    print(f"Обрабатываемая строка: {title_text}")  # Вывод строки для проверки
    # Поиск текста в формате "форма Х_обработанная", где Х — цифра
    match = re.search(r"форма (\d+)_обработанная", title_text, re.IGNORECASE)
    if match:
        form_number = int(match.group(1))  # Извлекаем число Х
        if form_number == 2:
            return "Хранение"
        elif form_number == 3:
            return "Входящий поток"
        elif form_number == 4:
            return "Исходящий поток"
    return "Неизвестная форма"


def process_table(sheet, start_row, end_row, date_col, columns_names):
    # Чтение и обработка данных для каждого блока таблицы
    dates = [sheet.cell(row=row, column=date_col).value for row in range(start_row, end_row + 1)]
    # Преобразуем даты в datetime, явным образом указывая, что день идет первым
    dates = pd.to_datetime(dates, errors='coerce',
                           dayfirst=True)  # Используем dayfirst=True, чтобы гарантировать правильный формат

    # Получаем все данные по столбцам (значения показателей)
    values = []
    for col in range(date_col + 1, sheet.max_column + 1):
        col_values = [sheet.cell(row=row, column=col).value for row in range(start_row, end_row + 1)]
        # Добавляем только непустые столбцы
        if any(val is not None for val in col_values):
            values.append(col_values)

    # Создаем DataFrame
    df = pd.DataFrame(values).T  # Транспонируем, чтобы строки стали столбцами
    df.columns = columns_names  # Используем список названий параметров
    df['Date'] = dates

    # Удаляем строки с пустыми датами
    df = df.dropna(subset=['Date'])

    # Сортируем по дате, чтобы она была в хронологическом порядке
    df = df.sort_values(by='Date')

    return df


def plot_data(dates, values, table_titles, file_path):
    import os

    output_dir = os.path.dirname(file_path)  # Папка исходного файла

    # Для каждой таблицы строим графики
    for i, (date_data, value_data, table_title) in enumerate(zip(dates, values, table_titles)):
        for column in value_data.columns:
            plt.figure(figsize=(10, 6))  # Размер графика

            plt.plot(date_data, value_data[column], marker='o', linestyle='-', color='#e80b16', label=column)

            # Название графика с измененным шрифтом
            plt.title(f'{table_title}. {column}', fontdict={'fontsize': 14, 'fontweight': 'bold', 'family': 'Arial'})
            plt.xlabel('Дата')  # Ось X - дата
            plt.ylabel('Значение')  # Ось Y - значение
            plt.xticks(date_data, rotation=45)  # Поворот подписей на оси X
            plt.grid(True)
            plt.legend()
            plt.tight_layout()

            # Название окна
            plt.gcf().canvas.manager.set_window_title(f'{table_title} - {column}')

            # Сохранение графика
            output_path = os.path.join(output_dir, f'{table_title}_{column}.png')
            plt.savefig(output_path, format='png', dpi=300)

            plt.show()


