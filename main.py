import pandas as pd
from openpyxl import Workbook
import re
import random
import os
from docx import Document
from docx.shared import Pt

# Поиск файла с иходными данными
def search_excel_file():
    print('Начинаю поиск таблицы с исходными данными...')
    files = os.listdir('./')
    for f in files:
        if '.xlsx' in f and f !='Результат.xlsx':
            return f
    return None

FILE_PATH = search_excel_file()

# Вычисление количества жил
def parse_expression(expr):
    expr = expr.replace(",", ".").replace("х", "x")
    numbers = re.findall(r"[\d.]+", expr)
    total = 1.0
    if len(numbers) > 2:
        numbers = numbers[:2]
    for num in numbers:
        total *= float(num)
    return int(round(total))

# Сохранение в Word
def make_word_result(result_list: list):
    doc = Document()
    last_place = ''

    table = doc.add_table(rows=1, cols=5)
    for result_element in result_list:
        value_title = list(result_element.keys())[0]
        place = result_element[value_title][0]
        values = result_element[value_title][1]

        # Заголовок блока (одна объединённая ячейка)
        if last_place == '':
            hdr_cells = table.rows[0].cells
            hdr_cells[0].merge(hdr_cells[-1])
            hdr_cells[0].text = place
            last_place = place
            hdr_cells[0].paragraphs[0].runs[0].font.size = Pt(12)
            row = table.add_row().cells
            merged_cell = row[0].merge(row[4])
            merged_cell.text = value_title
        elif last_place != '' and place != last_place:
            row = table.add_row().cells
            merged_cell = row[0].merge(row[4])
            merged_cell.text = place
            last_place = place
            row = table.add_row().cells
            merged_cell = row[0].merge(row[4])
            merged_cell.text = value_title
        else:
            row = table.add_row().cells
            merged_cell = row[0].merge(row[4])
            merged_cell.text = value_title

        for value in values:
            row = table.add_row().cells
            row[1].text = str(value[0])
            row[2].text = str(value[1])

            i = 1
            for vein in value[2]:
                row = table.add_row().cells
                row[0].text = str(i)
                row[1].text = f"Жила n×-{i}"
                row[3].text = str(vein)
                row[4].text = "Соответствует"
                i += 1

        # doc.add_paragraph()  # отступ между блоками

    doc.save("Результат.docx")

# Сохранение в Excel
def make_excel_result(result_list: list):
    wb = Workbook()
    ws = wb.active
    preview_place = ''
    for result_element in result_list:
        value_title = list(result_element.keys())[0]
        place = result_element[value_title][0]
        values = result_element[value_title][1]
        if place == preview_place:
            pass
        else:
            ws.append([place])
            preview_place = place
        ws.append([value_title])
        for value in values:
            ws.append(['', value[0], value[1]])
            i = 1
            for vein in value[2]:
                ws.append([i, f"Жила n×-{i}", '', vein, 'Соответствует'])
                i += 1
    wb.save('Результат.xlsx')


# Главная исполняющая функция
def main():
    try:
        df_dict = pd.read_excel(FILE_PATH, sheet_name=['Трассы', 'Помещения'])
        print('Таблица успешно загружена в программу!')
    except ValueError:
        return print('Таблица не имеет необходимых листов! Пожалуйста, убедитесь, что таблица подготовлена.\nПодсказка: должны быть листы "Трассы" и "Помещения"')
    except FileNotFoundError:
        return print('Программа не смогла найти подходящую таблицу! Убедитесь, что в папке с программой имеется файл с исходными данными.')
    
    df_data_sheet = df_dict['Трассы']
    df_headers_sheet = df_dict['Помещения']

    result_list = []

    search_list = df_headers_sheet['Конец'].to_list()
    for search_element in search_list:
        place = df_headers_sheet[df_headers_sheet['Конец'] == search_element]['Помещение'].to_list()[0]
        result_list.append({search_element: [place, []]})

    col_a = df_data_sheet.columns[0]  # A: Обозначение
    col_c = df_data_sheet.columns[2]  # C: Конец
    col_d = df_data_sheet.columns[3]  # D: Марка и т.п.
    col_e = df_data_sheet.columns[4]  # E: Сечение

    for result_element in result_list:
        value_title = list(result_element.keys())[0]
        for _, row in df_data_sheet.iterrows():
            if str(row[col_c]).strip() == value_title:
                value_a = str(row[col_a]).strip()
                value_d = str(row[col_d]).strip()
                expr_e = str(row[col_e]).strip()
                try:
                    count = parse_expression(expr_e)
                    rand_numbers = [random.randint(700, 900) for _ in range(count)]

                    result_element[value_title][1].append([value_a, value_d, rand_numbers])
                except Exception as e:
                    print(f"Ошибка при обработке строки: {value_d} — {e}")
    try:
        make_excel_result(result_list=result_list)
        make_word_result(result_list=result_list)
    except PermissionError:
        return print('Программа не смогла открыть файл, т. к. он уже открыт! Пожалуйста, закройте его.')
    print('Программа успешно обработала данные и вывела резульат в документы: Результат.xlsx')
    input('Нажмите Enter, чтобы завершить...')


if __name__ == "__main__":
    main()