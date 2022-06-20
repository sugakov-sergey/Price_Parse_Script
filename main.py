from openpyxl import load_workbook, Workbook
from itertools import chain
import input

data, line = [], []
not_valid_values = []


def load_price(filename, sheetname):
    """ Загружаем входящий прайс для обработки """
    global sheet_in
    book_in = load_workbook(filename=filename, data_only=True)
    sheet_in = book_in[sheetname]


def info(sheet):
    """ Информация о данных в переданном листе """
    # Структура данных
    print('\n', '-' * 12, 'Структура исходных данных', '-' * 13, '\n')
    for row in sheet.iter_rows(min_row=0, max_row=6, max_col=5, values_only=True):
        print(row)
    print('...')
    print(sheet.max_row, 'строк')

    # Статистика
    print('\n', '-' * 20, 'Статистика', '-' * 20, '\n')
    if not_valid_values: print(len(not_valid_values), 'невалидных строк')

    # Список невалидных строк
    if not_valid_values:
        print('\n', '-' * 14, 'Список невалидных строк', '-' * 14, '\n')
        [print(i) for i in not_valid_values]


def make_book(articul, title, price, price_dis, type_dis):
    """ Создаем новый список строк значений в нужном порядке """
    global data, line
    cols = [articul, title, price, price_dis, type_dis]
    data = [[sheet_in.cell(row, col).value for col in cols] for row in range(1, sheet_in.max_row)]


def clear_data(data):
    global not_valid_values
    for row in data[:]:
        _clear_articul_cell(row)
        _clear_price_cell(row)
        _clear_price_dis_cell(row)
        _del_recurring_values(row)
        if not is_valid():
            not_valid_values += [row]
            data.remove(row)


def _clear_articul_cell(row):
    """ Форматирует ячейку с артикулом"""
    global articul_is_valid
    articul_is_valid = False
    if isinstance(row[0], str):
        row[0] = row[0].strip()
        articul_is_valid = True


def _clear_price_cell(row):
    """ Форматирует ячейку цены в число с двумя знаками после запятой"""
    global price_is_valid
    price_is_valid = False
    if isinstance(row[2], str):
        if set(row[2]) <= set('0123456789., '):
            row[2] = row[2].strip().replace(',', '.').replace(' ', '')
            row[2] = round(float(row[2]), 2)
            price_is_valid = True
    elif isinstance(row[2], float | int):
        row[2] = round(float(row[2]), 2)
        price_is_valid = True


def _clear_price_dis_cell(row):
    """ Форматирует ячейку цены со скидкой в число с двумя знаками после запятой.
        А так же при наличии скидки, прописывает тип скидки """
    if row[3]:
        if isinstance(row[3], str):
            if set(row[3]) <= set('0123456789., '):
                row[3] = row[3].strip().replace(',', '.').replace(' ', '')
                row[3] = round(float(row[3]), 2)
        elif isinstance(row[3], float | int):
            row[3] = round(float(row[3]), 2)
        row[4] = 'Акция'


def _del_recurring_values(row):
    pass


def is_valid():
    """ Валидация данных в необходимых для загрузки ячейках """
    return all((articul_is_valid, price_is_valid))


def write_data_to_file(filename, data):
    """ Запись обработанных данных в файл """
    book = Workbook()
    sheet = book.active
    _make_table_view(sheet)

    for row in chain(title_table, data):
        sheet.append(row)

    book.save(filename)


def _make_table_view(sheet):
    """ Настраиваем таблицу для вывода данных """
    global title_table
    title_table = [['Артикул', 'Наименование', 'Цена', 'Цена со скидкой', 'Тип скидки']]
    # изменяем ширину колонки
    sheet.column_dimensions['A'].width = 30  # Артикул
    sheet.column_dimensions['B'].width = 60  # Наименование
    sheet.column_dimensions['C'].width = 15  # Цена
    sheet.column_dimensions['D'].width = 20  # Цена со скидкой
    sheet.column_dimensions['E'].width = 20  # Тип скидки
    # Фиксируем строку заголовка
    sheet.freeze_panes = 'A2'


# Укажите расположение прайса и название рабочего листа
load_price(filename='aquabiz.xlsx',
           sheetname='ПРАЙС ABBER и GEMY')
# Укажите номера колонок с соответствующими данными в загружаемом документе
make_book(input.articul, input.title, input.price, input.price_dis, input.type_dis)
clear_data(data)
# Укажите название создаваемого прайса
write_data_to_file(input.filename, data=data)
info(sheet=sheet_in)
