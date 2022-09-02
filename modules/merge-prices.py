import os
from openpyxl import load_workbook, Workbook

from start_app import _cell_alignment, _make_table_view

print("Внимание!\n"
      "Для правильного объедиения файлов прайсы должны быть обработаны \n"
      "и сохранены в папке 'input-for-merge'\n")

msg = 'Какой режим обработки прайса ?\n' \
      '1 - Минимальная цена при совпадении артикулов\n' \
      '2 - Приоритет цены по важности прайса\n\n' \
      'Введите номер варианта: '

merge_mode = input(msg)


def _get_files():
    """ Получем словарь обрабатываемых прайсов: ключ - имя файла, значение - путь файла """
    global files
    files = {}
    for file in os.listdir(r'.\input-for-merge'):
        path = os.path.abspath('input-for-merge\\' + file)
        files[file] = {'path': path}
    return files


def _get_priority():
    """ Запрашиваем приоритет цен в прайсах. Если приоритет прайса выше
        (значение больше), то цена перезаписывается. """
    msg = '\nВведите приоритет перезаписи прайсов. Если приоритет выше (значение больше),\n' \
          'то, при совпадении артикулов, цена перезапишется на цену с бОльшим значением.\n\n' \
          'Список файлов:\n'
    print(msg)
    for file in files.keys():
        print('+', file)
    print()
    for file in files.keys():
        files[file]['priority'] = int(input(f'Введите приоритет для файла "{file}": '))


def _write_data_to_file(filename, my_data):
    """ Запись обработанных данных в файл """
    book = Workbook()
    sheet = book.active
    for row in my_data:
        sheet.append(row)
    _make_table_view(sheet)
    _cell_alignment(sheet)
    book.save(filename)


def merge_prices(merge_mode):
    """ Объединяет прайсы в один исходя из выбранного режима """
    data_set = {'Артикул': ['Артикул', 'Наименование', 'Цена', 'Цена со скидкой', 'Тип скидки']}
    previous_priority = 0
    for file in files:
        book_in = load_workbook(filename=files[file]['path'], data_only=True)
        sheet_in = book_in.active
        current_priority = files[file]['priority']
        for current_row in sheet_in.values:
            if current_row[0] not in data_set.keys():
                data_set[current_row[0]] = list(current_row)
            else:
                if merge_mode == '1' and current_row[2] < data_set[current_row[0]][2]:
                    data_set[current_row[0]] = list(current_row)
                if merge_mode == '2' and current_priority > previous_priority:
                    data_set[current_row[0]] = list(current_row)
        previous_priority = current_priority

    new_data = list(data_set.values())
    _write_data_to_file('merged_price.xlsx', new_data)


_get_files()
if merge_mode == '2':
    _get_priority()
merge_prices(merge_mode)
