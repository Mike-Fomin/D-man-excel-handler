import openpyxl
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from parameters.load_parameters import load_params
from table_handler import set_bu_values, convert_table_to_value


def main() -> None:
    # Считываем необходимые параметры в переменные из файла параметров
    divisions, del_items, rules, corrections = load_params('Параметры.xlsx')

    # Получаем таблицу с проставленными БЮ
    handled_table: list[list] = set_bu_values(
        source_file='Исходник.xlsx',
        subdivisions=divisions,
        delete_items=del_items,
        handle_rules=rules,
        save_to_temp_file=False
    )

    # Получаем переходной формат таблицы
    headers, table = convert_table_to_value(handled_table, save_table_to_file=False)
    print(headers)
    print(*table, sep='\n')



if __name__ == '__main__':
    main()
