from parameters.load_parameters import load_params
from handlers.table_handler import set_bu_values, convert_table_to_value
from handlers.table_by_algorythm import new_table_by_algorythm, save_data_to_table
from handlers.margin_handler import margin_handler


def start_handle(path_to_source: str, path_to_margin: str) -> None:
    """ Функция запускает обработку файла """

    # Считываем необходимые параметры в переменные из файла параметров
    divisions, del_items, rules, corrections = load_params('parameters/Параметры.xlsx')

    # Получаем таблицу с проставленными БЮ
    workbook, handled_table = set_bu_values(
        source_file=path_to_source,
        subdivisions=divisions,
        delete_items=del_items,
        handle_rules=rules,
        save_to_temp_file=True
    )

    # Получаем переходной формат таблицы
    workbook, table_headers, table_data = convert_table_to_value(workbook, handled_table, save_table_to_file=True)

    # Получаем значения, просчитанные по алгоритму из параметров
    result_dict = new_table_by_algorythm(path_to_params='parameters/Параметры.xlsx', headers=table_headers, table=table_data)

    # Данные из файла маржи (с корректировками)
    margin_dict = margin_handler(path_to_file=path_to_margin, corrections=corrections)

    # Сохраняем конечный результат таблицы
    save_data_to_table(workbook, result_dict, margin_dict, extended_version=False)
    save_data_to_table(workbook, result_dict, margin_dict, extended_version=True)

    print('\nОбработка файла успешно завершена')

if __name__ == '__main__':
    pass