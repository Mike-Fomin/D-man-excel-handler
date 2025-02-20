from handlers.start_handle import start_handle


def main() -> None:
    """ Точка входа в программу """

    # Путь до файла исходника в корне проекта ('C://Users/{User_name}/Downloads/Исходник.xlsx' - пример другой папки)
    path_to_source: str = 'Данные для производства 01.25.xlsx'

    # Путь до файла маржи в корне проекта ('C://Users/{User_name}/Downloads/Маржа.xlsx' - пример другой папки)
    path_to_margin: str = 'Маржа.xlsx'

    start_handle(path_to_source, path_to_margin)


if __name__ == '__main__':
    main()
