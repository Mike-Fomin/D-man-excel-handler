# Excel Handler

## Описание
Excel Handler — это Python-инструмент для автоматизированной обработки Excel-файлов, связанных с производственными данными и маржой фирмы заказчика. Проект фильтрует исходные данные, распределяет затраты по цехам на основе алгоритма, рассчитывает маржу с учётом корректировок и сохраняет результаты в двух версиях: базовой и расширенной.

## Основные возможности
- **Загрузка параметров** из файла `parameters/Параметры.xlsx` (подразделения, статьи, правила, корректировки, алгоритм).
- **Обработка исходных данных** (например, `Данные для производства 01.25.xlsx`):
  - Фильтрация строк по подразделениям и удаление ненужных статей.
  - Проставление значений "БЮ" на основе правил.
  - Суммирование числовых данных по "БЮ" (лист "Переходник").
- **Расчёт затрат по цехам** (Пекарский, Кондитерский, Слойки, ПФ) на основе листа "Алгоритм".
- **Обработка маржи** из файла `Маржа.xlsx`:
  - Извлечение данных о выпуске и FC по месяцам для каждого цеха.
  - Применение корректировок из параметров.
  - Подсчёт итоговых значений.
- **Сохранение результатов** в `results/Результат.xlsx`:
  - Лист "БЮ": отфильтрованная исходная таблица.
  - Лист "Переходник": промежуточные данные с суммами.
  - Лист "Первая": базовая таблица с маржой и расходами.
  - Лист "Расширенная": то же + проценты от выпуска.
    
## Требования
- Python 3.x
- Библиотеки:
  - `openpyxl` — для работы с Excel-файлами.
