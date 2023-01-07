import csv
from re import sub
import os
# from typing import List
# import matplotlib.pyplot as plt
# import numpy as np
# import pdfkit
# import doctest
# from jinja2 import Environment, FileSystemLoader
# from openpyxl import Workbook
# from openpyxl.styles import Font, Border, Side
# from openpyxl.styles.numbers import FORMAT_PERCENTAGE_00


def custom_quit(msg: str) -> None:
    """
    Выход из программы с выводом сообщения на консоль.

    :param msg: Сообщение, выводимое на консоль.
    """

    print(msg)
    quit()


class Translator:
    """
    Класс для перевода валюты из международного формата на русский язык и с русского языка в числовой формат по
        заранее установленному курсу.

    """

    AZN: str = "Манаты"
    BYR: str = "Белорусские рубли"
    EUR: str = "Евро"
    GEL: str = "Грузинский лари"
    KGS: str = "Киргизский сом"
    KZT: str = "Тенге"
    RUR: str = "Рубли"
    UAH: str = "Гривны"
    USD: str = "Доллары"
    UZS: str = "Узбекский сум"

    def translate(self, key: str, dict_name: str = None) -> str:
        """
        Переводит из международного формата на русский язык. Если было передано имя словаря, возвращает из него
        значение по ключу.

        :param key: Международный формат валюты или название словаря.
        :param dict_name: Имя словаря, существующего в аттрибутах класса (на данный момент нет доступных).

        >>> t = Translator()
        >>> t.translate('USD')
        'Доллары'
        >>> t.RUR
        'Рубли'
        """
        if dict_name is not None:
            return self.__getattribute__(dict_name)[key]
        return self.__getattribute__(key)


class UserInterface:
    """
    Класс обработки ввода пользовательских данных.

    Attributes
    ----------
    file_name : str
        Путь до CSV-файла.
    profession_name : str
        Название профессии, введённое пользователем.
    """

    file_name: str
    profession_name: str

    def __init__(self, file_name: str = None, profession_name: str = None):
        """
        Инициализирует объект UserInterface, принимает название CSV-файла.

        :param file_name: Путь до CSV-файла. По-умолчанию '../vacancies_medium.csv'.
        :param profession_name: Название профессии для сбора статистики. По-умолчанию 'Программист'.

        >>> u = UserInterface()
        >>> u.f_name
        '../vacancies_medium.csv'
        >>> u.profession_name
        'Программист'
        >>> i = UserInterface('example.csv')
        >>> i.f_name
        'example.csv'
        >>> x = UserInterface(profession_name='Аналитик')
        >>> x.profession_name
        'Аналитик'
        """
        if file_name is not None:
            self.file_name = file_name
        else:
            self.file_name = "../vacancies_medium.csv"
        if profession_name is not None:
            self.profession_name = profession_name
        else:
            self.profession_name = 'Программист'


class CSV:
    """Класс для чтения и обработки CSV-файлов.

    Attributes
    ----------
    data : csv.reader
        Данные, полученные после прочтения CSV-файла при помощи функции reader библиотеки csv.
    title : list
        Список заголовков столбцов CSV-файла.
    rows : list
        Список строк с данными о вакансии. 1 строка = 1 вакансия.
    """

    data: csv.reader
    title: list
    rows: list

    def __init__(self, file_name: str):
        """
        Инициализирует объект CSV, пытается прочесть файл с переданным именем. Обрабатывает случаи пустого файла и
        отсутствия данных в файле.

        :param file_name: Путь до CSV-файла.

        """
        with open(file_name, 'r', newline='', encoding='utf-8-sig') as file:
            self.data = csv.reader(file)
            try:
                self.title = next(self.data)
            except StopIteration:
                custom_quit('Пустой файл')

            self.rows = [row for row in self.data
                         if len(list(filter(lambda word: word != '', row))) == len(self.title)]

            if len(self.rows) == 0:
                custom_quit('Нет данных')


class Salary:
    """
    Класс для предоставления зарплаты.

    Attributes
    ----------
    salary_from : int
        Нижняя граница вилки оклада.

    salary_to : int
        Верхняя граница вилки оклада.

    salary_currency : str
        Валюта оклада на русском языке.

    translator : Translator
        Переводчик валюты из международного формата на русский язык.
    """

    salary_from: int or float
    salary_to: int or float
    salary_currency: str
    currency_to_rub: {str, float} = {
        "Манаты": 35.68,
        "Белорусские рубли": 23.91,
        "Евро": 59.9,
        "Грузинский лари": 21.74,
        "Киргизский сом": 0.76,
        "Тенге": 0.13,
        "Рубли": 1,
        "Гривны": 1.64,
        "Доллары": 60.66,
        "Узбекский сум": 0.0055,
    }
    translator: Translator

    def __init__(self, salary_from: int or float = None, salary_to: int or float = None, salary_currency: str = None):
        """
        Инициализирует объект Salary, выполняет конвертацию для целочисленных полей.

        :param salary_from: Нижняя граница вилки оклада.
        :param salary_to: Верхняя граница вилки оклада.
        :param salary_currency: Валюта оклада на русском языке во множественном числе.
        """

        self.translator = Translator()
        if salary_from is not None:
            self.salary_from = salary_from
        if salary_to is not None:
            self.salary_to = salary_to
        if salary_currency is not None:
            self.salary_currency = salary_currency

    def set_field(self, key: str, value: str) -> None:
        """
        Устанавливает поле зарплаты, значение по ключу.

        :param key: Название поля.
        :param value: Значение поля. Валюта на русском языке во множественном числе. Числовые значения приводятся к int.
        """

        if key in ['salary_from', 'salary_to']:
            value = float(value)
        if key == 'salary_currency':
            value = self.translator.translate(value)
        self.__setattr__(key, value)

    def get_average_in_rur(self) -> int:
        """
        Вычисляет среднюю зарплату из вилки и переводит в рубли при помощи словаря - currency_to_rub.

        :returns: Средняя зарплата в рублях.
        """
        return int(self.currency_to_rub[self.salary_currency] * (self.salary_from + self.salary_to) // 2)


class Vacancy:
    """
    Класс вакансии используется для обработки данных о вакансиях из CSV-файлов.

    Attributes
    ----------
    name : str
        Название вакансии
    salary : Salary
        Вилка и валюта оклада
    area_name : str
        Название населённого пункта
    published_at : int
        Время публикации в формате - год.
    """

    name: str
    salary: Salary
    area_name: str
    published_at: int

    def __init__(self, fields: dict):
        """
        Инициализирует класс вакансии, используя переданные поля.

        :param fields: Словарь с полями вакансии. Доступные ключи - name, salary_from, salary_to, salary_currency,
        area_name, published_at

        >>> v = Vacancy({"name": 'Программист'})
        >>> v.t
        'Программист'
        >>> hasattr(v, 'area_name')
        False
        """
        for key, value in fields.items():
            if not self.check_salary(key, value):
                self.__setattr__(key, self.get_correct_field(key, value))

    def get_field(self, field: str) -> int or str:
        """
        Возвращает значение поля вакансии по ключу.

        :param field: Название поля вакансии.
        >>> v = Vacancy({'name': 'Аналитик'})
        >>> v.get_field('name')
        'Аналитик'
        """
        if field in 'salary':
            return self.salary.get_average_in_rur()
        return self.__getattribute__(field)

    def check_salary(self, key: str, value: str) -> bool:
        """
        Проверяет и устанавливает поле Salary, если его ещё нет

        :param key: Название поля зарплаты, такое как salary_from, salary_to, salary_currency.
        :param value: Значение поля зарплаты, числовое значение или международный формат валюты.
        :returns: Если название поля относится к зарплате, возвращает True и создаёт объект Salary, если его не было.

        """
        is_salary = False
        if key in ['salary_from', 'salary_to', 'salary_currency']:
            if not hasattr(self, 'salary'):
                self.salary = Salary()
            self.salary.set_field(key, value)
            is_salary = True
        return is_salary

    @staticmethod
    def get_correct_field(key: str, value: str or list) -> int or str:
        """
        Возвращает отформатированное поле вакансии. Сейчас форматирует только поле published_at.

        :param key: Название поля вакансии.
        :param value: Значение поля вакансии. Дату в формате YY-MM-DDTHH:MM:SS+MS преобразует в год в числовом формате.

        >>> v = Vacancy({'published_at': '2022-12-12T16:23:11+03'})
        >>> v.get_field('published_at')
        2022
        """

        if key == 'published_at':
            big, small = value[:19].split('T')
            year, month, day = big.split('-')
            return int(year)
        else:
            return value




def parse_html(line: str) -> str:
    """
    Убирает HTML-теги из строки.

    :param line: Строка для обработки.
    :returns: Возвращает строку без HTML-тегов.
    """
    line = sub('<.*?>', '', line)
    res = [' '.join(word.split()) for word in line.replace("\r\n", "\n").split('\n')]
    return res[0] if len(res) == 1 else res  # Спасибо Яндекс.Контесту за еще один костыль!


def parse_row_vacancy(header: list, row_vacs: list) -> dict:
    """
    Очищает строки от HTML-тегов и разбивает её на данные для вакансии.

    :param header: список заголовков из CSV-файла.
    :param row_vacs: список строк, прочитанных из CSV-файла.
    """
    return dict(zip(header, map(parse_html, row_vacs)))


def is_year_presented(dictionaries: list, year: str) -> {str: str or list} or None:
    """
    Определяет, существует ли словарь в списке словарей, в ключах которого есть переданный год.
    :param dictionaries: Список словарей, в которых проводится поиск.
    :param year: Год, который ищется среди ключей словарей.
    """
    for dictionary in dictionaries:
        if year in dictionary.keys():
            return dictionary
    return None


def get_vacancies_by_years(vacs_fields_dicts: list) -> list:
    """
    Обрабатывает список словарей с полями вакансии, возвращает список вакансий по годам.
    :param vacs_fields_dicts: Список словарей с полями вакансии.
    :returns: Список словарей типа {год: {поле вакансии: значение}}
    """
    result = []
    for vac_fields_dict in vacs_fields_dicts:
        vac_year = vac_fields_dict['published_at'][:4]
        existing_year_dict = is_year_presented(result, vac_year)
        if existing_year_dict is None:
            result.append({vac_year: [vac_fields_dict]})
        else:
            existing_year_dict[vac_year].append(vac_fields_dict)
    return result


def generate_csv_vacancies(year_vacancies: dict, year: str = None, path: str = None) -> None:
    """
    Создаёт CSV-файл с вакансиями, сгруппированными по году.

    :param year_vacancies: Словарь типа {год: список вакансий типа {поле: значение}}
    :param year: Год, который представлен в переданных вакансиях. Необязательный параметр.
    :param path: Путь, по которому будет сохранён CSV-файл (без слэша в конце). По-умолчанию папка "csvs_by_years"
    """
    if year is None:
        year = list(year_vacancies.keys())[0]
    if path is None:
        path = "csvs_by_years"

    fieldnames = list(year_vacancies[year][0].keys())
    file_name = f"{path}/vacancies_by_{year}.csv"
    with open(file_name, 'w', newline='', encoding='utf-8-sig') as file:
        writer = csv.DictWriter(file, fieldnames)
        writer.writeheader()

        for vac in year_vacancies[year]:
            writer.writerow(vac)


def generate_csvs_by_years(vacs_by_years_dicts: list) -> None:
    """
    Создаёт CSV-файлы вакансий, сгруппированных по годам. Один год - один файл.

    :param vacs_by_years_dicts: Список словарей типа {год: список вакансий типа {поле: значение}}
    """
    csvs_directory_name = "csvs_by_years"
    if not any(map(lambda name: name == csvs_directory_name, os.listdir(os.curdir))):
        os.mkdir(csvs_directory_name)
    for vacs in vacs_by_years_dicts:
        generate_csv_vacancies(vacs, list(vacs.keys())[0], csvs_directory_name)


if __name__ == '__main__':
    ui = UserInterface()
    csv_data = CSV(ui.file_name)
    title, row_vacancies = csv_data.title, csv_data.rows
    vacancies_fields_dictionaries = [parse_row_vacancy(title, row_vac) for row_vac in row_vacancies]
    vacancies_by_years_dictionaries = get_vacancies_by_years(vacancies_fields_dictionaries)
    generate_csvs_by_years(vacancies_by_years_dictionaries)