import re
import os
import openpyxl
from pathlib import Path
from find_ervk import Compair

# main_file_path = '/Volumes/KINGSTON/Выверка ЕРВК/Псков/отсутствуют в ЕРВК на 31 октября _Псковская область.xlsx'
plan_file_without_empty_object_adress_cells = '/Volumes/KINGSTON/Выверка ЕРВК/Омск/Из плана 2023 - нет в ЕРВК.xlsx'

true_plan_without_empty_object_adress_cells = '/Volumes/KINGSTON/Выверка ЕРВК/Омск/2023050795_полный_план.xlsx'
def compair_with_11_files():
    """
    делает выборку в списке проверок по ОГРН на наличие таких ОГРН в одном из 11 переданных в ЕРВК файлов.
    файл "plan_file_without_empty_object_adress_cells" должен иметь формат выгрузки из ЕРКНМ
    :return:
    """


    files = [
        'РПТН_01.xlsx',
        'РПТН_02.xlsx',
        'РПТН_03.xlsx',
        'РПТН_04.xlsx',
        'РПТН_05.xlsx',
        'РПТН_06.xlsx',
        'РПТН_07.xlsx',
        'РПТН_08.xlsx',
        'РПТН_09.xlsx',
        'РПТН_10.xlsx',
        'РПТН_11.xlsx',
    ]


    for file in files:
        compared_file_path = f'/Volumes/KINGSTON/Выверка ЕРВК/Омск/{file}'
        Compair(
            main_file_path=plan_file_without_empty_object_adress_cells,
            compared_file_path=compared_file_path,
            startswith_row_main=3,
            startswith_row_compared=2,
            operative_column_main='C',    # I
            alter_operative_column_main='C',   # I
            operative_column_compare='Y',
            alter_operative_column_compare='AI',
            info_column_compare='B',
            record_column_main='T'    # AH
        )

    print('по всем файлам прошелся!')


def compair_with_plan_in_ekrnm():
    Compair(
        main_file_path=plan_file_without_empty_object_adress_cells,
        compared_file_path=true_plan_without_empty_object_adress_cells,
        startswith_row_main=2, # 19
        startswith_row_compared=19,
        operative_column_main='C',   # I
        alter_operative_column_main='C',  # I
        operative_column_compare='I',
        alter_operative_column_compare='I',
        info_column_compare='AD',
        record_column_main='U'   # AI
    )


def main():

    # compair_with_11_files()
    compair_with_plan_in_ekrnm()


if __name__ == '__main__':
    main()
