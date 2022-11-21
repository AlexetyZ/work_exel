import os
import re

from sql import Database as db
from pxl import Exel_work
import openpyxl


def request_inspection_number_exists(find):
    result = db().get_terr_upravlenie_name(condition=find)
    if result != ():
        if find in result[0][0]:
            return 0
    else:
        return 1


def get_list_from_sh_column(*columns: str, wb_path: str, start_from_row: int = 1, reference_column: str = 'A',
                            del_last_empty_rows=False):
    """
    Если нужно получить значения определенного столбца в файле exel. Предварительно отрезуются пустые строки снизу.


    :param wb_path: путь до файла, в котором итерируются ячейки
    :param column: столбец, в котором итерируются ячейки
    :param start_from_row: начало итерации для удаления пустых строк
    :param reference_column: показательный столбец, по которому отсчитывается количество строк, то есть,
        не пустое значение ячейки этого столбца гарантирует, что вся строка подлежит отработке. По умолчанию "А"
    :return: Возвращает список значений столбца в ячейке exel
    """
    list = []
    if del_last_empty_rows is True:
        Exel_work().delete_last_empty_rows(wb_path, start_from_row)
    wb = openpyxl.load_workbook(wb_path)
    sh = wb.worksheets[0]
    for n, row in enumerate(sh[reference_column]):
        if n + 1 < start_from_row:
            continue
        corteg = []
        for column in columns:
            corteg.append(sh[f'{column}{n + 1}'].value)
        list.append(tuple(corteg))
    return list


def insert_terr_upravlenie_from_wb():
    """
    Заранее измени в списке планов название ТУ и даты плана на "2023" или какой там год у тебя.
    :return:
    """
    list = get_list_from_sh_column(
        'C',
        wb_path='/Users/aleksejzajcev/Desktop/Список_планов_КНМ.xlsx',
        start_from_row=4,
    )
    for name in list:
        if request_inspection_number_exists(name[0]) == 0:
            continue
        else:
            db().insert_terr_uprav(name[0])


def insert_plan_proverok_from_wb():
    list = get_list_from_sh_column(
        'C', 'A', 'D', 'G', 'F',
        wb_path='/Users/aleksejzajcev/Desktop/Список_планов_КНМ.xlsx',
        start_from_row=4
    )

    for row in list:
        ter_upr = db().get_terr_upravlenie_id(row[0])
        number = row[1]
        year = row[2]
        count = row[3]
        status = row[4]

        db().insert_plan_proverok(ter_upr, number, year, count, status)


def rename_terr_upravlenie(before, after):
    list = get_list_from_sh_column(
        'C',
        wb_path='/Users/aleksejzajcev/Desktop/Список_планов_КНМ.xlsx',
        start_from_row=1
    )
    for name in list:

        try:
            id = db().get_terr_upravlenie_id(name[0])

            new_name = str(name[0]).replace(before, after)
            print(new_name)

            db().change_names_knd_terr_upravlenie(
                new_name=new_name,
                terr_id=id
            )
        except:
            pass


def insert_inspections_from_wb():
    """
    Занесение данных о субъетах, объектах, проверках и их отношениях в базу данных MySQL для использования в приложении
    ДЖАНГО.
    В основу берется файл выгрузки плана из ЕРКНМ.
    Требования к файлу: 1. НЕ ДОЛЖНО БЫТЬ "'", никаких одиночных кавычек.
                        2. Обращать внимание на размещение информации в столбцах, если есть изменения, внести запрос
                        из таблицы в соответствии с функцией get_list_from_sh_column() ниже, а так же на окончание
                        шапки таблицы и начало непосредственной информации (в данный момент строка 19).
    :return: По окончании функции возвращается 0 в случае успеха, занесенные данные можно посмотреть в АДМИНЕ приложения
    """
    risk_table = {
        'Первый': 'чрезвычайно высокий риск',
        'Второй': 'высокий риск',
        'Третий': 'значительный риск',
        'Четвертый': 'средний риск',
        'Пятый': 'умеренный риск',
        'Шестой': 'низкий риск',
    }


    def replace_values(row_table):
        new_row = []
        for value in row_table:
            try:
                new_row.append(value.replace("'", ""))
            except:
                new_row.append(value)
        row_table = tuple(new_row)
        return row_table

    last_obj_risk_5 = None
    last_obj_risk_8 = None
    dir = '/Users/aleksejzajcev/Desktop/выгруженные планы на 2023'
    for file in os.listdir(dir):
    # file = '2023049566_полный_план.xlsx'
        print(file)
        to_pass = ['.DS_Store', '._Чукотка.xlsx']
        if file in to_pass or re.search('\._', file):
            continue
        plan_path = f'{dir}/{file}'
        plan_number = re.split('_', file)[0]
        list = get_list_from_sh_column(
            'B', 'E', 'J', 'I', 'F', 'AB', 'W', 'AF', 'AC', 'AD',
            wb_path=plan_path,
            start_from_row=19
        )
        subj_id = None
        obj_id = None
        insp_id = None

        for row in list:


            if row[0] is not None:
                print(row)
                row = replace_values(row)
                subj_name = row[0]

                if row[1] is not None:
                    subj_address = row[1]
                else:
                    subj_address = risk_table[row[9]]
                inn = row[2]
                ogrn = row[3]
                subj_id = db().insert_subject_with_return_id(
                    name=subj_name,
                    address=subj_address,
                    inn=inn,
                    ogrn=ogrn
                )
                # print(subj_id)
                obj_kind = row[4]
                if row[1] is not None:
                    obj_address = row[1]
                else:
                    obj_address = risk_table[row[9]]

                if row[5] is not None:
                    obj_risk = row[5]
                    last_obj_risk_5 = row[5]
                else:
                    if row[8] is not None:
                        obj_risk = risk_table[row[8]]
                        last_obj_risk_8 = risk_table[row[8]]
                    else:
                        if last_obj_risk_5 is not None:
                            obj_risk = last_obj_risk_5
                        else:
                            obj_risk = last_obj_risk_8

                obj_id = db().insert_object_with_return_id(
                    subject_id=subj_id,
                    kind=obj_kind,
                    address=obj_address,
                    risk_id=db().find_risk_id(obj_risk)
                )
                insp_kind = row[6]

                insp_number = row[7]

                insp_id = db().insert_inspection_with_return_id(
                    plan_id=db().get_plan_proverok_id(plan_number),
                    kind_id=db().get_kind_inspection_id(insp_kind),
                    number=insp_number
                )
                db().insert_m_to_m_object_inspection(
                    inspection_id=insp_id,
                    object_id=obj_id
                )

            else:
                if row[1] is not None:
                    row = replace_values(row)
                    if row[5] is not None:
                        obj_risk = row[5]
                        last_obj_risk_5 = row[5]
                    else:
                        if row[8] is not None:
                            obj_risk = risk_table[row[8]]
                            last_obj_risk_8 = risk_table[row[8]]
                        else:
                            if last_obj_risk_5 is not None:
                                obj_risk = last_obj_risk_5
                            else:
                                obj_risk = last_obj_risk_8
                    obj_id = db().insert_object_with_return_id(
                        subject_id=subj_id,
                        kind=row[4],
                        address=row[1],
                        risk_id=db().find_risk_id(obj_risk)
                    )
                    # print(f'{row[7]=}')
                    # inspection_id = db().get_inspection_id(row[7])

                    db().insert_m_to_m_object_inspection(
                        inspection_id=insp_id,
                        object_id=obj_id
                    )

                else:

                    continue
    return 0


def main():
    Exel_work().merge_files_in_directory('/Volumes/KINGSTON/готовые проф визиты/готовые', 1)


if __name__ == '__main__':
    main()
