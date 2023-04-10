import openpyxl
import os
import zipfile


def open_eplan_file():
    """
    Функция собирает все данные из txt файла с этикетками
    :return: вернуть простыню с данными
    """
    label = []  # пустой список для этикеток
    try:
        with open(
                "C:/Users/apolo/OneDrive/Рабочий стол/Рабочие проекты/Сборочный участок/Наклейки на устройства.txt",
                "r") as all_label:
            for item in all_label:
                # исключить абзацы
                if item != '\n':
                    label.append(item)

    except IOError:
        print("Нет файла с этикетка")

    return label


def delete_fail_and_save_to_file(all_data):
    """
    Функция удаляет ошибочные окончания у элементов списка и сохраняет их в файл
    :param all_data: простыня с данными
    :return: лист с уникальными этикетками
    """
    if not all_data == "":  # если файл не пусто, выполнить код ниже
        uniq_item = set(all_data)  # лист в множество для удаления повторений
        all_label = list(uniq_item)  # обратно в лист
        all_label.sort()  # отсортировать по алфавиту

        uniq_label = []  # лист с уникальными этикетками
        for item in all_label:
            new_item = item[:-1]
            uniq_label.append(new_item)

        return uniq_label

    else:  # если файл пусто, выдать ошибку
        print("Файл пустой")

        return False


def save_to_file_txt(label_for_print):
    if label_for_print:
        with open(
                "C:/Users/apolo/OneDrive/Рабочий стол/Рабочие проекты/Сборочный участок/Наклейки на устройства.txt",
                "w") as file:
            for item in label_for_print:
                file.write("\n%s" % item)
    else:
        print("Нет данных для записи в файл")


def save_to_file_excel(label_for_print):
    """
    Функция сохраняет полученный лист с ОУ устройств в файл Excel
    :param label_for_print: лист с ОУ
    :return:
    """
    try:
        workbook = openpyxl.Workbook()  # Создаём новый файл Excel
        sheet = workbook.active  # Выбираем активный лист, по умолчанию 1

        # Записывам каждый элемент списка в новую строку
        for i in range(len(label_for_print)):
            sheet.cell(row=i + 1, column=1, value=label_for_print[i])

        sheet.title = 'Наклейки для устройств'  # Изменить имя листа

        # Сохраняем файл
        workbook.save(
            "C:/Users/apolo/OneDrive/Рабочий стол/Рабочие проекты/Сборочный участок/Наклейки на устройства.xlsx")

        # Удалить старый файл txt
        os.remove("C:/Users/apolo/OneDrive/Рабочий стол/Рабочие проекты/Сборочный участок/Наклейки на устройства.txt")

    except IOError:
        print("Не могу сохранить файл Excel")


def created_zip_file():
    """
    Создать архив с данными
    :return:
    """
    pass



all_data_from_file = open_eplan_file()  # простыня с данными из файла
uniq_label_for_printing = delete_fail_and_save_to_file(all_data_from_file)  # уникальные этикетки (готов)
# save_to_file_txt(uniq_label_for_printing)  # Сохранение в файл txt
save_to_file_excel(uniq_label_for_printing)  # Сохранение в файл Excel
