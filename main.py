def open_eplan_file():
    """
    Функция собирает все данные из txt файла с этикетками
    :return: вернуть простыню с данными
    """
    label = []  # пустой список для этикеток
    try:
        with open(
                "C:/Users/apolo/OneDrive/Рабочий стол/Рабочие проекты/Сборочный участок/Наклейки на устройства.txt", "r") as all_label:
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


def save_to_file(label_for_print):
    if label_for_print:
        with open(
                "C:/Users/apolo/OneDrive/Рабочий стол/Рабочие проекты/Сборочный участок/Наклейки на устройства.txt",
                "w") as file:
            for item in label_for_print:
                file.write("\n%s" % item)
    else:
        print("Нет данных для записи в файл")


all_data_from_file = open_eplan_file()  # простыня с данными из файла
uniq_label_for_printing = delete_fail_and_save_to_file(all_data_from_file)  # уникальные этикетки (готов)
save_to_file(uniq_label_for_printing)


