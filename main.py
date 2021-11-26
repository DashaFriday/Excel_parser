import openpyxl as op
import numpy as np
import matplotlib.pyplot as plt
import time


# выбор всех уникальных строк из файла по указанному столбцу
def unique_label(sheet, column_char):
    count: int = 2
    j = column_char + str(count)
    list_of_unique_names = []

    while str(sheet[j].value) != "None":
        if sheet[j].value in list_of_unique_names:
            count += 1
            j = column_char + str(count)
            continue
        else:
            list_of_unique_names.append(sheet[j].value)
            count += 1
            j = column_char + str(count)

    return list_of_unique_names, count


# вывод инфы в консоль и добавление её в массив (счётчик, среднее, 2 перцентиля)
def print_info(count, sum, arr_of_values):
    text = []
    text.append('Count: ' + str(count))
    print(text[0])
    text.append('Average: ' + str(round(sum / count, 2)))
    print(text[1])
    text.append('Percentile 95: ' + str(round(np.percentile(arr_of_values, 95), 2)))
    print(text[2])
    text.append('Percentile 99: ' + str(round(np.percentile(arr_of_values, 99), 2)))
    print(text[3])
    return text


# буквы нужных столбцов (может позже сделать ввод пользователя)
col_elapse = 'A'
col_t_s = 'B'
col_label = 'C'
start_label = 'C2'

# выбор файла эксель (может позже сделать ввод пользователя)
# выбор активного листа
# ВАЖНО
# подразумевается, что файл может быть не отсортирован по именам, но должен быть отсортирован по времени
# также в файле стоит убрать / в именах
my_file = op.load_workbook('Agreggare_23.11_night.xlsx')
sheet = my_file.active

# получаем счётчик и массив с уникальными лейблами, сортируем массив уникальных имён (?)
main_array, counter_of_lines = unique_label(my_file.active, col_label)

main_array.sort()

# прогон кода для каждого лейбла
all_names = map(str, main_array)
for name in all_names:

    print(name)
    pass_counter = 0
    summa = 0
    arr_of_elapse = []
    arr_of_ts = []

    # начиная со второй строки (знаем, что первая - шапка таблицы), для каждой строки файла прогоняем:
    l = start_label
    for k in range(2, counter_of_lines):

        # задаём адреса ячеек обрабатываемой строки
        l = col_label + str(k)
        e = col_elapse + str(k)
        ts = col_t_s + str(k)

        # если имя совпадает с искомым добавляем счётчик, плюсуем сумму
        if str(sheet[l].value) == str(name):
            pass_counter += 1
            summa += int(str(sheet[e].value))

            # переводим таймстемп в человекочитаемый вид
            timestamp = int(str(sheet[ts].value)) / 1000
            time_format = time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime(timestamp))[11:]

            # добавляем полученные значения в массивы
            arr_of_elapse.append(sheet[e].value)
            arr_of_ts.append(time_format)


    # в массив выводим итоговую информацию по транзакции
    main_array = print_info(pass_counter, summa, arr_of_elapse)

    # настраиваем вид графика, данные и тд
    plt.style.use('seaborn-colorblind')

    plt.figure(figsize=(12, 6), facecolor='aliceblue')

    # для графика точками plt.scatter
    plt.plot(arr_of_ts, arr_of_elapse, label ='Время отклика')

    plt.suptitle(name)
    plt.subplots_adjust(left=0.06, bottom=0.15, top=0.94, right=0.8)
    plt.xticks(arr_of_ts[::5], rotation=90)
    plt.grid(axis='both', alpha=.2)
    plt.legend(loc='upper left')
    plt.figtext(0.83, 0.8, main_array[0])
    plt.figtext(0.83, 0.7, main_array[1])
    plt.figtext(0.83, 0.6, main_array[2])
    plt.figtext(0.83, 0.5, main_array[3])

    # вывод графика в окно
    plt.show()

    # # сохранение графика в папку проекта и закрытие графика
    # way = str(name) + '.png'
    # plt.savefig(way, bbox_inches = 'tight')
    # plt.close()

my_file.close()
