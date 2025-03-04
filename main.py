import pandas as pd
import re
import os

# Чтение исходного файла Excel
df = pd.read_excel('Самолет_объекты_25.xlsx')

# Определяем количество строк в исходном файле
num_rows = len(df)  # Количество строк

# Создаем списки для хранения результатов
values = []  # Для значений из столбца W (обрезанных и очищенных)
mon_tp = []  # Для найденных значений по шаблонам "ТП-XXXXXX" и "РП-XXXXXX"
mon_cell = []  # Для найденных значений "ввод XX" и "яч. XX"
mon_vru = []  # Для найденных значений "ВРУ-XX" или "ВРУ XX"
x_values = []  # Для значений из колонки X

# Основной цикл для обработки строк
for i in range(num_rows):  # Обрабатываем все строки
    value = df.iloc[i, 22]  # Чтение значения из столбца W (индекс 22)
    x_value = df.iloc[i, 23]  # Чтение значения из колонки X (индекс 23)

    # Заменяем все запятые на пробелы и убираем лишние пробелы
    if pd.notna(value):  # Если значение не пустое
        value = str(value).replace(',', ' ')  # Заменяем запятые на пробелы
        value = ' '.join(value.split())  # Убираем лишние пробелы

    # Проверяем, содержит ли значение подстроку "Квартира" (без учета регистра)
    if re.search(r'квартира', value, re.IGNORECASE):  # Поиск без учета регистра
        print(f"В строке {i + 1} найдена подстрока 'Квартира'. Пропускаем запись.")
        continue  # Пропускаем эту строку и переходим к следующей

    # Обрезаем значение с конца строки до первой с конца открывающейся скобки включительно
    if pd.notna(value):  # Если значение не пустое
        value_str = str(value)  # Преобразуем значение в строку
        # Используем rpartition для поиска последней открывающейся скобки
        before_last_bracket, bracket, after_last_bracket = value_str.rpartition('(')
        if bracket:  # Если скобка найдена
            value_str = before_last_bracket  # Оставляем часть до скобки
        value = value_str.strip()  # Убираем лишние пробелы

    # Поиск подстрок по шаблонам "ТП-XXXXXX" и "РП-XXXXXX", где XXXXXX от 1 до 6 цифр
    if pd.notna(value):  # Если значение не пустое
        match_tp = re.search(r'ТП-\d{1,6}', value, re.IGNORECASE)  # Шаблон для ТП (без учета регистра)
        match_rp = re.search(r'РП-\d{1,6}', value, re.IGNORECASE)  # Шаблон для РП (без учета регистра)

        if match_tp:
            result = match_tp.group(0)  # Найденное значение ТП
            mon_tp.append(result)  # Добавляем в список mon_tp
            value = value.replace(result, '').strip()  # Удаляем найденное значение из строки
        elif match_rp:
            result = match_rp.group(0)  # Найденное значение РП
            mon_tp.append(result)  # Добавляем в список mon_tp
            value = value.replace(result, '').strip()  # Удаляем найденное значение из строки
        else:
            mon_tp.append(None)  # Если шаблон не найден, добавляем None
    else:
        mon_tp.append(None)  # Если ячейка пуста, добавляем None

    # Поиск подстрок "ввод XX" и "яч. XX" (без учета регистра)
    if pd.notna(value):  # Если значение не пустое
        match_vvod = re.search(r'ввод\s+\d{1,2}', value, re.IGNORECASE)  # Шаблон для "ввод XX"
        match_yach = re.search(r'яч\.\s*\d{1,2}', value, re.IGNORECASE)  # Шаблон для "яч. XX"

        if match_vvod:
            result = match_vvod.group(0)  # Найденное значение "ввод XX"
            mon_cell.append(result)  # Добавляем в список mon_cell
            value = value.replace(result, '').strip()  # Удаляем найденное значение из строки
        elif match_yach:
            result = match_yach.group(0)  # Найденное значение "яч. XX"
            mon_cell.append(result)  # Добавляем в список mon_cell
            value = value.replace(result, '').strip()  # Удаляем найденное значение из строки
        else:
            mon_cell.append(None)  # Если шаблон не найден, добавляем None
    else:
        mon_cell.append(None)  # Если ячейка пуста, добавляем None

    # Поиск подстроки "ВРУ-XX" или "ВРУ XX" с конца строки (без учета регистра)
    if pd.notna(value):  # Если значение не пустое
        match_vru = re.search(r'ВРУ[-\s]\d{1,2}\s*$', value, re.IGNORECASE)  # Шаблон для "ВРУ-XX" или "ВРУ XX"
        if match_vru:
            result = match_vru.group(0).strip()  # Найденное значение "ВРУ-XX" или "ВРУ XX"
            mon_vru.append(result)  # Добавляем в список mon_vru
            value = value.replace(result, '').strip()  # Удаляем найденное значение из строки
        else:
            mon_vru.append(None)  # Если шаблон не найден, добавляем None
    else:
        mon_vru.append(None)  # Если ячейка пуста, добавляем None

    # Удаляем из values подстроки, равные mon_tp, mon_cell и mon_vru (если они есть)
    if pd.notna(value):  # Если значение не пустое
        if mon_tp[-1]:  # Если mon_tp не None
            value = value.replace(mon_tp[-1], '').strip()  # Удаляем найденное значение из строки
        if mon_cell[-1]:  # Если mon_cell не None
            value = value.replace(mon_cell[-1], '').strip()  # Удаляем найденное значение из строки
        if mon_vru[-1]:  # Если mon_vru не None
            value = value.replace(mon_vru[-1], '').strip()  # Удаляем найденное значение из строки

    values.append(value)  # Сохраняем очищенное значение в список values
    x_values.append(x_value)  # Сохраняем значение из колонки X

# Проверяем существование файла out_mon.xlsx и удаляем его, если он существует
if os.path.exists('out_mon.xlsx'):
    os.remove('out_mon.xlsx')
    print("Файл out_mon.xlsx удален.")

# Создаем новый DataFrame с 6 колонками (индексы 0-5)
out_df = pd.DataFrame(columns=range(6))  # Создаем DataFrame с 6 колонками

# Записываем значения из столбца W в 0-ю колонку (индекс 0)
out_df[0] = values  # Записываем значения в первую колонку

# Записываем найденные значения в 1-ю колонку (индекс 1)
out_df[1] = mon_tp  # Записываем результаты во вторую колонку

# Записываем найденные значения "ВРУ-XX" или "ВРУ XX" в 3-ю колонку (индекс 2)
out_df[2] = mon_vru  # Записываем значения в третью колонку

# Записываем найденные значения "ввод XX" и "яч. XX" в 4-ю колонку (индекс 3)
out_df[3] = mon_cell  # Записываем значения в четвертую колонку

# Записываем значения из колонки X в 5-ю колонку (индекс 4)
out_df[4] = x_values  # Записываем значения в пятую колонку

# Сохранение изменений в файл out_mon.xlsx
out_df.to_excel('out_mon.xlsx', index=False)

print(f"Файл out_mon.xlsx успешно создан. Обработано {num_rows} строк.")