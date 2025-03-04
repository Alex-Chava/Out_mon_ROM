import pandas as pd
import re
import os

# Чтение исходного файла Excel
df = pd.read_excel('Самолет_объекты_25.xlsx')

# Определяем количество строк в исходном файле
num_rows = len(df)  # Количество строк

# Создаем списки для хранения результатов
result = str()
mon_gk = [] # для значений ЖК
mon_tp = []  # Для найденных значений по шаблонам "ТП-XXXXXX" или "РП-XXXXXX"
mon_adr = [] # для значений адрес
mon_vru = []  # Для найденных значений "ВРУ-XX" или "ВРУ XX"
mon_cell = []  # Для найденных значений "ввод XX" или "яч. XX"
mon_npu = []  # Для значений из колонки X

# Основной цикл для обработки строк
for i in range(num_rows):  # Обрабатываем все строки
    value = df.iloc[i, 22]  # Чтение значения из столбца W (индекс 22)
    # Проверяем, содержит ли значение подстроку "Квартира" (без учета регистра)
    if pd.notna(value) and re.search(r'квартира', str(value), re.IGNORECASE):  # Поиск без учета регистра
        # print(f"В строке {i + 1} найдена подстрока 'Квартира'. Пропускаем строку.")
        continue  # Пропускаем эту строку и переходим к следующей

    # Поиск подстроки "ВРУ-XX" или "ВРУ XX" (без учета регистра)
    if pd.notna(value):  # Если значение не пустое
        match_vru = re.search(r'вру[-\s]\d{1,2}', str(value), re.IGNORECASE)  # Шаблон для "ВРУ-XX" или "ВРУ XX"
        result_vru = match_vru.group(0).strip() if match_vru else None
    else:
        result_vru = None
    mon_vru.append(result_vru)

    # Поиск подстроки "яч. XX", "ввод-XX" или "ввод XX"
    if pd.notna(value):  # Если значение не пустое
        match_cell = re.search(r'(яч\.\s*\d{1,2}|ввод[-\s]\d{1,2})', str(value), re.IGNORECASE)
        result_cell = match_cell.group(0).strip() if match_cell else None
    else:
        result_cell = None
    mon_cell.append(result_cell)

    result = df.iloc[i, 8]  # Чтение значения из колонки 8
    mon_gk.append(result)  # Если ячейка пуста, добавляем None
    result = df.iloc[i, 1]  # Чтение значения из колонки 1
    mon_tp.append(result)  # Если ячейка пуста, добавляем None
    result = df.iloc[i, 5]  # Чтение значения из колонки 5
    mon_adr.append(result)  # Если ячейка пуста, добавляем None
    result = df.iloc[i, 23]  # Чтение значения из колонки 23
    mon_npu.append(result)  # Если ячейка пуста, добавляем None


# Проверяем существование файла out_mon.xlsx и удаляем его, если он существует
if os.path.exists('out_mon.xlsx'):
    os.remove('out_mon.xlsx')
    print("Файл out_mon.xlsx удален.")

# Создаем новый DataFrame с 6 колонками (индексы 0-5)
out_df = pd.DataFrame(columns=range(6))  # Создаем DataFrame с 7 колонками

# Записываем значения из столбца W в 0-ю колонку (индекс 0)
out_df[0] = mon_gk  # Записываем значения в первую колонку

# Записываем найденные значения в 1-ю колонку (индекс 1)
out_df[1] = mon_tp  # Записываем результаты во вторую колонку

# Записываем найденные значения "ВРУ-XX" или "ВРУ XX" в 3-ю колонку (индекс 2)
out_df[2] = mon_adr  # Записываем значения в третью колонку

# Записываем найденные значения "ввод XX" и "яч. XX" в 4-ю колонку (индекс 3)
out_df[3] = mon_vru  # Записываем значения в четвертую колонку

# Записываем найденные значения "ввод XX" и "яч. XX" в 4-ю колонку (индекс 3)
out_df[4] = mon_cell  # Записываем значения в четвертую колонку

# Записываем значения из колонки X в 5-ю колонку (индекс 4)
out_df[5] = mon_npu  # Записываем значения в пятую колонку

# Сохранение изменений в файл out_mon.xlsx
out_df.to_excel('out_mon.xlsx', index=False)

print(f"Файл out_mon.xlsx успешно создан. Обработано {num_rows} строк.")