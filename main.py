import pandas as pd
import re
import os

# Чтение исходного файла Excel
df = pd.read_excel('Самолет_объекты_25.xlsx')

# Определяем количество строк в исходном файле
num_rows = len(df)  # Количество строк

# Создаем списки для хранения результатов
values = []  # Для значений из столбца W
results = []  # Для найденных значений по шаблонам "ТП-XXXXXX" и "РП-XXXXXX"
bracket_values = []  # Для значений в скобках
x_values = []  # Для значений из колонки X

# Основной цикл для обработки строк
for i in range(num_rows):  # Обрабатываем все строки
    value = df.iloc[i, 22]  # Чтение значения из столбца W (индекс 22)
    x_value = df.iloc[i, 23]  # Чтение значения из колонки X (индекс 23)

    # Проверяем, содержит ли значение подстроку "Квартира"
    if "Квартира" in str(value):  # Если найдена подстрока "Квартира"
        print(f"В строке {i + 1} найдена подстрока 'Квартира'. Пропускаем запись.")
        continue  # Пропускаем эту строку и переходим к следующей

    values.append(value)  # Сохраняем значение в список values
    x_values.append(x_value)  # Сохраняем значение из колонки X

    # Проверка, что значение не пустое
    if pd.notna(value):  # Если значение не NaN
        # Поиск подстрок по шаблонам "ТП-XXXXXX" и "РП-XXXXXX", где XXXXXX от 1 до 6 цифр
        match_tp = re.search(r'ТП-\d{1,6}', str(value))  # Шаблон для ТП
        match_rp = re.search(r'РП-\d{1,6}', str(value))  # Шаблон для РП

        if match_tp:
            results.append(match_tp.group(0))  # Добавляем найденное значение ТП
        elif match_rp:
            results.append(match_rp.group(0))  # Добавляем найденное значение РП
        else:
            results.append(None)  # Если шаблон не найден, добавляем None

        # Поиск значения в скобках в конце строки
        bracket_match = re.search(r'\(([^)]+)\)\s*$', str(value))  # Ищем текст в скобках в конце строки
        if bracket_match:
            bracket_values.append(bracket_match.group(1))  # Извлекаем текст внутри скобок
        else:
            bracket_values.append(None)  # Если скобок нет, добавляем None
    else:
        results.append(None)  # Если ячейка пуста, добавляем None
        bracket_values.append(None)  # Если ячейка пуста, добавляем None

# Проверяем существование файла out_mon.xlsx и удаляем его, если он существует
if os.path.exists('out_mon.xlsx'):
    os.remove('out_mon.xlsx')
    print("Файл out_mon.xlsx удален.")

# Создаем новый DataFrame с 6 колонками (индексы 0-5)
out_df = pd.DataFrame(columns=range(6))  # Создаем DataFrame с 6 колонками

# Записываем значения из столбца W в 0-ю колонку (индекс 0)
out_df[0] = values  # Записываем значения в первую колонку

# Записываем найденные значения в 1-ю колонку (индекс 1)
out_df[1] = results  # Записываем результаты во вторую колонку

# Записываем значения из скобок в 3-ю колонку (индекс 3)
out_df[3] = bracket_values  # Записываем значения в четвертую колонку

# Записываем значения из колонки X в 5-ю колонку (индекс 4)
out_df[4] = x_values  # Записываем значения в пятую колонку

# Сохранение изменений в файл out_mon.xlsx
out_df.to_excel('out_mon.xlsx', index=False)

print(f"Файл out_mon.xlsx успешно создан. Обработано {num_rows} строк.")