import pandas as pd
import numpy as np

# Определим количество строк и столбцов
num_rows = 9
num_columns = 35

# Генерируем пустые данные
data = np.full((num_rows, num_columns), '', dtype=object)

# Список значений для первой колонки
first_column_values = ['', '', 'анест', 'ночные', 'совм', 'ночные', 'совм', 'ночные', 'празд']

# Заполняем первую колонку данными
data[:, 0] = first_column_values

# Заполняем ячейки B2:P2 цифрами от 1 до 15
data[1, 1:16] = range(1, 16)

# Заполняем ячейку Q1 значением "итого дней"
data[0, 16] = "итого дней"

# Заполняем ячейки R2:AG2 цифрами от 16 до 31
data[1, 17:33] = range(16, 32)

# Заполняем ячейку AH1 значением "Всего дней"
data[0, 33] = "Всего дней"

# Заполняем ячейку AI1 значением "кол-во смен"
data[0, 34] = "кол-во смен"

# Запрос значения для первой половины месяца
hours_1_to_15 = input(
    "Введите значение первой половины месяца в формате десятичной дроби (например, '5.5' для 5 часов 30 минут): ")

# Записываем значение в ячейку Q3 в формате float
if hours_1_to_15:
    try:
        data[2, 16] = float(hours_1_to_15)
    except ValueError:
        print("Некорректный формат. Значение не будет учтено.")

# Переменная для хранения суммы значений с 1 по 15 день
sum_values = 0.0

# Получаем значение "итого дней" из ячейки Q3
try:
    itogo_days = float(hours_1_to_15)
except ValueError:
    itogo_days = 0.0
    print("Некорректное значение в 'итого дней'. Оно будет считаться как 0.")

# Переменная для отслеживания текущей строки ввода значений
current_row = 2


# Функция применения логики преобразования значений
def apply_format_logic():
    # Проходим по строкам 2 и 4
    for row_index in [2, 4]:
        for col_index in range(1, 16):  # Проходим по столбцам от B до P (1-15)
            value = data[row_index, col_index]
            if value == 24.0:
                # Разбиваем значение '24' на 16 и 8 в текущей строке
                data[row_index, col_index] = 16.0
                data[row_index, col_index + 1] = 8.0

                # В строке ниже (на одну строку ниже текущей) заполняем 2 и 6
                data[row_index + 1, col_index] = 2.0
                data[row_index + 1, col_index + 1] = 6.0


# Цикл ввода значений для дней
while True:
    day_input = input(
        "Введите номер дня от 1 до 15 (или 'end' для завершения, или 'format' для применения логики): ").strip()

    if day_input.lower() == 'end':
        break

    if day_input.lower() == 'format':
        apply_format_logic()
        continue

    try:
        day = int(day_input)
        if 1 <= day <= 15:
            value = input(
                f"Введите значение для дня {day} в формате десятичной дроби (например, '2.25' для 2 часов 15 минут): ").strip()
            col_index = day  # Индекс столбца для B3 - P3 (1 - 15)

            try:
                value_float = float(value)
                potential_sum = sum_values + value_float

                if potential_sum > itogo_days:
                    # Вычисляем остаток
                    remainder = potential_sum - itogo_days

                    # Записываем значение без остатка в текущую строку
                    data[current_row, col_index] = round(value_float - remainder, 2)

                    # Переносим остаток на строку через одну ниже
                    data[current_row + 2, col_index] = round(remainder, 2)

                    # Обновляем текущую строку и сумму
                    sum_values = remainder
                    current_row += 2
                else:
                    # Если сумма не превышает, просто записываем значение
                    data[current_row, col_index] = value_float
                    sum_values += value_float

            except ValueError:
                print(f"Некорректное значение: {value}. Оно не будет учтено.")
        else:
            print("Некорректный номер дня. Укажите номер от 1 до 15.")
    except ValueError:
        print(
            "Некорректный ввод. Пожалуйста, введите число от 1 до 15, 'end' для завершения или 'format' для применения логики.")

# Создаем DataFrame без заголовков
df = pd.DataFrame(data)

# Сохраняем DataFrame в Excel
df.to_excel('generated_table.xlsx', index=False, header=False)

print("Таблица успешно создана и сохранена в 'generated_table.xlsx'")
