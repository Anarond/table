import pandas as pd
import numpy as np

# Определим количество строк и столбцов
total_rows = 9
total_columns = 35

# Генерируем пустые данные
data = np.full((total_rows, total_columns), '', dtype=object)

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
#first_half_hours = input(
#    "Введите значение первой половины месяца (например, '5.5' для 5 часов 30 минут): ")

first_half_hours = 78.0

# Записываем значение в ячейку Q3 в формате float
if first_half_hours:
    try:
        data[2, 16] = float(first_half_hours)
    except ValueError:
        print("Некорректный формат. Значение не будет учтено.")

# Переменная для АККУМУЛИРОВАНИЯ суммы значений с 1 по 15 день
accumulated_hours = 0.0

# Получаем значение "итого дней" из ячейки Q3
try:
    total_days_first_half = float(first_half_hours)
except ValueError:
    total_days_first_half = 0.0
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

                # Проверяем, есть ли уже значение
                if data[row_index, col_index + 1] != '':
                    #print(row_index)
                    #print(col_index + 1)
                    data[row_index, col_index + 1]  = float(data[row_index, col_index + 1]) + 8
                else:
                    data[row_index, col_index + 1] = 8

                # В строке ниже (на одну строку ниже текущей) заполняем 2 и 6
                data[row_index + 1, col_index] = 2
                data[row_index + 1, col_index + 1] = 6

# Словарь с условными значениями
conditional_values = {
    'q': 24.0,
    'w': 7.8,
    'e': 9.75
}

# Цикл ввода значений для дней
while True:
    day_input = input(
        "Введите номер дня от 1 до 15 ('end' для завершения, 'format' для форматирования): ").strip()

    if day_input.lower() == 'end':
        break

    if day_input.lower() == 'format':
        apply_format_logic()
        continue

    try:
        day = int(day_input)
        if 1 <= day <= 15:
            value = input(
                f"Введите значение для дня {day} (q - 24 w - 7.8 e - 9.75): ").strip()

            # Проверяем, является ли введённое значение условным
            if value in conditional_values:
                value_float = conditional_values[value]  # Получаем числовое значение из словаря
            else:
                # Если значение не в словаре, пробуем преобразовать его в float
                value_float = float(value)

            col_index = day  # Индекс столбца для B3 - P3 (1 - 15)
            potential_total_hours = accumulated_hours + value_float

            # Сохраняем последнее введённое значение
            last_value_input = value_float

            if potential_total_hours > total_days_first_half:
                # Вычисляем остаток
                remaining_hours = potential_total_hours - total_days_first_half

                # Записываем значение без остатка в текущую строку
                data[current_row, col_index] = round(value_float - remaining_hours, 2)

                # Обрабатываем случай, если последний ввод был равен 24
                if last_value_input == 24:
                    required_value_for_remaining = 16 - (value_float - remaining_hours)
                    data[current_row + 2, col_index] = round(required_value_for_remaining, 2)
                    data[current_row + 2, col_index + 1] = round(24 - (value_float - remaining_hours) - required_value_for_remaining, 2)
                    data[current_row + 3, col_index] = 2
                    data[current_row + 3, col_index + 1] = 6
                else:
                    # Переносим остаток на строку через одну ниже
                    data[current_row + 2, col_index] = round(remaining_hours, 2)

                # Обновляем текущую строку и сумму
                accumulated_hours = remaining_hours
                current_row += 2
            else:
                # Если сумма не превышает, просто записываем значение
                data[current_row, col_index] = value_float
                accumulated_hours += value_float

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