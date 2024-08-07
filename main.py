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

# Функция для подсчета суммы значений в ячейках от start_col до end_col
def sum_values(start_col, end_col, row_index):
    total_value = 0.0
    for col in range(start_col, end_col + 1):
        value = data[row_index, col]
        if value:
            total_value += float(value)
    return total_value

# Запрашиваем у пользователя количество часов
hours_1_to_15 = input("Введите количество часов с 1 по 15 число (например, '5.5' для 5 часов 30 минут): ")
hours_total = input("Введите количество часов Итоговое (например, '10.75' для 10 часов 45 минут): ")

# Функция для форматирования числа
def format_value(value):
    return str(int(float(value))) if float(value).is_integer() else value

# Заполняем ячейки Q3 и AH3 введенными значениями
if hours_1_to_15:
    data[2, 16] = format_value(hours_1_to_15)
if hours_total:
    data[2, 33] = format_value(hours_total)

# Функция для установки значений в ячейки
def set_day_value(day, value, row_index):
    if 1 <= day <= 15:
        col_index = day  # B3 соответствует 1, C3 соответствует 2 и т.д.
    elif 16 <= day <= 31:
        col_index = day + 1  # R3 соответствует 17, S3 соответствует 18 и т.д.
    else:
        print("Некорректный номер дня. Он должен быть от 1 до 31.")
        return

    # Устанавливаем значение в соответствующую ячейку
    formatted_value = format_value(value)
    data[row_index, col_index] = formatted_value

    # Заполняем ячейку под текущей
    if float(value) == 16.0:
        data[row_index + 1, col_index] = "2"
    elif float(value) == 8.0:
        data[row_index + 1, col_index] = "6"

    # Заполняем следующую ячейку, если значение 16
    if float(value) == 16.0:
        if col_index < 31:
            next_col_index = col_index + 1
            data[row_index, next_col_index] = "8"
            data[row_index + 1, next_col_index] = "6"

# Изначально вводим значения в строку 3
current_row_index = 2

# Запрашиваем номер дня и значение до тех пор, пока пользователь не введет 'stop'
while True:
    day_input = input("Введите номер дня от 1 до 31 (или 'stop' для завершения): ").strip()
    if day_input.lower() == 'stop':
        break

    if day_input.lower() == 'next':
        current_row_index = 2
        continue

    try:
        day = int(day_input)
        value = input(f"Введите значение для дня {day} (например, '2.25' для 2 часов 15 минут): ")
        set_day_value(day, value, current_row_index)

        # Если работаем с днями 1-15
        if 1 <= day <= 15:
            # Считаем сумму значений в ячейках от B3 до P3
            sum_value = sum_values(1, 15, 2)

            # Если сумма значений равна или превышает hours_1_to_15, переключаемся на строку 5
            if hours_1_to_15 and sum_value >= float(hours_1_to_15):
                current_row_index = 4

        # Если работаем с днями 16-31
        if 16 <= day <= 31:
            # Считаем сумму значений в ячейках от R3 до AG3
            sum_value = sum_values(17, 32, 2)

            # Если сумма значений равна или превышает hours_total, переключаемся на строку 7
            if hours_total and sum_value >= float(hours_total):
                current_row_index = 4

    except ValueError:
        print("Пожалуйста, введите корректное число для дня.")

# Считаем сумму значений в ячейках от B3 до P3 и от R3 до AG3
start_col1 = 1
end_col1 = 15
sum_value_1 = sum_values(start_col1, end_col1, 2)

start_col2 = 17
end_col2 = 32
sum_value_2 = sum_values(start_col2, end_col2, 2)

# Выводим результаты
print(f"Сумма значений в ячейках B3:P3: {sum_value_1:.1f}")
print(f"Сумма значений в ячейках R3:AG3: {sum_value_2:.1f}")

# Создаем DataFrame без заголовков
df = pd.DataFrame(data)

# Сохраняем DataFrame в Excel
df.to_excel('generated_table.xlsx', index=False, header=False)

print("Таблица успешно создана и сохранена в 'generated_table.xlsx'")
