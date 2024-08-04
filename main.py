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


# Функция для преобразования значения в формат "часы и минуты"
def convert_to_hours_minutes(value):
    try:
        if '.' in value:
            hours_str, minutes_str = value.split('.')
            hours = int(hours_str)
            minutes = int(minutes_str)

            if minutes >= 60:
                print("Ошибка: Недопустимое количество минут. Минуты должны быть в пределах от 0 до 59.")
                return None

            return hours, minutes
        else:
            hours = int(value)
            return hours, 0
    except ValueError:
        print("Неверный формат. Пожалуйста, используйте только цифры и точку.")
        return None


# Функция для извлечения часов и минут из строки
def extract_hours_minutes(time_str):
    try:
        parts = time_str.split()
        hours = int(parts[0].replace('ч', ''))
        minutes = int(parts[1].replace('мин', ''))
        return hours, minutes
    except Exception as e:
        print(f"Ошибка при извлечении часов и минут: {e}")
        return 0, 0


# Функция для вычитания времени
def subtract_times(hours1, minutes1, hours2, minutes2):
    total_minutes1 = hours1 * 60 + minutes1
    total_minutes2 = hours2 * 60 + minutes2
    difference_minutes = total_minutes1 - total_minutes2
    difference_hours = difference_minutes // 60
    difference_minutes = difference_minutes % 60
    return difference_hours, difference_minutes


# Функция для подсчета суммы значений в ячейках от start_col до end_col
def sum_values(start_col, end_col, row_index):
    total_hours = 0
    total_minutes = 0
    for col in range(start_col, end_col + 1):
        value = data[row_index, col]
        if value:
            hours, minutes = extract_hours_minutes(value)
            total_hours += hours
            total_minutes += minutes
    total_hours += total_minutes // 60
    total_minutes = total_minutes % 60
    return total_hours, total_minutes


# Запрашиваем у пользователя количество часов
hours_1_to_15 = input("Введите количество часов с 1 по 15 число (например, '5.30' для 5 часов 30 минут): ")
hours_total = input("Введите количество часов Итоговое (например, '10.45' для 10 часов 45 минут): ")

converted_hours_1_to_15 = convert_to_hours_minutes(hours_1_to_15)
converted_hours_total = convert_to_hours_minutes(hours_total)

# Заполняем ячейки Q3 и AH3 введенными значениями
if converted_hours_1_to_15:
    data[2, 16] = f"{converted_hours_1_to_15[0]}ч {converted_hours_1_to_15[1]}мин"
if converted_hours_total:
    data[2, 33] = f"{converted_hours_total[0]}ч {converted_hours_total[1]}мин"


# Функция для установки значений в ячейки
def set_day_value(day, value, row_index):
    if 1 <= day <= 15:
        col_index = day  # B3 соответствует 1, C3 соответствует 2 и т.д.
    elif 16 <= day <= 31:
        col_index = day + 1  # R3 соответствует 17, S3 соответствует 18 и т.д.
    else:
        print("Некорректный номер дня. Он должен быть от 1 до 31.")
        return

    # Преобразуем введенное значение в формат "часы и минуты"
    converted_value = convert_to_hours_minutes(value)
    if converted_value:
        hours, minutes = converted_value
        data[row_index, col_index] = f"{hours}ч {minutes}мин"

# Заполняем ячейку под текущей
        if hours == 16:
            data[row_index + 1, col_index] = f"2ч 0мин"
        elif hours == 8:
            data[row_index + 1, col_index] = f"6ч 0мин"

        # Заполняем следующую ячейку, если значение 16
        if hours == 16:
            if col_index < 31:
                next_col_index = col_index + 1
                data[row_index, next_col_index] = f"8ч 0мин"
                data[row_index + 1, next_col_index] = f"6ч 0мин"


# Изначально вводим значения в строку 3
current_row_index = 2
sum_hours = 0
sum_minutes = 0

# Флаг для отслеживания, нужно ли переключаться на строку 5
switch_to_row_5 = False
switch_to_row_7 = False

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
        value = input(f"Введите значение для дня {day} (например, '2.15' для 2 часов 15 минут): ")
        set_day_value(day, value, current_row_index)

        # Если работаем с днями 1-15
        if 1 <= day <= 15:
            # Считаем сумму значений в ячейках от B3 до P3
            sum_hours, sum_minutes = sum_values(1, 15, 2)

            # Если сумма значений равна или превышает hours_1_to_15, переключаемся на строку 5
            if converted_hours_1_to_15 and sum_hours * 60 + sum_minutes >= converted_hours_1_to_15[0] * 60 + \
                    converted_hours_1_to_15[1]:
                current_row_index = 4
                #switch_to_row_5 = True

        # Если работаем с днями 16-31
        if 16 <= day <= 31:
            # Считаем сумму значений в ячейках от R3 до AG3
            sum_hours, sum_minutes = sum_values(17, 32, 2)

            # Если сумма значений равна или превышает hours_total, переключаемся на строку 7
            if converted_hours_total and sum_hours * 60 + sum_minutes >= converted_hours_total[0] * 60 + \
                    converted_hours_total[1]:
                current_row_index = 4
                #switch_to_row_7 = True

    except ValueError:
        print("Пожалуйста, введите корректное число для дня.")

# Считаем сумму значений в ячейках от B3 до P3 и от R3 до AG3
start_col1 = 1
end_col1 = 15
sum_hours_1, sum_minutes_1 = sum_values(start_col1, end_col1, 2)

start_col2 = 17
end_col2 = 32
sum_hours_2, sum_minutes_2 = sum_values(start_col2, end_col2, 4)

# Вычисляем разницу для обоих диапазонов
if converted_hours_1_to_15:
    hours1, minutes1 = converted_hours_1_to_15
    hours2, minutes2 = sum_hours_1, sum_minutes_1
    difference_hours_1, difference_minutes_1 = subtract_times(hours1, minutes1, hours2, minutes2)

if converted_hours_total:
    hours1, minutes1 = converted_hours_total
    hours2, minutes2 = sum_hours_2, sum_minutes_2
    difference_hours_2, difference_minutes_2 = subtract_times(hours1, minutes1, hours2, minutes2)

# Выводим результаты
print(f"Сумма значений в ячейках B3:P3: {sum_hours_1}ч {sum_minutes_1}мин")
print(f"Остаток времени для ячеек с 1 по 15: {difference_hours_1}ч {difference_minutes_1}мин")

print(f"Сумма значений в ячейках R3:AG3: {sum_hours_2}ч {sum_minutes_2}мин")
print(f"Остаток времени для ячеек с 16 по 31: {difference_hours_2}ч {difference_minutes_2}мин")

# Создаем DataFrame без заголовков
df = pd.DataFrame(data)

# Сохраняем DataFrame в Excel
df.to_excel('generated_table.xlsx', index=False, header=False)

print("Таблица успешно создана и сохранена в 'generated_table.xlsx'")
