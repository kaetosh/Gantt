import plotly.express as px
import pandas as pd
import random
from datetime import datetime, timedelta


# Функция для генерации случайных дат отпуска
def generate_vacation_dates(start_date, num_days, num_vacations, vacation_type):
    vacations = []
    for _ in range(num_vacations):
        start = start_date + timedelta(days=random.randint(0, num_days - 1))
        end = start + timedelta(days=random.randint(3, 14))  # Отпуск от 3 до 14 дней
        vacations.append((start, end, vacation_type))
    return vacations


# Генерация данных для 25 сотрудников с фамилиями
surnames = ['Иванов', 'Петров', 'Сидоров', 'Кузнецов', 'Смирнов',
            'Попов', 'Васильев', 'Зайцев', 'Морозов', 'Лебедев',
            'Соловьев', 'Борисов', 'Ковалев', 'Федоров', 'Григорьев',
            'Семенов', 'Алексеев', 'Егоров', 'Тихонов', 'Михайлов',
            'Кириллов', 'Сергеев', 'Денисов', 'Романов']
employees = [f'{surname} {name}' for surname in surnames for name in ['Андрей', 'Мария', 'Елена', 'Игорь']][:25]
data = []

# Задаем начальную дату и количество дней в году
start_date = datetime(2025, 1, 1)
num_days = 365

for employee in employees:
    # Генерируем от 1 до 2 отпусков для каждого сотрудника
    num_vacations = random.randint(1, 2)

    # Случайно выбираем тип отпуска
    for _ in range(num_vacations):
        if random.choice([True, False]):
            vacation_type = 'Отпуск по графику'
        else:
            vacation_type = 'Отпуск по беременности и родам'

        vacations = generate_vacation_dates(start_date, num_days, 1, vacation_type)

        for vac in vacations:
            data.append({
                'Employee': employee,
                'Start': vac[0],
                'Finish': vac[1],
                'Type': vac[2]
            })

# Создание DataFrame
df = pd.DataFrame(data)

# Создание диаграммы Ганта с пастельными тонами
fig = px.timeline(df, x_start='Start', x_end='Finish', y='Employee', title='Годовой график отпусков',
                  color='Type', color_discrete_sequence=px.colors.qualitative.Pastel)

# Настройка фона
fig.update_layout(plot_bgcolor='white', paper_bgcolor='white', font_color='black')

# Инвертируем ось Y
fig.update_yaxes(title='', autorange='reversed')

# Добавление горизонтальных линий по неделям
for week in range(0, 52):
    week_start = start_date + timedelta(weeks=week)
    fig.add_shape(type='line',
                  x0=week_start,
                  y0=-0.5,  # Начальная позиция Y
                  x1=week_start,
                  y1=len(employees) - 0.5,  # Конечная позиция Y
                  line=dict(color='LightGray', width=1, dash='dash'))

# Отображение графика
fig.show()
