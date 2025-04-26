import os

import plotly.express as px
import pandas as pd
import numpy as np


def generate_pastel_color() -> str:
    # Преобразуем цвет фона в RGB
    bg_rgb = (211, 211, 211)

    # Генерируем 10000 случайных цветов
    random_colors = np.random.randint(200, 256, size=(10000, 3))

    # Вычисляем абсолютные разности с фоновым цветом
    contrast = np.abs(random_colors - bg_rgb)

    # Проверяем, какие цвета соответствуют критерию контрастности
    valid_colors = random_colors[(contrast[:, 0] > 100) | (contrast[:, 1] > 100) | (contrast[:, 2] > 100)]

    if valid_colors.size > 0:
        # Возвращаем первый подходящий цвет с прозрачностью
        r, g, b = valid_colors[0]
        return f'rgba({r}, {g}, {b}, 1)'

    # Если не найден ни один подходящий цвет, возвращаем последний сгенерированный
    r, g, b = random_colors[-1]
    return f'rgba({r}, {g}, {b}, 1)'


def chart_write_html(df: pd.DataFrame) -> bool:

    # устанавливаем формат даты для столбцов с началом и окончанием объекта
    df['старт'] = pd.to_datetime(df['старт'], format='%d.%m.%Y')
    df['финиш'] = pd.to_datetime(df['финиш'], format='%d.%m.%Y')

    # помечаем категории объектов рандомным цветом
    unique_type_object = df['признак_объекта'].unique()

    color_map = {i: generate_pastel_color() for i in unique_type_object}
    df['colors'] = df['признак_объекта'].map(color_map)

    # создаем график
    fig = px.timeline(df,
                      x_start="старт",
                      x_end="финиш",
                      y="объект",
                      color="признак_объекта",
                      hover_name='данные',
                      color_discrete_map=color_map,
                      title=df['имя_диаграммы'][0],
                      color_discrete_sequence=px.colors.qualitative.Plotly,
                      hover_data={'данные': False,
                                  'старт': False,
                                  'финиш': False,
                                  'объект': False,
                                  'признак_объекта': False})
    fig.update_yaxes(autorange="reversed")  # в противном случае задачи перечисляются снизу вверх

    # # Добавление вертикальных линий для разграничения недель
    start_date = df['старт'].min()
    end_date = df['финиш'].max()

    # Находим первый понедельник перед минимальной датой
    first_monday = start_date - pd.Timedelta(days=start_date.weekday())

    tickvals_list = []
    ticktext_list = []
    for i in range((end_date - first_monday).days // 7 + 1):
        week_start = first_monday + pd.Timedelta(weeks=i)
        tickvals_list.append(week_start)
        ticktext_list.append(week_start.strftime('%d.%m.%Y'))

    # Определяем начало и конец недели
    fig.update_xaxes(
        tickvals=tickvals_list,
        ticktext=ticktext_list
    )
    # Настройки внешнего вида
    fig.update_layout(
        title_font=dict(size=24),
        yaxis_title=df['имя_вертикали'][0],
        legend_title=df['имя_легенды'][0],
        hoverlabel=dict(bgcolor="white", font_size=12, font_family="Arial"),
    )

    fig.update_layout(plot_bgcolor='#66CCCC')
    fig.write_html("project_timeline.html")
    if os.path.exists("project_timeline.html"):
        return True
    else:
        return False