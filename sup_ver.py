import plotly.express as px
import pandas as pd
import random


def generate_pastel_color(background_color='lightgray'):
    # Преобразуем цвет фона в RGB
    bg_rgb = {
        'lightgray': (211, 211, 211),
        'lightyellow': (255, 255, 224),
        # Добавьте другие цвета по мере необходимости
    }.get(background_color, (211, 211, 211))

    for _ in range(10000):
        r = random.randint(200, 255)
        g = random.randint(200, 255)
        b = random.randint(200, 255)
        # Проверяем контрастность
        if abs(r - bg_rgb[0]) > 100 or abs(g - bg_rgb[1]) > 100 or abs(b - bg_rgb[2]) > 50:
            return f'rgba({r}, {g}, {b}, 0.6)'  # Прозрачность 0.6 для пастельного оттенка
    return f'rgba({r}, {g}, {b}, 0.6)'  # Прозрачность 0.6 для пастельного оттенка

# читаем исходный файл
df = pd.read_excel('DataGant.xlsx')

# устанавливаем формат даты для столбцов с началом и окончанием объекта
df.Start = pd.to_datetime(df.Start)
df.Finish = pd.to_datetime(df.Finish)

# помечаем категории объектов рандомным цветом
unique_type_object = df.type_object.unique()
color_map = {i: generate_pastel_color() for i in unique_type_object}
df['colors'] = df.type_object.map(color_map)

# создаем график
fig = px.timeline(df,
                  x_start="Start",
                  x_end="Finish",
                  y="Task",
                  color="type_object",
                  hover_name='text',
                  color_discrete_map=color_map,
                  hover_data={'text': False,
                              'Start': False,
                              'Finish': False,
                              'Task': False,
                              'type_object': False})
fig.update_yaxes(autorange="reversed") # otherwise tasks are listed from the bottom up
# Добавление вертикальных линий для разграничения недель
start_date = df.Start.min()
end_date = df.Finish.max()

# Находим первую понедельник перед минимальной датой
first_monday = start_date - pd.Timedelta(days=start_date.weekday())

# Добавляем линии для каждой недели от первого понедельника до последнего воскресенья
for i in range((end_date - first_monday).days // 7 + 2):
    week_start = first_monday + pd.Timedelta(weeks=i)
    fig.add_shape(type="line",
                  x0=week_start,
                  y0=0,
                  x1=week_start,
                  y1=len(df),
                  line=dict(color="White", width=2, dash="dash"))
    # Добавление подписи с датой
    fig.add_annotation(
        x=week_start,
        y=len(df) + 1,  # Позиция по оси Y для аннотации
        text=week_start.strftime('%d.%m'),  # Формат даты
        showarrow=False,
        font=dict(size=10),
        xanchor='center'
    )

fig.update_layout(plot_bgcolor='#66CCCC')
fig.write_html("project_timeline.html")

=СЦЕПИТЬ("с ";ТЕКСТ(B2; "ДД МММ.");" на ";C2-B2;" дн.")
Task	Start	Finish	text	type_object	name_diagram
Иванов	25.02.2023	03.03.2023	с 25 фев. на 6 дн.	согласно графика	График отпусков
Иванов	10.03.2023	17.03.2023	с 10 мар. на 7 дн.	согласно графика	График отпусков
Здоровцев	13.03.2023	22.03.2023	с 13 мар. на 9 дн.	согласно графика	График отпусков
Милютьина	23.03.2023	30.03.2023	с 23 мар. на 7 дн.	перенос	График отпусков
Черевичкин	01.04.2023	07.04.2023	с 01 апр. на 6 дн.	согласно графика	График отпусков
Кнопкина	05.04.2023	18.04.2023	с 05 апр. на 13 дн.	согласно графика	График отпусков
Черевичкин	12.04.2023	23.04.2023	с 12 апр. на 11 дн.	перенос	График отпусков
Оборочков	20.04.2023	25.04.2023	с 20 апр. на 5 дн.	согласно графика	График отпусков
Хреновский	24.04.2023	03.05.2023	с 24 апр. на 9 дн.	согласно графика	График отпусков
Черевичкин	02.05.2023	07.05.2023	с 02 май. на 5 дн.	согласно графика	График отпусков
