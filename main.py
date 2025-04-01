import plotly.express as px
import pandas as pd
import random
import sys

def generate_pastel_color(background_color='lightgray'):
    # Преобразуем цвет фона в RGB
    bg_rgb = {
        'lightgray': (211, 211, 211),
        'lightyellow': (255, 255, 224),
    }.get(background_color, (211, 211, 211))

    r, g, b = 0, 0, 0  # Инициализируем переменные
    for _ in range(10000):
        r = random.randint(200, 255)
        g = random.randint(200, 255)
        b = random.randint(200, 255)
        # Проверяем контрастность
        if abs(r - bg_rgb[0]) > 100 or abs(g - bg_rgb[1]) > 100 or abs(b - bg_rgb[2]) > 100:
            return f'rgba({r}, {g}, {b}, 0.8)'  # Прозрачность для пастельного оттенка
    return f'rgba({r}, {g}, {b}, 0.8)'  # Прозрачность для пастельного оттенка

correct_columns = ['object', 'start', 'finish', 'data', 'legend_title', 'type_object', 'diagram_title', 'yaxis_title']

def main():
    print(f"""
Убедитесь, что в папке со скриптом находится файл DataGant.xlsx.
Убедитесь, что файл DataGant.xlsx содержит следующие столбцы:
{correct_columns}.
Убедитесь, что файл DataGant.xlsx содержит заполненную как минимум первую строку.
Бары на графике окрашиваются в случайные цвета. Запускайте скрипт несколько раз, пока не получите удачный вариант.
          """)
    input('Для продолжения нажмите Enter или закройте программу.')

    try:
        df = pd.read_excel('DataGant.xlsx')
    except Exception as e:
        print(f'Ошибка при загрузки DataGant.xlsx: {e}')
        print('Закройте программу и проверьте наличе файла DataGant.xlsx в папке со скриптом.')
        input()
        sys.exit()

    # проверяем, что файл с исходными данными
    if set(correct_columns) != set(df.columns):
        print('Ошибка при загрузки DataGant.xlsx')
        print(f'Закройте программу и проверьте наличе следующих столбцов в DataGant.xlsx: {correct_columns}.')
        input()
        sys.exit()

    # проверяем, что в файле заполнена хотя бы одна строка
    if df.iloc[0].isnull().any():
        print('Ошибка при работе с DataGant.xlsx')
        print('Закройте программу и проверьте наличе заполненной первой строки.')
        input()
        sys.exit()

    # устанавливаем формат даты для столбцов с началом и окончанием объекта
    df.start = pd.to_datetime(df.start)
    df.finish = pd.to_datetime(df.finish)

    # помечаем категории объектов рандомным цветом
    unique_type_object = df.type_object.unique()
    color_map = {i: generate_pastel_color() for i in unique_type_object}
    df['colors'] = df.type_object.map(color_map)

    # создаем график
    fig = px.timeline(df,
                      x_start="start",
                      x_end="finish",
                      y="object",
                      color="type_object",
                      hover_name='data',
                      color_discrete_map=color_map,
                      title = df.diagram_title[0],
                      color_discrete_sequence=px.colors.qualitative.Plotly,
                      hover_data={'data': False,
                                  'start': False,
                                  'finish': False,
                                  'object': False,
                                  'type_object': False})
    fig.update_yaxes(autorange="reversed") # в противном случае задачи перечисляются снизу вверх

    # Добавление вертикальных линий для разграничения недель
    start_date = df.start.min()
    end_date = df.finish.max()

    # Находим первый понедельник перед минимальной датой
    first_monday = start_date - pd.Timedelta(days=start_date.weekday())

    tickvals_list = []
    ticktext_list = []
    for i in range((end_date - first_monday).days // 7+1):
        week_start = first_monday + pd.Timedelta(weeks=i)
        tickvals_list.append(week_start)
        ticktext_list.append(week_start.strftime('%d.%m.%y'))

    # Определяем начало и конец недели
    fig.update_xaxes(
    tickvals=tickvals_list,
    ticktext=ticktext_list
)
    # Настройки внешнего вида
    fig.update_layout(
        title_font=dict(size=24),
        yaxis_title = df.yaxis_title[0],
        legend_title = df.legend_title[0],
        hoverlabel = dict(bgcolor="white", font_size=12, font_family="Arial"),
    )

    fig.update_layout(plot_bgcolor='#66CCCC')
    fig.write_html("project_timeline.html")
    print('Проверьте файл project_timeline.html в текущей папке.')
    input('Программу можно закрыть.')

if __name__ == "__main__":
   main()
