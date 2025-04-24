import pandas as pd
import os
import webbrowser
import asyncio

from textual.app import App, ComposeResult
from textual.widgets import Header, Footer, Static, LoadingIndicator
from textual.reactive import reactive

from example_data import EXAMPLE
from chartGant import chart_write_html

correct_columns = ['object', 'start', 'finish', 'data', 'legend_title', 'type_object', 'diagram_title', 'yaxis_title']

TEXT_INTRODUCTION = """Ganttify позволит легко и быстро создать диаграмму Ганта на основе заполненной таблицы в excel-файле.
Будет предложено заполнить файл DataGant.xlsx, по данным которого Ganttify сформирует файл project_timeline.html, содержащий интерактивную диаграмму Ганта, будь то график отпусков или дорожная карта проекта.
Следуйте инструкции ниже."""

STEP1 = """На данном этапе Ganttify создаст файл DataGant.xlsx, который уже будет заполнен примером Графика отпусков в качестве образца.
Заполните файл своими данными по образцу.
Пожалуйста, не редактируйте заголовки столбцов.
Важно, чтобы как минимум первая строка таблицы была заполнена полностью.
Не забудьте сохранить файл после заполнения.
После вернитесь к окну Ganttify для продолжения.

Чтобы открыть файл DataGant.xlsx для заполнения, нажмите Ctrl+r или соответствующую кнопку внизу.
"""

STEP_1 = """На данном этапе Ganttify создаст файл DataGant.xlsx, содержащий следующие колонки:

| Колонка        | Описание                                                                                                                                               | Примеры заполнения                                    |
|----------------|--------------------------------------------------------------------------------------------------------------------------------------------------------|------------------------------------------------------|
| `object`       | Элементы, которые Ganttify отобразит на оси ординат диаграммы (например, ФИО сотрудника для графика отпусков или этап для дорожной карты проекта). | Иванов И.И.                                  |
| `start`        | Исходная дата соответствующего object.                                                                                                              | 01.01.2023                                           |
| `finish`       | Конечная дата соответствующего object.                                                                                                             | 14.01.2023                                           |
| `data`         | Дополнительная информация, которая будет отображаться при наведении курсора на бар диаграммы (например, количество дней отпуска для графика отпусков или краткое описание этапа для дорожной карты проекта). | Иванов И.И. с 01.01.2023 на 14 дней                                             |
| `legend_title` | Заголовок легенды, достаточно заполнить только первую строку (например, вид отпуска для графика отпусков или ответственный для дорожной карты проекта). | Вид отпуска                       |
| `type_object`  | Дополнительный признак соответствующего object, бары одного признака будут иметь свой цвет, соответствующий легенде (например, основной или дополнительный вид отпуска или ФИО ответственного для дорожной карты проекта). | Основной                                             |
| `diagram_title`| Название диаграммы, достаточно заполнить только первую строку (например, График отпусков инженерного отдела или Капитальный ремонт склада готовой продукции). | График отпусков инженерного отдела                   |
| `yaxis_title`  | Заголовок оси ординат (например, Сотрудник для графика отпусков или Этап для дорожной карты проекта).                                               | Сотрудник                                            |


Файл DataGant.xlsx уже будет заполнен примером Графика отпусков для образца.

Заполните файл своими данными по образцу.

Пожалуйста, не редактируйте заголовки столбцов.

Важно, чтобы как минимум первая строка таблицы была заполнена полностью. Не забудьте сохранить файл после заполнения.

После вернитесь к окну Ganttify для продолжения.

Чтобы открыть файл DataGant.xlsx для заполнения, нажмите Ctrl+r.
"""
class HeaderApp(App):
    CSS = """
    .introduction {
        height: 20%;
        border: solid green;
    }
    .steps {
        height: 80%;
        border: solid green;
    }
    """

    BINDINGS = [
        ("ctrl+r", "open_file", "Открыть DataGant.xlsx для редактирования"),
        ("ctrl+d", "open_diagram", "Сформировать диаграмму Ганта"),
    ]

    loading_indicator = reactive(None)

    def compose(self) -> ComposeResult:
        yield Header(show_clock=True, icon='*')
        yield Footer()
        yield Static(TEXT_INTRODUCTION, classes='introduction')
        yield Static(STEP1, id='example', classes='steps')
        self.loading_indicator = LoadingIndicator()  # Инициализация индикатора загрузки

    def on_mount(self) -> None:
        self.title = "Ganttify"
        self.sub_title = "мастер по созданию диаграммы Ганта"

    async def action_open_file(self) -> None:
        central_widget = self.query_one('#example')
        central_widget.mount(self.loading_indicator)
        self.loading_indicator.visible = True  # Показываем индикатор загрузки

        if not os.path.isfile('DataGant.xlsx'):
            example_file = pd.DataFrame(EXAMPLE)
            example_file.to_excel('DataGant.xlsx', index=False)
        os.startfile('DataGant.xlsx')
        text = """Если файл DataGant.xlsx готов, нажмите ctrl+d, чтобы сформировать и открыть диаграмму Ганта.
        Сформированная диаграмма Ганта будет сохранена под именем project_timeline.html и расположена в папке, откуда запущен Ganttify.
        Чтобы вернуться к редактированию DataGant.xlsx, нажмите ctrl+r"""
        central_widget.update(text)
        await asyncio.sleep(2)
        self.loading_indicator.visible = False  # Скрываем индикатор загрузки после завершения

    async def action_open_diagram(self) -> None:
        central_widget = self.query_one('#example')
        central_widget.mount(self.loading_indicator)
        self.loading_indicator.visible = True  # Показываем индикатор загрузки
        await asyncio.sleep(2)


        try:
            df = pd.read_excel('DataGant.xlsx')
            if set(correct_columns) != set(df.columns):
                text = """Ошибка при загрузки DataGant.xlsx: отсутствует один или несколько необходимых столбцов.
                          Возможно файл был неверно отредактирован.
                          Пожалуйста, не изменяйте заголовки столбцов.
                          Нажмите ctrl+r, чтобы снова сформировать и открыть для редактирования DataGant.xlsx"""
                self.query_one('#example').update(text)

            elif df.iloc[0].isnull().any():
                text = """Ошибка при загрузки DataGant.xlsx: не заполнена первая строка таблицы.
                          Возможно файл был неверно отредактирован.
                          Пожалуйста, заполните как минимум первую строку таблицы полностью.
                          Нажмите ctrl+r, чтобы снова сформировать и открыть для редактирования DataGant.xlsx"""
                self.query_one('#example').update(text)

            else:
                if chart_write_html(df):
                    webbrowser.open(f'file://{os.path.abspath("project_timeline.html")}')
                else:
                    text = """Ошибка при формировании project_timeline.html"""
                    self.query_one('#example').update(text)
        except FileNotFoundError:
            text = """Ошибка при загрузки DataGant.xlsx: файл не найден.
                      Возможно файл был перемещен или удален.
                      Нажмите ctrl+r, чтобы снова сформировать и открыть для редактирования DataGant.xlsx"""
            self.query_one('#example').update(text)
        self.loading_indicator.visible = False  # Скрываем индикатор загрузки после завершения

if __name__ == "__main__":
    app = HeaderApp()
    app.run()

