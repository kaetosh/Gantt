import pandas as pd
import os
import webbrowser
import asyncio

from textual.app import App, ComposeResult
from textual.widgets import Header, Footer, Static, LoadingIndicator
from textual.reactive import reactive

from chartGant import chart_write_html
from data_text import (correct_columns,
                       EXAMPLE,
                       TEXT_INTRODUCTION,
                       TEXT_DATAGANTFILE,
                       TEXT_OPEN_DATAGANTFILE,
                       TEXT_MISSING_COL,
                       TEXT_MISSING_DATA_COL,
                       TEXT_CREATE_CHART,
                       TEXT_ERR_CREATE_CHART,
                       TEXT_DATAGANTFILE_NOT_FIND)

class HeaderApp(App):
    CSS = """
    .introduction {
        height: auto;
        border: solid green;
    }
    .steps {
        height: auto;
        border: solid green;
    }
    .indicator {
        height: auto;
    }
    """

    BINDINGS = [
        ("ctrl+r", "open_file", "Открыть DataGant.xlsx для редактирования"),
        ("ctrl+d", "open_diagram", "Сформировать и открыть диаграмму Ганта"),
    ]

    loading_indicator = reactive(None)

    def compose(self) -> ComposeResult:
        self.loading_indicator = LoadingIndicator(classes='indicator')
        yield self.loading_indicator
        yield Header(show_clock=True, icon='<>')
        yield Footer()
        yield Static(TEXT_INTRODUCTION, classes='introduction')
        yield Static(TEXT_DATAGANTFILE, id='example', classes='steps')

    def on_mount(self) -> None:
        self.title = "Ganttify"
        self.sub_title = "мастер по созданию диаграммы Ганта"
        self.loading_indicator.visible = False

    async def action_open_file(self) -> None:
        self.loading_indicator.visible = True
        await asyncio.sleep(2)

        central_widget = self.query_one('#example')
        central_widget.update(TEXT_OPEN_DATAGANTFILE)

        if not os.path.isfile('DataGant.xlsx'):
            example_file = pd.DataFrame(EXAMPLE)
            example_file.to_excel('DataGant.xlsx', index=False)
        os.startfile('DataGant.xlsx')
        self.loading_indicator.visible = False

    async def action_open_diagram(self) -> None:
        self.loading_indicator.visible = True
        await asyncio.sleep(2)

        central_widget = self.query_one('#example')

        try:
            df = pd.read_excel('DataGant.xlsx')
            if set(correct_columns) != set(df.columns):
                missing_columns = set(correct_columns) - set(df.columns)
                missing_columns_str = ', '.join(missing_columns)
                error_message = TEXT_MISSING_COL.format(missing_columns_str=missing_columns_str)
                self.query_one('#example').update(error_message)

            elif df.iloc[0].isnull().any():
                null_columns = df.iloc[0].isnull()
                missing_columns = null_columns[null_columns].index.tolist()
                missing_columns_str = ', '.join(missing_columns)
                error_message = TEXT_MISSING_DATA_COL.format(missing_columns_str=missing_columns_str)
                self.query_one('#example').update(error_message)

            else:
                if chart_write_html(df):
                    webbrowser.open(f'file://{os.path.abspath("project_timeline.html")}')
                    central_widget.update(TEXT_CREATE_CHART)
                else:
                    self.query_one('#example').update(TEXT_ERR_CREATE_CHART)
        except FileNotFoundError:
            self.query_one('#example').update(TEXT_DATAGANTFILE_NOT_FIND)
        self.loading_indicator.visible = False

if __name__ == "__main__":
    app = HeaderApp()
    app.run()

