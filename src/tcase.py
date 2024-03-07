"""
Описание одного теста.
"""
from dataclasses import dataclass
from pathlib import Path

# openpyxl
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet


@dataclass
class Action:
    action: str     # do
    result: str     # expected result


@dataclass
class TCase:
    name: str
    automated: bool = False
    preconditions: list[Action] = None
    steps: list[Action] = None
    postconditions: list[Action] = None
    testdata: str = ''
    comments: str = ''
    iterations: str = ''
    priority: str = 'Medium'
    state: str = 'NotReady'
    tags: str = ''


class TRow:
    """ Row in xlsx worksheet in TestIt import."""
    @staticmethod
    def parse_headers_from_str():
        testit_xlsx_headers = 'Id	Direction	Section	TestCaseName	Automated	Preconditions	\
            Steps	Postconditions	ExpectedResult	TestData	Comments	Iterations	Priority	State	\
            CreatedDate	CreatedById	Tags'
        res = {name: index for index, name in enumerate(testit_xlsx_headers.split(), 1)}
        print(res)
        return res

    testit_xlsx_headers = parse_headers_from_str()
    # {'Id': 1, 'Direction': 2, ...}

    @classmethod
    def get_row_data(cls, ws: Worksheet, row_index):
        d = {}
        for header, i in cls.testit_xlsx_headers.items():
            d[header] = ws.cell(row=row_index, column=i).value
        return d



def load_xlsx(filename: str):
    wb = load_workbook(filename=filename)
    ws = wb.active
    print(type(ws))
    print(ws.max_column, ws.max_row)
    row = TRow.get_row_data(ws, row_index=11)
    print(row)



if __name__ == '__main__':
    filename = '../xlsx/Test IT - casino web.xlsx'
    load_xlsx(filename=filename)
