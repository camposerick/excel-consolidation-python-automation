from project import is_excel_file, is_xlsb, define_sheet
from openpyxl import Workbook

def test_is_excel_file():
    files = ['test_1.xlsx', 'test_2.xlsx', 'test_3.txt']
    assert is_excel_file(files) == False
    

def test_is_xlsb():
    files = ['test_1.xlsx', 'test_2.xlsx', 'test_3.xlsx']
    assert is_xlsb(files) == False

def test_define_sheet(monkeypatch):
    wb = Workbook()
    ws = wb.active
    ws.title = "Mysheet1"
    ws2 = wb.create_sheet("Mysheet2")
    wb.save("test.xlsx")
    files = ['test.xlsx']
    
    monkeypatch.setattr('builtins.input', lambda _: '2')
    
    assert define_sheet(files)[1] == "Mysheet2"