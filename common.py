from datetime import datetime
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
import calendar
import settings

def IsFromMonth(date, month_str:str, year_str:str):
    first_day = datetime(int(year_str), int(month_str), 1)
    
    last_day = calendar.monthrange(int(year_str), int(month_str))[-1]
    last_day = datetime(int(year_str), int(month_str), last_day)
    
    if ((first_day <= date) and (date <= last_day)):
        return True
    else:
        return False

def LoadSheet(file:str) -> Worksheet:
    file_path = settings.sheets_folder + file
    try:
        wb = load_workbook(file_path)
    except:
        return None
    return wb.active

def FormatPrint(title:str, value, size = settings.print_size) -> str:
    value_str = f"R$ {value:_.2f}"
    value_str = value_str.replace('.', ',').replace('_', '.')
    txt = f'{title}:'
    txt += ' ' * (size - len(txt) -len(value_str)) + value_str
    return txt

def StockCCVFilePath(month:str, year:str) -> str:
    sub_folder = year + ' ' + month + '\\'
    file_name = year + ' ' + month + ' Estoque CCV.xlsx'
    return sub_folder + file_name