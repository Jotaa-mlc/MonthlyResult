from datetime import datetime
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
import calendar
import settings

def IsFromMonth(date, month_str:str, year_str:str):
    first_day = datetime(int(month_str), int(year_str), 1)
    
    last_day = calendar.monthrange(int(month_str), int(year_str))[-1]
    last_day = datetime(int(month_str), int(year_str), last_day)
    
    if ((first_day <= date) and (date <= last_day)):
        return True
    else:
        return False

def LoadSheet(file:str) -> Worksheet:
    file_path = settings.sheets_folder + file
    wb = load_workbook(file_path)
    return wb.active