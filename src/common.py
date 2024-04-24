from datetime import datetime
from openpyxl import load_workbook, Workbook
import calendar
import settings

def IsFromMonth(date, month:int, year:int):
    first_day = datetime(year, month, 1)
    
    last_day = calendar.monthrange(year, month)[-1]
    last_day = datetime(year, month, last_day)
    
    if ((first_day <= date) and (date <= last_day)):
        return True
    else:
        return False

def LoadSheet(file:str) -> Workbook | None:
    file_path = settings.sheets_folder + file
    try:
        wb = load_workbook(filename=file_path)
    except FileNotFoundError:
        print(f"ERRO: Não foi possível encontrar o arquivo {file_path}")
        return None
    except Exception as error:
        print(f"ERRO DESCONHECIDO: Não foi possível carregar o arquivo {file_path}")
        #print(error)
        return None
    return wb

def FormatPrint(title:str, value, size = settings.print_size) -> str:
    value_str = f"R$ {value:_.2f}"
    value_str = value_str.replace('.', ',').replace('_', '.')
    txt = f'{title}:'
    txt += ' ' * (size - len(txt) -len(value_str)) + value_str
    return txt

def SubFilePath(month:int, year:int, estoque:bool) -> str:
    year_str = "{:04n}".format(year)
    month_str = "{:02n}".format(month)
    sub_folder = year_str + ' ' + month_str + '/'
    file_name = year_str + ' ' + month_str + (' Estoque CCV.xlsx' if estoque else ' Lucratividades.xlsx')
    return sub_folder + file_name