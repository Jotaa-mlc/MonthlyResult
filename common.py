from datetime import datetime
import calendar

def isFromMonth(date, month_str:str, year_str:str):
    first_day = datetime(int(month_str), int(year_str), 1)
    
    last_day = calendar.monthrange(int(month_str), int(year_str))[-1]
    last_day = datetime(int(month_str), int(year_str), last_day)
    
    if ((first_day <= date) and (date <= last_day)):
        return True
    else:
        return False