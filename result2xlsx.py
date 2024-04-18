from month_result import MonthResult
from bills_plan import BillsPlan
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font, PatternFill, Border, Side
import settings

def GenerateHeaders( export_file:Worksheet) -> None:
    thick = Side(style='thick', color='000000')
    thin = Side(style='thin', color='000000')
    
    title_font = Font(size = 13, bold = True, color = 'FFFFFF')
    title_fill = PatternFill("solid", fgColor = '000000')
    title_border = Border(top=thick, bottom=thick, left=thick, right=thick)
    
    account_font = Font(size=12, bold=True)
    account_border = title_border
    
    sub_account_font = Font(size=11)
    sub_account_border = Border(top=thin, bottom=thin, left=thin, right=thin)
    
    export_file['A1'].value = 'Cód. Conta'
    export_file['B1'].value = 'Descrição'
    export_file['A1:B1'].font = title_font
    export_file['A1:B1'].fill = title_fill
    export_file['A1:B1'].border = title_border
    
    model = BillsPlan()
    
    header_row = 2
    for account_index in range(len(model.accounts)):
        export_file["A" + str(header_row)].value = account_index + 1
        export_file["B" + str(header_row)].value = model.accounts[account_index].description
        export_file["A" + str(header_row) + ":B" + str(header_row)].font = account_font
        export_file["A" + str(header_row) + ":B" + str(header_row)].font = account_border
        
        header_row += 1
        for sub_account_index in range(len(model.accounts[account_index].sub_accounts)):
            export_file["A" + str(header_row)].value = sub_account_index + 1
            export_file["B" + str(header_row)].value = model.accounts[account_index].sub_accounts[sub_account_index].description
            export_file["A" + str(header_row) + ":B" + str(header_row)].font = sub_account_font
            export_file["A" + str(header_row) + ":B" + str(header_row)].font = sub_account_border
            
            header_row += 1
    
    

def Export2XLSX(results:list[MonthResult]) -> None:
    for result in results:
        result.PrintBillsPlan()