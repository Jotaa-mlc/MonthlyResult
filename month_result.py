import settings
import datetime
from bills_plan import BillsPlan
from openpyxl import load_workbook

class MonthResult:
    def __init__(self, _month:str, _year:str) -> None:
        self.month = _month
        self.year = _year
        self.bills_plan = BillsPlan()
    
    def Calc_receitas(self) -> None:
        file_path = settings.sheets_folder + settings.document_file
        wb = load_workbook(file_path, read_only = True)
        document_ws = wb.active

        for payment_form in document_ws['V:AA']:
            for doc in payment_form:
                match doc.column_letter:
                    case 'V': #DINHEIRO
                        self.bills_plan.account[0].sub_accounts[0].value += doc.value
                    case 'W': #CHEQUE
                        self.bills_plan.account[0].sub_accounts[1].value += doc.value
                    case 'X': #TRANSFERENCIA / PIX
                        self.bills_plan.account[0].sub_accounts[5].value += doc.value
                    case 'Y': #CARTAO
                        self.bills_plan.account[0].sub_accounts[2].value += doc.value
                    case 'Z': #FATURADO
                        self.bills_plan.account[0].sub_accounts[3].value += doc.value
                    case 'AA': #FINANCEIRA
                        self.bills_plan.account[0].sub_accounts[3].value += doc.value
        
        self.bills_plan.accounts[0].CalculateAcount()