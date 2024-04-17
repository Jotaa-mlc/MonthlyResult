import settings
import common
import datetime
from bills_plan import BillsPlan
from openpyxl import load_workbook

class MonthResult:
    def __init__(self, _month:str, _year:str) -> None:
        self.month = _month
        self.year = _year
        self.bills_plan = BillsPlan()
    
    def CalcReceitas(self) -> None:
        document_ws = common.LoadSheet(settings.document_file)

        for payment_form in document_ws['V:AA']:
            for doc in payment_form:
                is_pedido = document_ws["A" + str(doc.row)].value == "Pedido   Saída"
                is_fechado = document_ws["C" + str(doc.row)].value == "Fechado"
                is_from_month = common.IsFromMonth(document_ws["I" + str(doc.row)].value, self.month, self.year) if is_fechado else False
                if is_pedido and is_from_month:
                    match doc.column_letter:
                        case 'V': #DINHEIRO
                            self.bills_plan.accounts[0].sub_accounts[0].value += doc.value
                        case 'W': #CHEQUE
                            self.bills_plan.accounts[0].sub_accounts[1].value += doc.value
                        case 'X': #TRANSFERENCIA / PIX
                            self.bills_plan.accounts[0].sub_accounts[5].value += doc.value
                        case 'Y': #CARTAO
                            self.bills_plan.accounts[0].sub_accounts[2].value += doc.value
                        case 'Z': #FATURADO
                            self.bills_plan.accounts[0].sub_accounts[3].value += doc.value
                        case 'AA': #FINANCEIRA
                            self.bills_plan.accounts[0].sub_accounts[3].value += doc.value
        
        self.bills_plan.accounts[0].CalculateAccount()
    
    def CalcDespesas(self):
        payments_ws = common.LoadSheet(settings.payment_file)
        
        for payment in payments_ws:#achar uma forma de iterar por linha!!!
            is_quitado = payment["H" + str(payment.row)].value == "Quitada"
            is_2plano_contas = payment["M" + str(payment.row)].value = "True"
            is_from_month = common.IsFromMonth(payment["P" + str(payment.row)].value, self.month, self.year) if is_quitado else False
            if (is_quitado and is_from_month) and is_2plano_contas:
                cod_conta = payment["N" + str(payment.row)].value - 1
                cod_subconta = payment["O" + str(payment.row)].value - 1
                
                self.bills_plan.accounts[cod_conta].sub_accounts[cod_subconta].value += payment["S" + str(payment.row)]
            
        
        
        