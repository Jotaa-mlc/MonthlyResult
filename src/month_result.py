import settings, common
from bills_plan import BillsPlan

class MonthResult:
    def __init__(self, _month:int, _year:int) -> None:
        self.month = _month
        self.year = _year
        self.result = 0.0
        self.result_percentage = 0.0
        self.stock_cost = 0.0
        self.bills_plan = BillsPlan()
    
    def CalcResult(self) -> None:
        for account in self.bills_plan.accounts:
            self.result += account.value
        
        self.result_percentage = self.result / self.bills_plan.accounts[0].value
    
    def CalcStockCost(self) -> None:
        file_path = common.SubFilePath(self.month, self.year, True)
        stock_wb = common.LoadSheet(file_path)
        
        if stock_wb == None:
            return None
        
        stock_ws = stock_wb.active
        stock_columns = settings.stockCCV_columns()
        
        for product_row in range(1, stock_ws.max_row):
            amount = stock_ws[stock_columns.amount + str(product_row)].value
            cost = stock_ws[stock_columns.cost + str(product_row)].value
            self.stock_cost += amount * cost
        
        stock_wb.close()
    
    def CalcReceitas(self) -> None:
        document_wb = common.LoadSheet(settings.document_file)
        
        if document_wb == None:
            return None
        
        document_ws = document_wb.active
        document_columns = settings.document_columns()
        
        for document_row in range(1, document_ws.max_row):
            is_pedido = document_ws[document_columns.tp_doc + str(document_row)].value == "Pedido   Saída"
            is_fechado = document_ws[document_columns.status_doc + str(document_row)].value == "Fechado"
            is_from_month = common.IsFromMonth(document_ws[document_columns.dt_doc + str(document_row)].value, self.month, self.year) if is_fechado else False
            
            if is_pedido and is_fechado and is_from_month:
                self.bills_plan.accounts[0].sub_accounts[0].value += document_ws[document_columns.pmt_dinheiro + str(document_row)].value #DINHEIRO
                self.bills_plan.accounts[0].sub_accounts[1].value += document_ws[document_columns.pmt_cheque + str(document_row)].value #CHEQUE
                self.bills_plan.accounts[0].sub_accounts[5].value += document_ws[document_columns.pmt_transferencia + str(document_row)].value #TRANSFERENCIA / PIX
                self.bills_plan.accounts[0].sub_accounts[2].value += document_ws[document_columns.pmt_cartao + str(document_row)].value #CARTAO
                self.bills_plan.accounts[0].sub_accounts[3].value += document_ws[document_columns.pmt_faturado + str(document_row)].value #FATURADO
                self.bills_plan.accounts[0].sub_accounts[4].value += document_ws[document_columns.pmt_financeira + str(document_row)].value #FINANCEIRA
        
        self.bills_plan.accounts[0].CalculateAccount()
        document_wb.close()
    
    def CalcCMV(self) -> None:
        file_path = common.SubFilePath(self.month, self.year, False)
        profits_wb = common.LoadSheet(file_path)
        cmv = 0.0
        
        if profits_wb == None:
            self.bills_plan.accounts[2].sub_accounts[0].value = cmv
            return None
        
        profits_ws = profits_wb.active
        profits_columns = settings.profits_columns()
        
        for product_row in range(1, profits_ws.max_row):
            cmv -= profits_ws[profits_columns.cost_prod + str(product_row)].value
        
        self.bills_plan.accounts[2].sub_accounts[0].value = cmv
        profits_wb.close()
    
    def CalcDespesas(self) -> None:
        #CALCULANDO DESPESAS DA TABELA DE CONTAS A PAGAR
        payments_wb = common.LoadSheet(settings.payment_file)
        cashier_wb = common.LoadSheet(settings.cashier_file)
        
        if payments_wb == None or cashier_wb == None:
            return None
        
        payments_ws = payments_wb.active
        payments_columns = settings.payment_columns()
        cashier_ws = cashier_wb.active
        cahsier_columns = settings.cashier_columns()
        
        for payment_row in range(1, payments_ws.max_row):
            is_quitado = payments_ws[payments_columns.status_pmt + str(payment_row)].value == "Quitada"
            is_2plano_contas = payments_ws[payments_columns.is_2plano_contas + str(payment_row)].value == "True"
            is_from_month = common.IsFromMonth(payments_ws[payments_columns.dt_pmt + str(payment_row)].value, self.month, self.year) if is_quitado else False
            
            if is_quitado and is_from_month and is_2plano_contas:
                cod_conta = payments_ws[payments_columns.cod_conta + str(payment_row)].value - 1
                cod_sub_conta = payments_ws[payments_columns.cod_sub_conta + str(payment_row)].value - 1
                
                self.bills_plan.accounts[cod_conta].sub_accounts[cod_sub_conta].value -= payments_ws[payments_columns.pmt_value + str(payment_row)].value
        
        #CALCULANDO DESPESAS DA TABELA DE LANCAMENTOS LIVRO CAIXA
        for payment_row in range(1, cashier_ws.max_row):
            is_2plano_contas = cashier_ws[cahsier_columns.is_2plano_contas + str(payment_row)].value == "True"
            is_from_month = common.IsFromMonth(cashier_ws[cahsier_columns.dt_pmt + str(payment_row)].value, self.month, self.year) if is_2plano_contas else False
            
            if is_2plano_contas and is_from_month:
                is_saida = cashier_ws[cahsier_columns.tp_pmt + str(payment_row)].value == "Saída"
                cod_conta = cashier_ws[cahsier_columns.cod_conta + str(payment_row)].value - 1
                cod_sub_conta = cashier_ws[cahsier_columns.cod_sub_conta + str(payment_row)].value - 1
                
                payment_value = cashier_ws[cahsier_columns.pmt_value + str(payment_row)].value
                payment_value = payment_value if is_saida else (-1 * payment_value)
                                
                self.bills_plan.accounts[cod_conta].sub_accounts[cod_sub_conta].value -= payment_value
        
        #CALCULANDO O CUSTO DA MERCADORIA VENDIDA (EXCLUNDO DUPLA CONTAGEM DE FRETE E ICMS)
        self.CalcCMV()
        self.bills_plan.accounts[2].sub_accounts[2].value = 0 #DIFERENCA DE ICMS
        self.bills_plan.accounts[2].sub_accounts[3].value = 0 #FRETE
            
        for account in self.bills_plan.accounts:
            account.CalculateAccount()
            account.participation = account.value / self.bills_plan.accounts[0].value
        
        payments_wb.close()
        cashier_wb.close()
            
    def PrintBillsPlan(self) -> None:
        for account in self.bills_plan.accounts:
            account.PrintAccount()
            
        print(common.FormatPrint('Resultado', self.result, settings.print_size + 4))
        print(common.FormatPrint('Estoque - Custo', self.stock_cost, settings.print_size + 4))