def init():
    global sheets_folder
    sheets_folder = "Acompanhamento/"
    global document_file
    document_file = "GERAL - Documentos.xlsx"
    global payment_file
    payment_file = "GERAL - Contas a pagar.xlsx"
    global recieve_file
    recieve_file = "GERAL - Contas a receber.xlsx"
    global cashier_file
    cashier_file = "GERAL - Lancamentos Livro Caixa.xlsx"
    global export_file
    export_file = "Resultado Hidraucenter Nordeste.xlsx"
    global print_size
    print_size = 50
class document_columns():
    def __init__(self) -> None:
        self.tp_doc = 'A'
        self.status_doc = 'C'
        self.dt_doc = 'I'
        self.pmt_dinheiro = 'V'
        self.pmt_cheque = 'W'
        self.pmt_transferencia = 'X'
        self.pmt_cartao = 'Y'
        self.pmt_faturado = 'Z'
        self.pmt_financeira = 'AA'
class payment_columns():
    def __init__(self) -> None:
        self.status_pmt = 'H'
        self.is_2plano_contas = 'M'
        self.cod_conta = 'N'
        self.cod_sub_conta = 'O'
        self.dt_pmt = 'P'
        self.pmt_value = 'S'
class cashier_columns():
    def __init__(self) -> None:
        self.dt_pmt = 'C'
        self.tp_pmt = 'D'
        self.is_2plano_contas = 'H'
        self.cod_conta = 'I'
        self.cod_sub_conta = 'J'
        self.pmt_value = 'G'
class stockCCV_columns():
    def __init__(self) -> None:
        self.amount = 'C'
        self.cost = 'E'
class profits_columns():
    def __init__(self) -> None:
        self.cost_prod = 'E'
class export_sheets():
    def __init__(self) -> None:
        self.geral = 'GERAL'
        self.participation = 'Participações'