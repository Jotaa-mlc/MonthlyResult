from common import FormatPrint

class Sub_account:
    def __init__(self, _description:str):
        self.description = _description
        self.value = 0.0
class Account:
    def __init__(self, _description:str, _sub_accounts:list[Sub_account]):
        self.description = _description
        self.value = 0.0
        self.participation = 0.0
        self.sub_accounts = _sub_accounts
        
    def CalculateAccount(self) -> None:
        self.value = 0.0
        for sub_account in self.sub_accounts:
            self.value += sub_account.value
    
    def PrintAccount(self) -> None:
        account_value = f'{self.value:_.2f}'.replace('.', ',').replace('_', '.')
        account_participation = f'{self.participation*100:_.2f}'.replace('.', ',').replace('_', '.')
        print(f"{self.description} - R$ {account_value} - {account_participation}%")
        
        for sub_account in self.sub_accounts:
            print(' ' * 4 + FormatPrint(sub_account.description, sub_account.value))