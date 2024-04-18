import common

class Sub_account:
    def __init__(self, _description:str, _value:float = 0):
        self.description = _description
        self.value = _value
class Account:
    def __init__(self, _description:str, _sub_accounts:list[Sub_account], _value:float = 0):
        self.description = _description
        self.value = _value
        self.sub_accounts = _sub_accounts
        
    def CalculateAccount(self) -> float:
        self.value = 0
        for sub_account in self.sub_accounts:
            self.value += sub_account.value
        return self.value
    
    def PrintAccount(self) -> None:
        account_value = f"{self.value:_.2f}"
        account_value = account_value.replace('.', ',').replace('_', '.')
        print(f"{self.description} - R$ {account_value}")
        
        for sub_account in self.sub_accounts:
            print(' ' * 4 + common.FormatPrint(sub_account.description, sub_account.value))