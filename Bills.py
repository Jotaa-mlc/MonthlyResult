class Sub_account:
    def __init__(self, _description:str, _value:float = 0):
        self.description = _description
        self.value = _value
class Account:
    def __init__(self, _description:str, _sub_accounts:list[Sub_account], _value:float = 0):
        self.description = _description
        self.value = _value
        self.sub_accounts = _sub_accounts
        
    def CalculateAcount(self) -> float:
        self.value = 0
        for sub_account in self.sub_accounts:
            self.value += sub_account.value
        return self.value