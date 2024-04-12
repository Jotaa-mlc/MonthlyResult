class Account:
    def __init__(self, _description:str, _value:float = 0):
        self.description = _description
        self.value = _value
        self.sub_accounts = list()
        
    def AddSub_account(self, _description:str):
        self.sub_accounts.append(Sub_account(_description))
        
    def CalculateAcount(self):
        self.value = 0
        for sub_account in self.sub_accounts:
            self.value += sub_account.value
        return self.value

class Sub_account:
    def __init__(self, _description:str, _value:float = 0) -> None:
        self.description = _description
        self.value = _value