import pandas as pd


class ExcelReader:
    
    def __init__(self, list_of_entreprises : list[str]):
        self.entreprises = list_of_entreprises
        self.entreprises_contacts : dict =  self.load_entreprises_contact()
        return None
    
    
    
    
    def load_entreprises_contact(self) -> dict:
        """ return dict like [ENTREPRISE-NAME]-> DATAFRAME """
        entreprises_contacts = {}
        if(self.entreprises != []) :
            for entreprise in self.entreprises :
                entreprises_contacts[entreprise] = pd.read_excel(f"./data/contacts/{entreprise} - contacts.xlsx")
        
        return entreprises_contacts
   