from docx import Document
from excelReader import ExcelReader
from letterGenetor import LetterGenerator

def main():
    domain = "TELECOM"
    entreprises_list =  ["ORANGE"]
    entreprises_contacts =  ExcelReader(entreprises_list).entreprises_contacts
    lg = LetterGenerator(f"./data/lettres de motivation/Lettre de motivation - {domain}.docx")
    
    for entreprise in entreprises_list :
        lg.generate_cover_letter(entreprise_name=entreprise,entreprise_contact=entreprises_contacts[entreprise])
        print("done !")

if __name__ == "__main__":
    main()
