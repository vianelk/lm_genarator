from excelReader import ExcelReader
from letterGenetor import LetterGenerator

def main():
    domain = "TELECOM"
    entreprises_list =  ["ORANGE","SFR","BOUYGUES TELECOM"]
    entreprises_skills =  {
        "ORANGE":"CGP, Python, Spark/Scala,Hive, Hadoop",
        "SFR":"CGP, Python, Hadoop",
        "BOUYGUES TELECOM":" Python, Hadoop, Hive, Impala, GCP"
    }
    entreprises_contacts =  ExcelReader(entreprises_list).entreprises_contacts
    lg = LetterGenerator(f"./data/lettres de motivation/Lettre de motivation - {domain}.docx",skills_entreprises = entreprises_skills)
    
    for entreprise in entreprises_list :
        lg.generate_cover_letter(entreprise_name=entreprise,entreprise_contact=entreprises_contacts[entreprise])
        print("done !")

if __name__ == "__main__":
    main()
