from docx import Document
import copy
import pandas as pd


class LetterGenerator:
    
    def __init__(self, cover_letter_path : str) -> None:
        self.cover_letter_path = cover_letter_path
        pass
    
    def generate_cover_letter(self, entreprise_name : str, entreprise_contact : pd.DataFrame):
        
        for index, rows in entreprise_contact.iterrows():
            
            document = self.replace_letter_values(doc= Document(self.cover_letter_path),
                                                  adress=rows["LIEU"],
                                                  entreprise_name=entreprise_name,
                                                  fonction=rows["FONCTION"],
                                                  name=rows["NOM"],
                                                  sexe=rows["SEXE"])

            document.save(f"./data/lettres de motivation/{entreprise_name}/lettre de motivation - {rows["NOM"]}.docx")
        
        
        return ''
    
    def replace_letter_values(self,doc :Document,entreprise_name :str,name :str,adress:str,fonction:str,sexe :str) -> Document:
        
        for para in doc.paragraphs:
            voyelles = ['a', 'e', 'i', 'o', 'u', 'y', 'A', 'E', 'I', 'O', 'U', 'Y']
            article =  "d'" if entreprise_name[0] in voyelles else 'de '
            article2 = "Chère " if sexe == "F" else "Cher "
            first_name, second_name = name.split(" ")
            self.replace_text_in_paragraph(para,"[ENTREPRISE]",entreprise_name)
            self.replace_text_in_paragraph(para,"[ARTICLE + ENTREPRISE]",article + entreprise_name)
            self.replace_text_in_paragraph(para,"[NOM]", first_name.capitalize() + " " + second_name.upper())
            self.replace_text_in_paragraph(para,"[ADRESSE]",adress)
            self.replace_text_in_paragraph(para,"XXX",article2 + fonction)
            self.replace_text_in_paragraph(para,"XX",(article2.lower() + fonction))
            self.replace_text_in_paragraph(para,"[FONCTION]",fonction)
             
        
        return doc
    def replace_text_in_paragraph(self,paragraph, old_text, new_text):
        for run in paragraph.runs:
            if old_text in run.text:
               
                run.text = run.text.replace(old_text, new_text)