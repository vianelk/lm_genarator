from docx import Document
import copy
import pandas as pd


class LetterGenerator:
    
    def __init__(self, cover_letter_path : str,skills_entreprises : str) -> None:
        self.cover_letter_path = cover_letter_path
        self.skills_entreprises = skills_entreprises
        pass
    
    def generate_cover_letter(self, entreprise_name : str, entreprise_contact : pd.DataFrame):
        
        for index, rows in entreprise_contact.iterrows():
            
            document = self.replace_letter_values(doc= Document(self.cover_letter_path),
                                                  adress=rows["LIEU"],
                                                  entreprise_name=entreprise_name,
                                                  fonction=rows["FONCTION"],
                                                  name=rows["NOM"],
                                                  sexe=rows["SEXE"])

            document.save(f"./data/lettres de motivation/{entreprise_name}/ {rows["NOM"]} - lettre de motivation.docx")
            
            first_name, _= rows["NOM"].split(" ")
            self.generate_email(fichier=f"./data/lettres de motivation/{entreprise_name}/email.txt",nom=first_name,skills=self.skills_entreprises[entreprise_name], entreprise=entreprise_name)
        
        
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
    
    def generate_email(self,fichier, nom, entreprise, skills):
        nouveau_texte = f"""
CANDIDATURE DATA ENGINEER EN ALTERNANCE CHEZ ORANGE
        
Bonjour {nom},
J'espère que vous allez bien.

Je me permets de vous contacter pour vous proposer ma candidature pour une alternance en data engineering au sein de {entreprise}.

Actuellement étudiant en 4ème année Big data & AI,
je suis passionné par le domaine du big data et je souhaite développer mes compétences en rejoignant une équipe innovante telle que la vôtre.
Mes expériences sur les technologies cloud et data ({skills} etc...) me permettront de contribuer efficacement à vos projets.
Je vous joins mon CV, ma lettre de motivation ainsi qu'une lettre de recommandation pour appuyer ma candidature.
Je suis disponible immédiatement pour tout entretien afin de vous exposer mes motivations.

Dans l'attente de votre retour, je vous prie d'agréer l'expression de mes salutations distinguées.
Vianel KONG
0758334658
        
        --------------------------------------------------------------------------------------------------------------
        
        
        """

        try:
            with open(fichier, 'a') as file:  # Ouvrir le fichier en mode ajout ('a' pour append)
                file.write(nouveau_texte + '\n')  # Écrire le nouveau texte suivi d'un saut de ligne
        except FileNotFoundError:
            print(f"Le fichier {fichier} n'existe pas.")
        except Exception as e:
            print(f"Une erreur est survenue: {e}")