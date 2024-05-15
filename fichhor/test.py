from pandas import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import  get_column_letter


def lire_fichier_entree(nom_fichier, feuille):
    try:
        df = pd.read_excel(nom_fichier, sheet_name=feuille)
        return df
    except Exception as e:
        print("Une erreur s'est produite lors de la lecture du fichier d'entrée :", str(e))
        return None


def ajuster_largeur_colonnes(nom_fichier):
    try:
        wb = load_workbook(nom_fichier)
        ws = wb.active
        for column in ws.columns:
            max_length = 0
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            ws.column_dimensions[column[0].column_letter].width = adjusted_width
        wb.save(nom_fichier)  # N'oubliez pas de sauvegarder le classeur Excel après l'ajustement
    except Exception as e:
        print("Une erreur s'est produite lors de l'ajustement de la largeur des colonnes :", str(e))


def ecrire_fichier_sortie(df_code, df_coef,nom_fichier):
    try:
        coef=[]
        retour=[]
        i=0
        total=0
        general = []
        for index, row in df_coef.iterrows():
                temp=row["Nombre d'heures équivalent TD"]
                coef.append(temp)

        for nom_colonne in df_code.columns[1:]:
            df_subset = df_code[['Référentiel', nom_colonne]]

            for index, row in df_subset.iterrows():
                ligne=row[nom_colonne]
                if isinstance(ligne, int):
                    retour.append(ligne*coef[i])
                    total += ligne*coef[i]
                else:
                    retour.append(0)
                i+=1

            df_subset = df_code[['Référentiel']]
            df_subset = df_subset.assign(temp=retour)
            df_subset = df_subset.rename(columns={'temp': nom_colonne})
            df_subset = pd.concat([ df_subset, pd.DataFrame([{"Référentiel": "Total", nom_colonne: total}])], ignore_index=True)
            nom_fichier_sortie = nom_fichier.split('.')[0] + f"_{nom_colonne}.xlsx"
            df_subset.to_excel(nom_fichier_sortie, index=False)
            ajuster_largeur_colonnes(nom_fichier_sortie)
            print("Le fichier Excel de sortie a été créé avec succès :", nom_fichier_sortie)
            i=0
            total = 0
            retour=[]
        #df_subset = df_code[['Référentiel']]
        #df_subset = df_subset.assign(temp=retour)
        print("Les données ont été écrites avec succès dans les fichiers Excel de destination.")
    except Exception as e:
        print("Une erreur s'est produite lors de l'écriture des fichiers de sortie :", str(e))


fichier_ens = "D:/zCour/stage/premier_test/fichor_sortie.xlsx"
feuille_entree_ens = "Sheet1"
df_coeficien = lire_fichier_entree(fichier_ens, feuille_entree_ens)


nom_fichier_entree = "D:/zCour/stage/premier_test/fichor_ens_sortie.xlsx" 
feuille_entree = "Sheet1"  
df_entree = lire_fichier_entree(nom_fichier_entree, feuille_entree)


if df_entree is not None:
    nom_fichier_horaire = "fichor_sortie.xlsx"
    ecrire_fichier_sortie(df_entree, df_coeficien, nom_fichier_horaire)