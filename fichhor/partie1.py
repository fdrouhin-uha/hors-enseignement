import pandas as pd

def lire_fichier(chemin_fichier,feuille):
    try:
        donnee_excel = pd.read_excel(chemin_fichier, sheet_name=feuille)
        return donnee_excel
    except FileNotFoundError:
        print("Le fichier spécifié n'existe pas.")
        return None
    except Exception as e:
        print("Une erreur s'est produite lors de la lecture du fichier Excel :", str(e))
        return None
    
def recuperer_code_etape(chemin_fichier, feuille):
    try:
        df = pd.read_excel(chemin_fichier, sheet_name=feuille)
        code_etape, intitule_code_etape = df.columns[0:2]
        return code_etape, intitule_code_etape
    except Exception as e:
        print("Une erreur s'est produite lors de la récupération des deux premières colonnes :", str(e))
        return None
    
def recuperer_coeficient(chemin_fichier, feuille):
    try:
        df = pd.read_excel(chemin_fichier, sheet_name=feuille)
        stage, SAE_responsable, SAE_suivi, apprenti  = df.columns[2:4]
        return stage, SAE_responsable, SAE_suivi, apprenti
    except Exception as e:
        print("Une erreur s'est produite lors de la récupération des deux premières colonnes :", str(e))
        return None

def lire_fichier_entree(nom_fichier, feuille):
    try:
        df = pd.read_excel(nom_fichier, sheet_name=feuille)
        return df
    except Exception as e:
        print("Une erreur s'est produite lors de la lecture du fichier d'entrée :", str(e))
        return None
        
def ecrire_fichier_sortie(df, nom_fichier):
    try:
        nouveau_df = pd.DataFrame(columns=['Référentiel', 'Code étape'])
        for index, row in df.iterrows():
            code_etape = row['Code Etape']
            intitule_code_etape = row['Intitulé code étape']
            for colonne in ['Stage', 'SAE responsable', 'SAE suivi', 'Apprenti']:
                valeur = row[colonne]
                if pd.notnull(valeur):
                    referenciel = f"{intitule_code_etape}_{colonne.replace(' ', '_')}"
                    nouveau_df = nouveau_df.append({'Référentiel': referenciel, 'Code étape': code_etape}, ignore_index=True)
        nouveau_df.to_excel(nom_fichier, index=False)
        print("Les données ont été écrites avec succès dans le fichier Excel de destination :", nom_fichier)
    except Exception as e:
        print("Une erreur s'est produite lors de l'écriture du fichier de sortie :", str(e))

nom_fichier_entree = "donnees_entree.xlsx"  
feuille_entree = "Feuil1"  
df_entree = lire_fichier_entree(nom_fichier_entree, feuille_entree)

if df_entree is not None:
    nom_fichier_sortie = "fichier_sortie.xlsx"  
    ecrire_fichier_sortie(df_entree, nom_fichier_sortie)