import pandas as pd

def lire_fichier_entree(nom_fichier, feuille):
    try:
        df = pd.read_excel(nom_fichier, sheet_name=feuille)
        return df
    except Exception as e:
        print("Une erreur s'est produite lors de la lecture du fichier d'entrée :", str(e))
        return None

def ecrire_fichier_sortie(df, nom_fichier):
    try:
        data = []  # Initialiser une liste pour stocker les données
        for index, row in df.iterrows():
            code_etape = row['Code Etape']
            intitule_code_etape = row['Intitulé code étape']
            if 'FA' in intitule_code_etape:
                referenciel = intitule_code_etape
            else:
                referenciel = f"{intitule_code_etape}_Stage"
            for colonne in ['Stage', 'SAE responsable', 'SAE suivi', 'Apprenti']:
                valeur = row[colonne]
                if pd.notnull(valeur):
                    # Convertir les valeurs float en chaînes de caractères
                    valeur = str(valeur)
                    if 'heures' in colonne or 'heure' in colonne:
                        remarque = f"{valeur} heures par {colonne.split()[0]}"
                    else:
                        remarque = ''
                    # Ajouter les données à la liste
                    data.append({'Référentiel': referenciel, 'Code étape': code_etape, 'Nombre d\'heures équivalent TD': valeur, 'Remarque': remarque})

        # Créer un DataFrame à partir de la liste de données
        nouveau_df = pd.DataFrame(data)

        # Écrire le DataFrame dans un fichier Excel
        nouveau_df.to_excel(nom_fichier, index=False)
        print("Les données ont été écrites avec succès dans le fichier Excel de destination :", nom_fichier)
    except Exception as e:
        print("Une erreur s'est produite lors de l'écriture du fichier de sortie :", str(e))


nom_fichier_entree = "Codes étapes TC.xlsx" 
feuille_entree = "Feuil1"  
df_entree = lire_fichier_entree(nom_fichier_entree, feuille_entree)

if df_entree is not None:
    nom_fichier_sortie = "fichier_sortie.xlsx"
    ecrire_fichier_sortie(df_entree, nom_fichier_sortie)
