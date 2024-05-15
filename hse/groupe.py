import pandas as pd

def sortie_groupe(df, nom_fichier):
    try:
        data = [] 
        data_unique = []
        for index, row in df.iterrows():
            groupe_all = row["Groupes d'étudiants imposés (noms)"]
            if isinstance(groupe_all, str):
                groupes = groupe_all.split(', ')
                data.extend(groupes)  # Utilisation de extend pour ajouter chaque élément de la liste séparément

        for i in data: 
            if i not in data_unique: 
                data_unique.append(i) 
        data_unique.sort()
        nouveau_df = pd.DataFrame(data_unique, columns=["Groupes"])  # Nom de la colonne pour le DataFrame de sortie
        nouveau_df.to_excel(nom_fichier, index=False)
        print("Les données ont été écrites avec succès dans le fichier Excel de destination :", nom_fichier)
    except Exception as e:
        print("Une erreur s'est produite lors de l'écriture du fichier de sortie :", str(e))


# Assurez-vous de définir votre fonction lire_fichier_entree pour récupérer les données du fichier Excel
def lire_fichier_entree(nom_fichier, feuille):
    try:
        return pd.read_excel(nom_fichier, feuille)
    except Exception as e:
        print("Une erreur s'est produite lors de la lecture du fichier d'entrée :", str(e))
        return None

nom_fichier_entree = "D:\zCour\stage\premier_test\hse\HSE.xlsx"
feuille_entree = "Feuil1"  
df_entree = lire_fichier_entree(nom_fichier_entree, feuille_entree)

if df_entree is not None:
    nom_fichier_sortie = "hse_groupe_sortie.xlsx"
    sortie_groupe(df_entree, nom_fichier_sortie)
