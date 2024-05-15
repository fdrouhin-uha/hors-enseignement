import pandas as pd 

def lire_fichier(chemin_fichier,feuille):
    try:
        donnee_excel = pd.read_excel(chemin_fichier,sheet_name=feuille)
        
        data = []  # Initialiser une liste pour stocker les données
        for index, row in donnee_excel.iterrows():
            groupe = row["Groupes d'étudiants imposés (noms)"]
            data.append(groupe)

        return donnee_excel, data
    except FileNotFoundError:
        print("Le fichier spécifié n'existe pas.")
        return None
    except Exception as e:
        print("Une erreur s'est produite lors de la lecture du fichier Excel :", str(e))
        return None
    


def ecrire_fichier(donnee, nom_fichier):
    try:
        df = pd.DataFrame(donnee)
        df.to_excel(nom_fichier, index=False)
        print("Les données ont été écrites avec succès dans le fichier", nom_fichier)
    except Exception as e:
        print("Une erreur s'est produite lors de l'écriture dans le fichier Excel :", str(e))



def fusionner_colonnes(dataframe, colonne1, colonne2, nouvelle_colonne):
    try:
        dataframe[nouvelle_colonne] = dataframe[colonne1].astype(str) + ' ' + dataframe[colonne2].astype(str)
        if colonne1 in dataframe.columns and colonne2 in dataframe.columns:
            dataframe.drop(columns=[colonne1, colonne2], inplace=True)
        elif colonne1 in dataframe.columns:
            dataframe.drop(columns=[colonne1], inplace=True)
        elif colonne2 in dataframe.columns:
            dataframe.drop(columns=[colonne2], inplace=True)
        return dataframe
    except Exception as e:
        print("Une erreur s'est produite lors de la fusion des colonnes :", str(e))
        return None



def inserer_donnee_cellule(nom_fichier, feuille, ligne, colonne, donnee):
    try:
        df = pd.read_excel(nom_fichier, sheet_name=feuille)
        df.iloc[ligne-1, colonne-1] = donnee
        df.to_excel(nom_fichier, index=False, sheet_name=feuille)
        print("Donnée insérée avec succès dans la cellule ({}, {}) de la feuille '{}'.".format(ligne, colonne, feuille))
    except Exception as e:
        print("Une erreur s'est produite lors de l'insertion des données dans le fichier Excel :", str(e))


#nom_fichier_excel = "donnee.xlsx"
#feuille = "Feuil1"
#ligne = 3
#colonne = 2
#donnee = "Nouvelle donnée"
#inserer_donnee_cellule(nom_fichier_excel, feuille, ligne, colonne, donnee)




chemin_fichier = "HSE.xlsx"
feuille="Feuil1"
donnee, data = lire_fichier(chemin_fichier,feuille)
if donnee is not None:
    print("Contenu du fichier Excel :")
    print(donnee)
    print(data)
#
#
#donnee = {'Nom': ['Alice', 'Bob', 'Charlie'],
#           'Age': [25, 30, 35],
#           'Ville': ['Paris', 'New York', 'Londres']}
#nom_fichier = "donnee.xlsx"
#ecrire_fichier(donnee, nom_fichier)


#donnee = {'Prénom': ['Alice', 'Bob', 'Charlie'],
#           'Nom': ['Dupont', 'Smith', 'Brown'],
#           'Âge': [25, 30, 35]}
#df = pd.DataFrame(donnee)
#print("DataFrame avant fusion :")
#print(df)
#
#df = fusionner_colonnes(df, 'Prénom', 'Nom', 'Nom complet')
#
#print("DataFrame après fusion :")
#print(df)
