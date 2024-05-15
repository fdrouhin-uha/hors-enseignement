import pandas as pd

#def inserer_donnees_ligne(nom_fichier, feuille, ligne, colonne, donnee):
#    try:
#        df = pd.read_excel(nom_fichier, sheet_name=feuille)
#        nb_colonnes_existantes = len(df.columns)
#        nb_colonnes_a_ajouter = max(0, colonne + len(donnee) - nb_colonnes_existantes)
#        if nb_colonnes_a_ajouter > 0:
#            for i in range(nb_colonnes_a_ajouter):
#                df[f'Nouvelle colonne {nb_colonnes_existantes + i + 1}'] = None
#        
#        for i in range(len(donnee)):
#            df.iloc[ligne-1 + i, colonne-1] = donnee[i]
#        df.to_excel(nom_fichier, index=False, sheet_name=feuille)
#        print("Donnée insérée avec succès dans la cellule ({}, {}) de la feuille '{}'.".format(ligne, colonne, feuille))
#    except Exception as e:
#        print("Une erreur s'est produite lors de l'insertion des données dans le fichier Excel :", str(e))
#
#
#nom_fichier_excel = "donnee.xlsx"
#feuille = "Feuil1"
#ligne = 3
#colonne = 2
#donnee = ["Nouvelle donnée 1", "Nouvelle donnée 2", "Nouvelle donnée 3"]
#inserer_donnees_ligne(nom_fichier_excel, feuille, ligne, colonne, donnee)

import pandas as pd

def inserer_donnees_ligne(nom_fichier, feuille, ligne, colonne_debut, donnees):
    try:
        # Lire le fichier Excel et charger les données dans un DataFrame
        df = pd.read_excel(nom_fichier, sheet_name=feuille)
        
        # Ajouter des colonnes si elles n'existent pas
        nb_colonnes_existantes = len(df.columns)
        nb_colonnes_a_ajouter = max(0, colonne_debut + len(donnees) - nb_colonnes_existantes)
        if nb_colonnes_a_ajouter > 0:
            for i in range(nb_colonnes_a_ajouter):
                df[f'Nouvelle colonne {nb_colonnes_existantes + i + 1}'] = None
        
        # Insérer les données dans la ligne spécifiée
        for i in range(len(donnees)):
            df.iloc[ligne-1, colonne_debut-1 + i] = donnees[i]
        
        # Écrire les données modifiées dans le fichier Excel
        df.to_excel(nom_fichier, index=False, sheet_name=feuille)
        print("Données insérées avec succès dans la ligne {} de la feuille '{}'.".format(ligne, feuille))
    except Exception as e:
        print("Une erreur s'est produite lors de l'insertion des données dans le fichier Excel :", str(e))

# Exemple d'utilisation de la fonction
nom_fichier_excel = "donnee.xlsx"  # Nom du fichier Excel
feuille = "Feuil1"  # Nom de la feuille dans le fichier Excel
ligne = 3  # Numéro de ligne où insérer les données
colonne_debut = 1  # Numéro de la colonne où commencer l'insertion des données
donnees = ["Nouvelle donnée 1", "Nouvelle donnée 2", "Nouvelle donnée 3"]  # Données à insérer
inserer_donnees_ligne(nom_fichier_excel, feuille, ligne, colonne_debut, donnees)
