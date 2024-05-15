import pandas as pd

def sortie_groupe(df, nom_fichier):
    try:
        data_groupe = []
        data_ens = []
        data_ens_unique = []
        data_groupe_unique = []
        retour = []
        test = []
        heures_agreg = {} 
        for index, row in df.iterrows():
            groupe_all = row["Groupes d'étudiants imposés (noms)"]
            ens_all = row["Enseignants"]
            type = row["Type"]
            heures = row["Durée (h)"]
        
            heures_nombre = convertir_heure_en_nombre(heures)

            if isinstance(groupe_all, str):
                groupes = groupe_all.split(', ')
                data_groupe.extend(groupes)
            if type in ["TD", "TP", "CM"]:
                if isinstance(ens_all, str):
                    ens = ens_all.split(', ')
                    data_ens.extend(ens)

        for i in data_groupe: 
            if i not in data_groupe_unique: 
                data_groupe_unique.append(i)
        
        for i in data_ens: 
            if i not in data_ens_unique: 
                data_ens_unique.append(i) 
        
        for i in data_groupe_unique:
            for j in data_ens_unique:
                retour.append({'Groupe':i, 'Enseignants': j})
                a = [i,j]
                test.append(a)
        z=0
        for index, row in df.iterrows():
            groupe_all = row["Groupes d'étudiants imposés (noms)"]
            ens_all = row["Enseignants"]
            type_enseignement = row["Type"]
            heures = row["Durée (h)"]
            z+=1
            
            for i in test:
                if type_enseignement in ["TD", "TP", "CM"]:
                    # Créer une clé basée sur [groupe, enseignant]
                    temp1,temp2=i
                    key = (temp1,temp2)
                    print(key)
                    # Vérifier si la combinaison [groupe, enseignant] existe déjà dans le dictionnaire
                    if key in heures_agreg:
                        # Ajouter les heures à la valeur existante
                        heures_agreg[key] += heures_nombre
                    else:
                        # Initialiser les heures à cette valeur si elle n'existe pas encore
                        heures_agreg[key] = heures_nombre
            if z == 2000:
                break
        # Créer une liste pour stocker les données agrégées
        rendu = []
        for key, heures in heures_agreg.items():
            groupe, enseignant = key
            heures_totales, minutes = divmod(heures, 60)
            heure_formattee = '{:02d}h{:02d}'.format(heures_totales, minutes)
            rendu.append({'Groupe': groupe, 'Enseignant': enseignant, 'Heure': heure_formattee})
        
        data_groupe_unique.sort()
        nouveau_df = pd.DataFrame(rendu)
        nouveau_df.to_excel(nom_fichier, index=False)
        print("Les données ont été écrites avec succès dans le fichier Excel de destination :", nom_fichier)
    except Exception as e:
        print("Une erreur s'est produite lors de l'écriture du fichier de sortie :", str(e))

def convertir_heure_en_nombre(heure_str):
    # Séparer l'heure et les minutes en fonction de 'h'
    heures, minutes = heure_str.split('h')
    # Convertir les heures et les minutes en nombres entiers
    heures = int(heures)
    minutes = int(minutes)
    # Calculer le total en minutes
    total_minutes = heures * 60 + minutes
    return total_minutes

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
    nom_fichier_sortie = "hse_groupe_enseignant_sortie.xlsx"
    sortie_groupe(df_entree, nom_fichier_sortie)
