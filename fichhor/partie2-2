import sys
import os
from pandas import pandas as pd
from openpyxl import load_workbook
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QFileDialog, QMessageBox

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
        wb.save(nom_fichier)
    except Exception as e:
        print("Une erreur s'est produite lors de l'ajustement de la largeur des colonnes :", str(e))

def ecrire_fichier_sortie(df_code, df_coef, nom_fichier, dossier_sortie):
    try:
        coef = []
        retour = []
        i = 0
        total = 0
        general = []
        generaltot = 0
        general_sans_coef = []
        for index, row in df_coef.iterrows():
            temp = row["Nombre d'heures équivalent TD"]
            coef.append(temp)
        for y in range(len(coef)):
            general.append(0)
            general_sans_coef.append(0)
        for nom_colonne in df_code.columns[1:]:
            df_subset = df_code[['Référentiel', nom_colonne]]
            for index, row in df_subset.iterrows():
                ligne = row[nom_colonne]
                if isinstance(ligne, int):
                    retour.append(ligne * coef[i])
                    total += ligne * coef[i]
                    general[i] += ligne * coef[i]
                    general_sans_coef[i] += ligne
                else:
                    retour.append(0)
                i += 1

            df_subset = df_code[['Référentiel']]
            df_subset = df_subset.assign(temp=retour)
            df_subset = df_subset.rename(columns={'temp': nom_colonne})
            df_subset = pd.concat([df_subset, pd.DataFrame([{"Référentiel": "Total", nom_colonne: total}])], ignore_index=True)
            nom_fichier_sortie = os.path.join(dossier_sortie, f"{nom_fichier.split('.')[0]}_{nom_colonne}.xlsx")
            df_subset.to_excel(nom_fichier_sortie, index=False, sheet_name='Feuil1')
            ajuster_largeur_colonnes(nom_fichier_sortie)
            print("Le fichier Excel de sortie a été créé avec succès :", nom_fichier_sortie)

            generaltot += total
            i = 0
            total = 0
            retour = []

        df_total = df_code[['Référentiel']]
        df_total = df_total.assign(total=general)
        df_total = df_total.assign(Total=general_sans_coef)
        df_total = pd.concat([df_total, pd.DataFrame([{"Référentiel": "total", "total": generaltot}])], ignore_index=True)
        fichier_total_path = os.path.join(dossier_sortie, "fichier_total.xlsx")
        df_total.to_excel(fichier_total_path, index=False, sheet_name='Feuil1')
        ajuster_largeur_colonnes(fichier_total_path)
        print("Les données ont été écrites avec succès dans les fichiers Excel de destination.")
    except Exception as e:
        print("Une erreur s'est produite lors de l'écriture des fichiers de sortie :", str(e))

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Traitement des fichiers de référentiel et liste d\'enseignants')

        # Layout principal
        layout = QVBoxLayout()

        # Ligne pour sélectionner le fichier de référentiel
        self.referenciel_layout = QHBoxLayout()
        self.referenciel_label = QLabel('Fichier de référentiel:')
        self.referenciel_edit = QLineEdit()
        self.referenciel_button = QPushButton('Parcourir...')
        self.referenciel_button.clicked.connect(self.browse_referenciel)
        self.referenciel_layout.addWidget(self.referenciel_label)
        self.referenciel_layout.addWidget(self.referenciel_edit)
        self.referenciel_layout.addWidget(self.referenciel_button)
        layout.addLayout(self.referenciel_layout)

        # Ligne pour sélectionner la feuille du fichier de référentiel
        self.feuille_ref_layout = QHBoxLayout()
        self.feuille_ref_label = QLabel('Feuille de référentiel:')
        self.feuille_ref_edit = QLineEdit('Feuil1')
        self.feuille_ref_layout.addWidget(self.feuille_ref_label)
        self.feuille_ref_layout.addWidget(self.feuille_ref_edit)
        layout.addLayout(self.feuille_ref_layout)

        # Ligne pour sélectionner le fichier de liste d'enseignants
        self.enseignants_layout = QHBoxLayout()
        self.enseignants_label = QLabel('Fichier de liste d\'enseignants:')
        self.enseignants_edit = QLineEdit()
        self.enseignants_button = QPushButton('Parcourir...')
        self.enseignants_button.clicked.connect(self.browse_enseignants)
        self.enseignants_layout.addWidget(self.enseignants_label)
        self.enseignants_layout.addWidget(self.enseignants_edit)
        self.enseignants_layout.addWidget(self.enseignants_button)
        layout.addLayout(self.enseignants_layout)

        # Ligne pour sélectionner la feuille du fichier de liste d'enseignants
        self.feuille_ens_layout = QHBoxLayout()
        self.feuille_ens_label = QLabel('Feuille de liste d\'enseignants:')
        self.feuille_ens_edit = QLineEdit('Feuil1')
        self.feuille_ens_layout.addWidget(self.feuille_ens_label)
        self.feuille_ens_layout.addWidget(self.feuille_ens_edit)
        layout.addLayout(self.feuille_ens_layout)

        # Ligne pour sélectionner le dossier de sortie
        self.dossier_sortie_layout = QHBoxLayout()
        self.dossier_sortie_label = QLabel('Dossier de sortie:')
        self.dossier_sortie_edit = QLineEdit()
        self.dossier_sortie_button = QPushButton('Parcourir...')
        self.dossier_sortie_button.clicked.connect(self.browse_dossier_sortie)
        self.dossier_sortie_layout.addWidget(self.dossier_sortie_label)
        self.dossier_sortie_layout.addWidget(self.dossier_sortie_edit)
        self.dossier_sortie_layout.addWidget(self.dossier_sortie_button)
        layout.addLayout(self.dossier_sortie_layout)

        # Bouton pour exécuter le traitement
        self.run_button = QPushButton('Exécuter')
        self.run_button.clicked.connect(self.run_processing)
        layout.addWidget(self.run_button)

        self.setLayout(layout)

    def browse_referenciel(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self, "Sélectionner le fichier de référentiel", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
        if fileName:
            self.referenciel_edit.setText(fileName)

    def browse_enseignants(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self, "Sélectionner le fichier de liste d'enseignants", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
        if fileName:
            self.enseignants_edit.setText(fileName)

    def browse_dossier_sortie(self):
        options = QFileDialog.Options()
        folder = QFileDialog.getExistingDirectory(self, "Sélectionner le dossier de sortie", options=options)
        if folder:
            self.dossier_sortie_edit.setText(folder)

    def run_processing(self):
        referenciel_file = self.referenciel_edit.text()
        feuille_ref = self.feuille_ref_edit.text()
        enseignants_file = self.enseignants_edit.text()
        feuille_ens = self.feuille_ens_edit.text()
        dossier_sortie = self.dossier_sortie_edit.text()

        if not os.path.exists(referenciel_file) or not os.path.exists(enseignants_file):
            QMessageBox.critical(self, "Erreur", "Veuillez sélectionner des fichiers valides.")
            return

        if not dossier_sortie:
            dossier_sortie = "sortie"

        if not os.path.exists(dossier_sortie):
            os.makedirs(dossier_sortie)

        df_ref = lire_fichier_entree(referenciel_file, feuille_ref)
        df_ens = lire_fichier_entree(enseignants_file, feuille_ens)

        if df_ref is not None and df_ens is not None:
            nom_fichier = "fichier_total.xlsx"
            ecrire_fichier_sortie(df_ens, df_ref, nom_fichier, dossier_sortie)
            QMessageBox.information(self, "Terminé", "Le traitement a été effectué avec succès.")

def main():
    app = QApplication(sys.argv)
    mainWindow = MainWindow()
    mainWindow.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
