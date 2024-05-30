import sys
import os
import pandas as pd
from openpyxl import load_workbook
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QFileDialog, QMessageBox


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle('Traitement des fichiers de code étape et liste d\'enseignants')

        layout = QVBoxLayout()

        self.code_etape_layout = QHBoxLayout()
        self.code_etape_label = QLabel('Fichier de code étape:')
        self.code_etape_edit = QLineEdit()
        self.code_etape_button = QPushButton('Parcourir...')
        self.code_etape_button.clicked.connect(self.recherche_code_etape)
        self.code_etape_layout.addWidget(self.code_etape_label)
        self.code_etape_layout.addWidget(self.code_etape_edit)
        self.code_etape_layout.addWidget(self.code_etape_button)
        layout.addLayout(self.code_etape_layout)

        self.feuille_code_layout = QHBoxLayout()
        self.feuille_code_label = QLabel('Feuille de code étape:')
        self.feuille_code_edit = QLineEdit('Feuil1')
        self.feuille_code_layout.addWidget(self.feuille_code_label)
        self.feuille_code_layout.addWidget(self.feuille_code_edit)
        layout.addLayout(self.feuille_code_layout)
        
        self.enseignants_layout = QHBoxLayout()
        self.enseignants_label = QLabel('Fichier de liste d\'enseignants:')
        self.enseignants_edit = QLineEdit()
        self.enseignants_button = QPushButton('Parcourir...')
        self.enseignants_button.clicked.connect(self.recherche_enseignants)
        self.enseignants_layout.addWidget(self.enseignants_label)
        self.enseignants_layout.addWidget(self.enseignants_edit)
        self.enseignants_layout.addWidget(self.enseignants_button)
        layout.addLayout(self.enseignants_layout)

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

        self.run_button = QPushButton('Exécuter')
        self.run_button.clicked.connect(self.run_processing)
        layout.addWidget(self.run_button)

        self.setLayout(layout)
    
    def recherche_code_etape(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self, "Sélectionner le fichier de code étape", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
        if fileName:
            self.code_etape_edit.setText(fileName)
    
    def recherche_enseignants(self):
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
        code_etape_file = self.code_etape_edit.text()
        feuille_code = self.feuille_code_edit.text()
        enseignants_file = self.enseignants_edit.text()
        feuille_ens = self.feuille_ens_edit.text()
        dossier_sortie = self.dossier_sortie_edit.text()

        if not os.path.exists(code_etape_file) or not os.path.exists(enseignants_file):
            QMessageBox.critical(self, "Erreur", "Veuillez sélectionner des fichiers valides.")
            return

        if not dossier_sortie:
            dossier_sortie = "sortie/referenciel"
        else:
            dossier_sortie = dossier_sortie+"/referenciel"

        if not os.path.exists(dossier_sortie):
            
            os.makedirs(dossier_sortie)

        df_ens = lire_fichier_entree(enseignants_file, feuille_ens)
        df_entree = lire_fichier_entree(code_etape_file, feuille_code)

        if df_entree is not None:
            nom_fichier_horaire = os.path.join(dossier_sortie, "fichor_sortie.xlsx")
            nom_fichier_horaire_ens = os.path.join(dossier_sortie, "fichor_ens_sortie.xlsx")
            ecrire_fichier_sortie(df_entree, df_ens, nom_fichier_horaire, nom_fichier_horaire_ens)

        QMessageBox.information(self, "Terminé", "Le traitement a été effectué avec succès.")

def lire_fichier_entree(nom_fichier, feuille):
    try:
        df = pd.read_excel(nom_fichier, sheet_name=feuille)
        return df
    except Exception as e:
        print("Une erreur s'est produite lors de la lecture du fichier d'entrée :", str(e))
        return None

def ecrire_fichier_sortie(df_code, df_ens, nom_fichier, nom_fich_ens):
    try:
        data = []
        for index, row in df_code.iterrows():
            code_etape = row['Code Etape']
            intitule_code_etape = row['Intitulé code étape']
            referenciel = intitule_code_etape
            if isinstance(code_etape, str):
                for col in df_code.columns:
                    if col not in ['Code Etape', 'Intitulé code étape'] and col in row:
                        heures = row[col] if pd.notnull(row[col]) else 0
                        if heures != 0:  # Ignorer les lignes avec 0 ou NaN
                            data.append({
                                'Référentiel': f"{referenciel} - {col}", 
                                'Code étape': code_etape, 
                                'Nombre d\'heures équivalent TD': heures, 
                                'Remarque': f'heures par {col.lower()}'
                            })
        data.append({
            'Référentiel': "Loi ORE : mettre 1 si l'enseignant a étudié des dossiers, 0 sinon", 
            'Code étape': " ", 
            'Nombre d\'heures équivalent TD': 0.5, 
            'Remarque': "h pour relecture du dossier"
        })
        
        nouveau_df = pd.DataFrame(data)
        nouveau_df.to_excel(nom_fichier, index=False, sheet_name='Feuil1')

        ajuster_largeur_colonnes(nom_fichier)

        data_ens = cree_ens(nouveau_df, df_ens)

        nouveau_df_ens = pd.DataFrame(data_ens)
        nouveau_df_ens = nouveau_df_ens.fillna('')
        nouveau_df_ens = nouveau_df_ens.transpose()
        nouveau_df_ens.reset_index(inplace=True)
        nouveau_df_ens = nouveau_df_ens.rename(columns={'index': 'Référentiel'})
        nouveau_df_ens.to_excel(nom_fich_ens, index=False, sheet_name='Feuil1')

        ajuster_largeur_colonnes(nom_fich_ens)
        
        print("Les données ont été écrites avec succès dans le fichier Excel de destination :", nom_fichier)
        print("La fiche d'enseignant a été écrite dans le fichier de destination", nom_fich_ens)
    except Exception as e:
        print("Une erreur s'est produite lors de l'écriture du fichier de sortie :", str(e))

def cree_ens(nouveau_df, df_ens):
    try:
        data = {}
        for index, row in nouveau_df.iterrows():
            referenciel = row['Référentiel']
            if isinstance(referenciel, str):
                if referenciel not in data:
                    data[referenciel] = {}
                for index, row_ens in df_ens.iterrows():
                    enseignant = ""
                    if isinstance(row_ens['NOM'], str) or isinstance(row_ens['Prénom'], str):
                        enseignant = f"{row_ens['NOM']} {row_ens['Prénom']}"
                    if enseignant not in data[referenciel]:
                        data[referenciel][enseignant] = ""
        return data
    except Exception as e:
        print("Une erreur s'est produite lors de la creation du fichier enseignants", str(e))

def ajuster_largeur_colonnes(nom_fichier):
    try:
        wb = load_workbook(nom_fichier)
        ws = wb.active
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = max_length + 2
            ws.column_dimensions[column_letter].width = adjusted_width
        wb.save(nom_fichier)
    except Exception as e:
        print("Une erreur s'est produite lors de l'ajustement de la largeur des colonnes :", str(e))

def main():
    app = QApplication(sys.argv)
    mainWindow = MainWindow()
    mainWindow.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
