from main import *
import sys
import pandas as pd
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QVBoxLayout, QPushButton,
                             QWidget, QDialog, QLabel, QLineEdit, QDialogButtonBox, QHBoxLayout, QMessageBox, QHeaderView)

class App(QMainWindow):
    def __init__(self):
        super().__init__()
        self.title = 'Générateur de PDF'
        self.left = 100
        self.top = 100
        self.width = 800
        self.height = 600
        self.current_page = 1
        self.records_per_page = 10
        self.dataFrame = pd.DataFrame() 
        self.load_data_from_excel()  
        self.initUI()
        self.load_data_to_table()

    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)

        # Layouts
        main_layout = QVBoxLayout()
        search_layout = QHBoxLayout()

        # Créer les zones de recherche
        self.search_dossier = QLineEdit(self)
        self.search_dossier.setPlaceholderText("Rechercher par numéro de dossier")
        self.search_datval = QLineEdit(self)
        self.search_datval.setPlaceholderText("Rechercher par date de validation")
        self.search_datech = QLineEdit(self)
        self.search_datech.setPlaceholderText("Rechercher par date de paiement")
        self.search_devach = QLineEdit(self)
        self.search_devach.setPlaceholderText("Rechercher par devise")

        self.search_dossier.returnPressed.connect(self.on_search_click)
        self.search_datval.returnPressed.connect(self.on_search_click)
        self.search_datech.returnPressed.connect(self.on_search_click)
        self.search_devach.returnPressed.connect(self.on_search_click)
        
        # Bouton pour exécuter la recherche
        self.search_button = QPushButton('Rechercher', self)
        self.search_button.clicked.connect(self.on_search_click)

        # Ajouter les éléments de recherche au layout de recherche
        search_layout.addWidget(self.search_dossier)
        search_layout.addWidget(self.search_datval)
        search_layout.addWidget(self.search_datech)
        search_layout.addWidget(self.search_devach)
        search_layout.addWidget(self.search_button)
        
        main_layout.addLayout(search_layout)

        # Créer tableau
        self.tableWidget = QTableWidget()
        self.tableWidget.setColumnCount(7)
        self.tableWidget.setHorizontalHeaderLabels([
            "Numéro Dossier", "Date Validation", "Montant Net", "Devise",
            "Taux de Marge", "Marge Revenant", "Date de Paiement"
        ])
        main_layout.addWidget(self.tableWidget)
        self.tableWidget.setSortingEnabled(True)
        self.tableWidget.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tableWidget.setSelectionBehavior(QTableWidget.SelectRows)
        self.tableWidget.setSelectionMode(QTableWidget.SingleSelection)
        self.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        pagination_layout = QHBoxLayout()
        self.btn_previous = QPushButton('Précédent', self)
        self.btn_next = QPushButton('Suivant', self)
        self.lbl_page_info = QLabel(self)
        self.btn_previous.clicked.connect(self.on_previous_click)
        self.btn_next.clicked.connect(self.on_next_click)
        pagination_layout.addWidget(self.btn_previous)
        pagination_layout.addStretch(1)
        pagination_layout.addWidget(self.lbl_page_info)
        pagination_layout.addStretch(1)
        pagination_layout.addWidget(self.btn_next)
        main_layout.addLayout(pagination_layout)

        self.button_generate_pdf = QPushButton('Générer PDF', self)
        self.button_generate_pdf.clicked.connect(self.on_click)
        main_layout.addWidget(self.button_generate_pdf)

        central_widget = QWidget()
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)
        self.show()

    def max_pages(self):
        # Retourne le nombre total de pages
        total_records = len(self.filtered_dataFrame)
        return (total_records + self.records_per_page - 1) // self.records_per_page
    
    def update_page_info(self):
        # Met à jour le label avec le numéro de la page actuelle et le nombre total de pages
        total_pages = self.max_pages()
        self.lbl_page_info.setText(f"Page {self.current_page} de {total_pages}")
        self.btn_previous.setEnabled(self.current_page > 1)
        self.btn_next.setEnabled(self.current_page < total_pages)

    def load_data_from_excel(self):
        directory = 'Rapport'
        self.original_dataFrame = pd.DataFrame()  # Stocker les données originales
        for filename in os.listdir(directory):
            if filename.endswith('.xlsx'):
                file_path = os.path.join(directory, filename)
                df = pd.read_excel(file_path)
                self.original_dataFrame = pd.concat([self.original_dataFrame, df], ignore_index=True)
        self.filtered_dataFrame = self.original_dataFrame.copy()  # Utiliser pour filtrage et affichage


    def load_data_to_table(self):
        self.update_page_info()  # Assurez-vous de mettre à jour l'info de pagination
        start = (self.current_page - 1) * self.records_per_page
        end = start + self.records_per_page
        current_data = self.filtered_dataFrame.iloc[start:end]

        self.tableWidget.setRowCount(len(current_data))
        column_names = ['DOSSIER', 'DATVAL', 'MNETAC', 'DEVACH', 'TAUXACH', 'MARACH', 'DATECH']
        for row_index, row_data in enumerate(current_data.itertuples(index=False)):
            for column_index, column_name in enumerate(column_names):
                data = getattr(row_data, column_name)
                self.tableWidget.setItem(row_index, column_index, QTableWidgetItem(str(data)))
        self.update_page_info()  # Mettre à jour l'info de pagination après chargement des données

    def on_search_click(self):
        # Réinitialiser les données filtrées
        self.filtered_dataFrame = self.original_dataFrame.copy()
        # Appliquer les filtres
        filters = {
            'DOSSIER': self.search_dossier.text(),
            'DATVAL': self.search_datval.text(),
            'DATECH': self.search_datech.text(),
            'DEVACH': self.search_devach.text()
        }
        mask = pd.Series(True, index=self.filtered_dataFrame.index)
        for col, value in filters.items():
            if value:
                mask &= self.filtered_dataFrame[col].astype(str).str.contains(value, case=False, na=False)
        self.filtered_dataFrame = self.filtered_dataFrame[mask]
        self.current_page = 1  # Réinitialiser la pagination
        self.load_data_to_table()


    def on_previous_click(self):
        if self.current_page > 1:
            self.current_page -= 1
            self.load_data_to_table()

    def on_next_click(self):
        self.current_page += 1
        self.load_data_to_table()

    def on_click(self):
        row = self.tableWidget.currentRow()
        if row == -1:
            QMessageBox.warning(self, "Selection Error", "Please select an operation to generate its PDF.")
            return
        num_dossier = self.tableWidget.item(row, 0).text()

        buyer_dialog = QDialog()
        buyer_dialog.setWindowTitle("Entrer Nom de l'Acheteur")
        layout = QVBoxLayout()
        label = QLabel("Nom de l'acheteur:")
        self.buyer_name_input = QLineEdit()
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(lambda: self.validate_buyer_name(buyer_dialog, num_dossier))
        button_box.rejected.connect(buyer_dialog.reject)
        layout.addWidget(label)
        layout.addWidget(self.buyer_name_input)
        layout.addWidget(button_box)
        buyer_dialog.setLayout(layout)
        buyer_dialog.exec()

    def validate_buyer_name(self, dialog, num_dossier):
        buyer_name = self.buyer_name_input.text().strip()
        if buyer_name:
            generatePDF(num_dossier, buyer_name)
            dialog.accept()
        else:
            error_dialog = QMessageBox()
            error_dialog.setWindowTitle("Erreur de saisie")
            error_dialog.setText("Le nom de l'acheteur ne peut pas être vide.")
            error_dialog.setIcon(QMessageBox.Warning)
            error_dialog.exec()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())
