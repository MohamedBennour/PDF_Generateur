from main import *
import sys
import pyodbc
from PyQt5.QtWidgets import (QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QVBoxLayout, QPushButton,
                             QWidget, QDialog, QLabel, QLineEdit, QDialogButtonBox, QHBoxLayout, QMessageBox)

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
        self.initUI()

    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)

         # Layouts
        main_layout = QVBoxLayout()
        search_layout = QHBoxLayout()

        # Créer les zones de recherche
        self.search_numope = QLineEdit(self)
        self.search_numope.setPlaceholderText("Rechercher par numéro d'opération")
        self.search_datval = QLineEdit(self)
        self.search_datval.setPlaceholderText("Rechercher par date de validation")
        self.search_datech = QLineEdit(self)
        self.search_datech.setPlaceholderText("Rechercher par date de paiement")
        self.search_devach = QLineEdit(self)
        self.search_devach.setPlaceholderText("Rechercher par devise")
        
        # Bouton pour exécuter la recherche
        self.search_button = QPushButton('Rechercher', self)
        self.search_button.clicked.connect(self.on_search_click)

        # Ajouter les éléments de recherche au layout de recherche
        search_layout.addWidget(self.search_numope)
        search_layout.addWidget(self.search_datval)
        search_layout.addWidget(self.search_datech)
        search_layout.addWidget(self.search_devach)
        search_layout.addWidget(self.search_button)
        
        main_layout.addLayout(search_layout)

        # Créer tableau
        self.tableWidget = QTableWidget()
        self.tableWidget.setColumnCount(7)
        self.tableWidget.setHorizontalHeaderLabels([
            "Numéro Opération", "Date Validation", "Montant Net", "Devise",
            "Taux de Marge", "Marge Revenant", "Date de Paiement"
        ])
        main_layout.addWidget(self.tableWidget)

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

        self.load_data_from_db()

    def load_data_from_db(self):
        offset = (self.current_page - 1) * self.records_per_page
        query = "SELECT NUMOPE, DATVAL, MNETAC, DEVACH, TAUXACH, MARACH, DATECH FROM hope ORDER BY NUMOPE OFFSET ? ROWS FETCH NEXT ? ROWS ONLY"
        try:
            with pyodbc.connect('DRIVER={SQL Server}; SERVER=DESKTOP-KIO15OT;DATABASE=BH_BANK; Trusted_Connection=yes;') as conn:
                with conn.cursor() as cursor:
                    cursor.execute(query, (offset, self.records_per_page))
                    records = cursor.fetchall()
                    self.tableWidget.setRowCount(len(records))
                    for row_index, row_data in enumerate(records):
                        for column_index, data in enumerate(row_data):
                            self.tableWidget.setItem(row_index, column_index, QTableWidgetItem(str(data)))
            self.update_page_info()
        except Exception as e:
            QMessageBox.critical(self, "Database Error", f"Error connecting to database: {e}")

    def update_page_info(self):
        self.lbl_page_info.setText(f"Page {self.current_page}")

    def on_search_click(self):
        # Reset to first page on a new search
        self.current_page = 1
        self.load_data_from_db()

    def on_previous_click(self):
        if self.current_page > 1:
            self.current_page -= 1
            self.load_data_from_db()

    def on_next_click(self):
        # Here you might want to check if there are still records to fetch to decide whether to increment page
        self.current_page += 1
        self.load_data_from_db()

    def on_click(self):
        row = self.tableWidget.currentRow()
        if row == -1:
            QMessageBox.warning(self, "Selection Error", "Please select an operation to generate its PDF.")
            return
        num_ope = self.tableWidget.item(row, 0).text()

        buyer_dialog = QDialog()
        buyer_dialog.setWindowTitle("Entrer Nom de l'Acheteur")
        layout = QVBoxLayout()
        label = QLabel("Nom de l'acheteur:")
        self.buyer_name_input = QLineEdit()
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(lambda: self.validate_buyer_name(buyer_dialog, num_ope))
        button_box.rejected.connect(buyer_dialog.reject)
        layout.addWidget(label)
        layout.addWidget(self.buyer_name_input)
        layout.addWidget(button_box)
        buyer_dialog.setLayout(layout)
        buyer_dialog.exec()

    def validate_buyer_name(self, dialog, num_ope):
        buyer_name = self.buyer_name_input.text().strip()
        if buyer_name:
            print(f"Génération du PDF pour l'opération {num_ope} avec l'acheteur {buyer_name}")
            generatePDF(num_ope, buyer_name)
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
