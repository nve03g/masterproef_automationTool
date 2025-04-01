import sys
import json
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableView, QVBoxLayout, QPushButton, QWidget, QFileDialog, QComboBox
from PyQt5.QtGui import QStandardItemModel, QStandardItem

class ExcelEditor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Data Editor")
        self.setGeometry(100, 100, 800, 600)
        
        self.df = None
        self.model = QStandardItemModel()
        self.table_view = QTableView(self)
        self.table_view.setModel(self.model)
        self.table_view.setEditTriggers(QTableView.AllEditTriggers)
        self.table_view.setSelectionMode(QTableView.SingleSelection)
        
        self.load_button = QPushButton("Laad Excel-bestand", self)
        self.load_button.clicked.connect(self.browse_file)
        
        self.save_button = QPushButton("Opslaan als nieuwe Excel", self)
        self.save_button.clicked.connect(self.save_file)
        self.save_button.setEnabled(False)  # Alleen activeren als er een bestand geladen is
        
        self.dropdown1 = QComboBox(self)
        self.dropdown2 = QComboBox(self)
        self.dropdown1.setVisible(False)
        self.dropdown2.setVisible(False)
        
        layout = QVBoxLayout()
        layout.addWidget(self.load_button)
        layout.addWidget(self.dropdown1)
        layout.addWidget(self.dropdown2)
        layout.addWidget(self.table_view)
        layout.addWidget(self.save_button)
        
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)
        
        self.load_config()
    
    def load_config(self):
        try:
            with open("config.json", "r") as file:
                self.config = json.load(file)
        except Exception as e:
            print(f"Fout bij het laden van config.json: {e}")
            self.config = {"profiles": []}
    
    def browse_file(self):
        filepath, _ = QFileDialog.getOpenFileName(self, "Selecteer een bestand", "", "Excel bestanden (*.xlsx *.xls *.xlsm)")
        
        if filepath:
            self.load_excel(filepath)
            self.load_dropdown_data(filepath)
    
    def load_excel(self, filepath):
        self.df = pd.read_excel(filepath, sheet_name=None)  # Lees alle sheets in
        first_sheet = list(self.df.keys())[0]  # Standaard de eerste sheet tonen
        self.model = QStandardItemModel(len(self.df[first_sheet]), len(self.df[first_sheet].columns))
        self.model.setHorizontalHeaderLabels(self.df[first_sheet].columns)
        
        for row in range(len(self.df[first_sheet])):
            for col in range(len(self.df[first_sheet].columns)):
                item = QStandardItem(str(self.df[first_sheet].iloc[row, col]))
                self.model.setItem(row, col, item)
        
        self.table_view.setModel(self.model)
        self.save_button.setEnabled(True)
        self.dropdown1.setVisible(True)
        self.dropdown2.setVisible(True)
    
    def load_dropdown_data(self, filepath):
        # Laad profielen uit config.json
        self.dropdown1.clear()
        self.dropdown1.addItems(self.config.get("profiles", []))
        
        # Laad sheetnamen uit het geladen Excel-bestand
        self.dropdown2.clear()
        self.dropdown2.addItems(self.df.keys())
    
    def save_file(self):
        if self.df is not None:
            current_sheet = self.dropdown2.currentText()
            for row in range(len(self.df[current_sheet])):
                for col in range(len(self.df[current_sheet].columns)):
                    self.df[current_sheet].iloc[row, col] = self.model.item(row, col).text()
            
            save_path, _ = QFileDialog.getSaveFileName(self, "Opslaan als", "", "Excel bestanden (*.xlsx)")
            if save_path:
                with pd.ExcelWriter(save_path) as writer:
                    for sheet_name, sheet_data in self.df.items():
                        sheet_data.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"Bestand opgeslagen als '{save_path}'")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    editor = ExcelEditor()
    editor.show()
    sys.exit(app.exec_())
