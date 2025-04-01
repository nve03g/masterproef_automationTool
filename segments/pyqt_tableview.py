import sys
import json
import os
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
        
        self.profile_dropdown = QComboBox(self)
        self.profile_dropdown.setVisible(False)
        
        self.sheet_dropdown = QComboBox(self)
        self.sheet_dropdown.setVisible(False)
        self.sheet_dropdown.currentIndexChanged.connect(self.update_table_view)
        
        layout = QVBoxLayout()
        layout.addWidget(self.load_button)
        layout.addWidget(self.profile_dropdown)
        layout.addWidget(self.sheet_dropdown)
        layout.addWidget(self.table_view)
        layout.addWidget(self.save_button)
        
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)
        
        self.load_config()
    
    def load_config(self):
        try:
            with open("pyqt_tableview_config.json", "r") as file:
                self.config = json.load(file)
        except Exception as e:
            print(f"Fout bij het laden van pyqt_tableview_config.json: {e}")
            self.config = {"profiles": []}
    
    def browse_file(self):
        filepath, _ = QFileDialog.getOpenFileName(self, "Selecteer een bestand", "", "Excel bestanden (*.xlsx *.xls *.xlsm)")
        
        if filepath:
            self.current_file = os.path.basename(filepath)
            self.load_excel(filepath)
            self.load_dropdown_data(filepath)
    
    def load_excel(self, filepath):
        self.df = pd.read_excel(filepath, sheet_name=None)  # Lees alle sheets in
        self.all_sheets = list(self.df.keys())
        self.sheet_dropdown.clear()
        self.sheet_dropdown.addItems(self.all_sheets)
        self.profile_dropdown.setVisible(True)
        self.sheet_dropdown.setVisible(True)
        self.sheet_dropdown.setCurrentIndex(0)
        self.update_table_view()
        self.save_button.setEnabled(True)
        
    def update_table_view(self):
        selected_sheet = self.sheet_dropdown.currentText()
        if selected_sheet and self.df:
            sheet_data = self.df[selected_sheet]
            self.model = QStandardItemModel(len(sheet_data), len(sheet_data.columns))
            self.model.setHorizontalHeaderLabels(sheet_data.columns)
            
            for row in range(len(sheet_data)):
                for col in range(len(sheet_data.columns)):
                    item = QStandardItem(str(sheet_data.iloc[row, col]))
                    self.model.setItem(row, col, item)
            
            self.table_view.setModel(self.model)
            
    def load_dropdown_data(self, filepath):
        # Laad profielen uit pyqt_tableview_config.json
        self.profile_dropdown.clear()
        self.profile_dropdown.addItems(self.config.get("profiles", []))
        
        # # Laad sheetnamen uit het geladen Excel-bestand
        # self.sheet_dropdown.clear()
        # self.sheet_dropdown.addItems(self.df.keys())
    
    def save_file(self):
        if self.df is not None:
            current_sheet = self.sheet_dropdown.currentText()
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
