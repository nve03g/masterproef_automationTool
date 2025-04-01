import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableView, QVBoxLayout, QPushButton, QWidget, QFileDialog
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
        
        layout = QVBoxLayout()
        layout.addWidget(self.load_button)
        layout.addWidget(self.table_view)
        layout.addWidget(self.save_button)
        
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)
    
    def browse_file(self):
        filepath, _ = QFileDialog.getOpenFileName(self, "Selecteer een bestand", "", "Excel bestanden (*.xlsx *.xls *.xlsm)")
        
        if filepath:
            self.load_excel(filepath)
    
    def load_excel(self, filepath):
        self.df = pd.read_excel(filepath)
        self.model = QStandardItemModel(len(self.df), len(self.df.columns))
        self.model.setHorizontalHeaderLabels(self.df.columns)
        
        for row in range(len(self.df)):
            for col in range(len(self.df.columns)):
                item = QStandardItem(str(self.df.iloc[row, col]))
                self.model.setItem(row, col, item)
        
        self.table_view.setModel(self.model)
        self.save_button.setEnabled(True)
    
    def save_file(self):
        if self.df is not None:
            for row in range(len(self.df)):
                for col in range(len(self.df.columns)):
                    self.df.iloc[row, col] = self.model.item(row, col).text()
            
            save_path, _ = QFileDialog.getSaveFileName(self, "Opslaan als", "", "Excel bestanden (*.xlsx)")
            if save_path:
                self.df.to_excel(save_path, index=False)
                print(f"Bestand opgeslagen als '{save_path}'")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    editor = ExcelEditor()
    editor.show()
    sys.exit(app.exec_())
