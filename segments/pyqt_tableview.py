import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableView, QVBoxLayout, QPushButton, QWidget
from PyQt5.QtGui import QStandardItemModel, QStandardItem

class ExcelEditor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Data Editor")
        self.setGeometry(100, 100, 800, 600)

        # Laad Excel data
        self.df = pd.read_excel("AlarmList_file_ingevuld.xlsx")

        # Zet de DataFrame om naar een model
        self.model = QStandardItemModel(len(self.df), len(self.df.columns))
        self.model.setHorizontalHeaderLabels(self.df.columns)

        for row in range(len(self.df)):
            for col in range(len(self.df.columns)):
                item = QStandardItem(str(self.df.iloc[row, col]))
                self.model.setItem(row, col, item)

        # Maak de table view
        self.table_view = QTableView(self)
        self.table_view.setModel(self.model)
        self.table_view.setEditTriggers(QTableView.AllEditTriggers)
        self.table_view.setSelectionMode(QTableView.SingleSelection)
        
        # Opslaan knop
        self.save_button = QPushButton("Opslaan als nieuwe Excel", self)
        self.save_button.clicked.connect(self.save_file)

        # Layout
        layout = QVBoxLayout()
        layout.addWidget(self.table_view)
        layout.addWidget(self.save_button)
        
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def save_file(self):
        # Zet de gewijzigde data terug naar een DataFrame
        for row in range(len(self.df)):
            for col in range(len(self.df.columns)):
                self.df.iloc[row, col] = self.model.item(row, col).text()

        # Sla het gewijzigde DataFrame op als een nieuw Excel-bestand
        self.df.to_excel("pyqt_tableview_result.xlsx", index=False)
        print("Bestand opgeslagen als 'pyqt_tableview_result.xlsx'")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    editor = ExcelEditor()
    editor.show()
    sys.exit(app.exec_())
