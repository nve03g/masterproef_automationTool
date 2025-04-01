import os
import sys
import json
import logging
import warnings
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableView, QVBoxLayout, QPushButton, QWidget, QFileDialog, QComboBox
from PyQt5.QtGui import QStandardItemModel, QStandardItem

# REMARK: last data row isn't correctly calculated

# Log errors in logfile

logging.basicConfig(
    filename="error_log.txt",
    level=logging.ERROR,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

# Extra: log fouten ook naar de terminal
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.ERROR)
console_handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
logging.getLogger().addHandler(console_handler)

# We get a warning that openpyxl no longer supports dropdown lists in Excel, 
# but that's not a problem because we’re performing that check 
# through the python code, so this warning may be ignored.
warnings.simplefilter("ignore", UserWarning) 
# this warning is possibly suppressing other relevant UserWarnings

class ExcelEditor(QMainWindow):
    """ 
    GUI class showing file browse button, dropdown for profile selection, 
    sheet selection, table view to display dataframe and buttong to save as new file. 
    """
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Pfizer Automation Tool")
        self.setGeometry(100, 100, 800, 600)  # It's important to define window size!

        self.df = None
        self.model = QStandardItemModel()
        self.table_view = QTableView(self)
        self.table_view.setModel(self.model)
        self.table_view.setEditTriggers(QTableView.AllEditTriggers)
        self.table_view.setSelectionMode(QTableView.SingleSelection)
        
        self.load_button = QPushButton("Load Excel file", self)
        self.load_button.clicked.connect(self.browse_file)
        
        self.save_button = QPushButton("Save as new Excel file", self)
        self.save_button.clicked.connect(self.save_file)
        self.save_button.setEnabled(False)  # Only activate button if a file is selected
        
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
            with open("config.json", "r") as file:
                self.config = json.load(file)
        except Exception as e:
            print(f"Error loading config.json: {e}")
            self.config = {"profiles": []}
        
    def browse_file(self):
        filepath, _ = QFileDialog.getOpenFileName(self, "Select Excel file", "", "Excel files (*.xlsx *.xls *.xlsm)")
            
        if filepath:
            self.current_file = os.path.basename(filepath)
            self.load_excel(filepath)
            self.load_dropdown_data(filepath)

    def load_excel(self, filepath):
        self.df = pd.read_excel(filepath, sheet_name=None)  # Load all sheets
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
        # Load profiles from config.json
        self.profile_dropdown.clear()
        self.profile_dropdown.addItems(self.config.get("profiles", []))
        
        # # Load sheetnames from the selected Excel file
        # self.sheet_dropdown.clear()
        # self.sheet_dropdown.addItems(self.df.keys())
    
    def save_file(self):
        if self.df is not None:
            current_sheet = self.sheet_dropdown.currentText()
            for row in range(len(self.df[current_sheet])):
                for col in range(len(self.df[current_sheet].columns)):
                    self.df[current_sheet].iloc[row, col] = self.model.item(row, col).text()
            
            save_path, _ = QFileDialog.getSaveFileName(self, "Save as", "", "Excel files (*.xlsx)")
            if save_path:
                with pd.ExcelWriter(save_path) as writer:
                    for sheet_name, sheet_data in self.df.items():
                        sheet_data.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"File saved as '{save_path}'")
        

if __name__ == "__main__":
    app = QApplication(sys.argv)
    editor = ExcelEditor()
    editor.show()
    sys.exit(app.exec_())
    