import sys
import json
import warnings
import pandas as pd
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QComboBox, QPushButton, QTableView, QFileDialog, QFrame, QLabel, QScrollArea
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QStandardItemModel, QStandardItem

# REMARK: last data row isn't correctly calculated

# We get a warning that openpyxl no longer supports dropdown lists in Excel, 
# but that's not a problem because we’re performing that check 
# through the python code, so this warning may be ignored.
warnings.simplefilter("ignore", UserWarning) 


class ExcelTableView(QWidget):
    """ 
    GUI class showing dropdown for profile selection, 
    sheet selection and table view to display dataframe. 
    """
    def __init__(self, processor):
        super().__init__()
        self.processor = processor
        
        self.setWindowTitle("Pfizer Automation Tool")
        self.setGeometry(100, 100, 1000, 800)  # It's important to define window size!

        # layout
        main_layout = QVBoxLayout(self)
        controls_layout = QHBoxLayout()

        # Dropdown menu for profile selection.
        self.profile_dropdown = QComboBox(self)
        self.profile_dropdown.addItems(list(self.processor.profiles.keys()))
        self.profile_dropdown.currentTextChanged.connect(self.update_profile)
        controls_layout.addWidget(QLabel("Select profile:"))
        controls_layout.addWidget(self.profile_dropdown)
        
        # File browse button.
        self.file_button = QPushButton("Browse", self)
        self.file_button.clicked.connect(self.browse_file)
        controls_layout.addWidget(self.file_button)

        # Dropdown menu for sheet selection.
        self.sheet_dropdown = QComboBox(self)
        self.sheet_dropdown.currentTextChanged.connect(self.update_sheet)
        controls_layout.addWidget(QLabel("Select sheet:"))
        controls_layout.addWidget(self.sheet_dropdown)
        
        main_layout.addLayout(controls_layout)
        
        # treeview widget (tableview in pyqt5)
        self.table_view = QTableView(self)
        self.table_model = QStandardItemModel(self)
        self.table_view.setModel(self.table_model)
        main_layout.addWidget(self.table_view)

        # Show initial data in the table view.
        self.load_tableview()
        
    def browse_file(self):
        filepath, _ = QFileDialog.getOpenFileName(self, "Open Excel file", "", "Excel files (*.xlsx *.xls)")
    
        if filepath:
            self.processor.set_filepath(filepath)
            self.update_profile()
        
    def update_profile(self):
        """ Update user profile and load correct data into table view. """
        new_profile = self.profile_dropdown.currentText()
        # Store currently selected sheet.
        current_sheet = self.sheet_dropdown.currentText()
        self.processor.set_profile(new_profile)
        # Update sheet dropdown options based on selected profile.
        self.update_sheet_options(current_sheet)
        
    def update_sheet_options(self, previous_sheet=None):
        """ Update the available sheet options in dropdown list according to current profile and try to keep the previous selection. """
        # Get the sheets for current user profile out of config file.
        available_sheets = self.processor.get_config_sheets()
        self.sheet_dropdown.clear()
        self.sheet_dropdown.addItems(available_sheets)
        
        # Keep previously selected sheet open if it's still available, otherwise fall back to first available sheet.
        if previous_sheet in available_sheets:
            self.sheet_dropdown.setCurrentText(previous_sheet)
        elif available_sheets:
            self.sheet_dropdown.setCurrentText(available_sheets[0])
            
        # Automatically load the selected sheet.
        self.update_sheet()
        
    def update_sheet(self, event=None):
        """ Update the table view with data from selected sheet. """
        sheet_name = self.sheet_dropdown.currentText()
        self.load_tableview(sheet_name)
        
    def load_tableview(self, sheet_name=None):
        """ Reload table view with correct sheet, columns and data. """
        # Delete the current table view data.
        self.table_model.clear()
        
        if sheet_name is None:
            # By default take the current selected sheet.
            sheet_name = self.sheet_dropdown.currentText()
        
        # Get dataframe for selected sheet.
        df = self.processor.get_dataframe(sheet_name)
        if df is not None:
            # Set headers.
            self.table_model.setHorizontalHeaderLabels(list(df.columns))
            
            # Add rows to model
            for _, row in df.iterrows():
                row_items = [QStandardItem(str(cell)) for cell in row.tolist()]
                self.table_model.appendRow(row_items)
        
        
class ExcelProcessor:
    """
    initialize Excel file processor
    - config_path : str, path to JSON config file
    - profile : str, huidige hardcoded profielaanwijzing (supplier, developer, operator)
    """
    def __init__(self, config_path, profile):
        self.load_config(config_path)
        # Dictionary containing all data {sheetname: DataFrame}.
        self.dataframes = {}
        self.profile = profile
        self.filepath = None  # User always has to choose a file before tool works.
        
    def load_config(self, config_path):
        """ Load configuration parameters from JSON file. """
        with open(config_path, 'r', encoding="utf-8") as f:
            config = json.load(f)
        
        self.default_profile = config["default_profile"]
        self.headerrows = config["header_rows"]
        self.index_start = config["index_start"]
        self.columnnames = config["column_names"]
        self.profiles = config["profiles"]
        
    def set_filepath(self, filepath):
        self.filepath = filepath
        self.load_excel()        
    
    def get_config_sheets(self):
        """ Return list of available sheets for current profile. """
        available_sheets = []
        for sheet in self.headerrows.keys():
            if sheet in self.profiles[self.profile]:
                available_sheets.append(sheet)
        return available_sheets
        
    def set_profile(self, profile):
        """ Set active profile (OR default profile). """
        if profile not in self.profiles:
            self.profile = self.default_profile
        else:
            self.profile = profile
                
    def load_excel(self):
        """ Load the Excel sheets into dataframes, with the specified header row per sheet and columns allowed depending on the indicated profile. """
        # Don't try to load an Excel file when the file has not been selected yet.
        if not self.filepath:
            return
        
        # Open the Excel-file.
        xls = pd.ExcelFile(self.filepath)
        
        for sheet, header_row in self.headerrows.items():
            if sheet in xls.sheet_names:
                df = pd.read_excel(self.filepath, sheet_name=sheet, header=header_row-1)
                
                # Only use profile-assigned columns.
                allowed_columns = self.profiles[self.profile][sheet]
                valid_columns = [col for col in allowed_columns if col in df.columns]
                df = df[valid_columns]
                    
                if sheet in self.index_start:
                    # Delete excess rows that are right below the header row(s).
                    if (self.index_start[sheet] - self.headerrows[sheet] - 2) >= 0:
                        # We have to delete the first x rows in the dataframe.
                        # -2+1 because else range(0,0), then we get an empty list (row 0 doesn't get dropped from the dataframe)
                        df = df.drop([i for i in range(self.index_start[sheet]-self.headerrows[sheet]-1)])
                    
                    # Adjust indices to match Excel row indeces.
                    df.index = range(self.index_start[sheet], self.index_start[sheet] + len(df))                                 
                    
                self.dataframes[sheet] = df
            else:
                print(f"Warning: {sheet} not found in {self.filepath}.")
        
    def get_dataframe(self, sheetname):
        """ Get a specific sheet (DataFrame). """
        return self.dataframes.get(sheetname)
    


if __name__ == "__main__":
    config_file = "config.json"
    current_profile = "operator"  # Will later be set through GUI dropdown list.
    processor = ExcelProcessor(config_file, profile=current_profile)
    
    app = QApplication(sys.argv)
    window = ExcelTableView(processor)
    window.show()
    sys.exit(app.exec_())


# I use these lines to print a list of all column names for a specific sheet in the right syntax, to be able to put it correctly in the config file.
# df_alarmlist = processor.get_dataframe("Alarmlist")
# print(list(df_alarmlist.columns.values))

# df_cp = processor.get_dataframe("Color Pictures")
# print(list(df_cp.columns.values))
