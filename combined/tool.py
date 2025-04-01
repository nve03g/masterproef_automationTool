import os
import sys
import json
import logging
import warnings
import pandas as pd
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QComboBox, QPushButton, QTableView, QFileDialog, QLabel
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
        
        # self.profile_dropdown.addItems(list(self.processor.profiles.keys()))
        # Verzamel unieke profielen uit de config
        unique_profiles = set()
        for profiles in self.processor.profiles.values():
            unique_profiles.update(profiles.keys())  # Verzamel unieke profielen uit alle bestanden

        self.profile_dropdown.addItems(list(unique_profiles))

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
        # # Ask to browse file before opening tool.
        # self.browse_file()
        
    def browse_file(self):
        excel_path, _ = QFileDialog.getOpenFileName(self, "Open Excel file", "", "Excel files (*.xlsx *.xls *.xlsm)")
        
        print(f"Geselecteerd bestand: {excel_path}")  # Debugging
    
        if excel_path:
            self.processor.set_excel_filepath(excel_path)
            self.update_profile()

    def update_profile(self):
        """ Update user profile and load correct data into table view. """
        new_profile = self.profile_dropdown.currentText()
        self.processor.set_profile(new_profile)
        
        # Store previous selected sheet.
        previous_sheet = self.sheet_dropdown.currentText()
        # Update sheet dropdown options based on selected profile.
        self.update_sheet_options(previous_sheet)
        
        # Refresh the table view with the new profile's rights.
        self.update_sheet()
        
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
        # self.profile = profile
        self.excel_path = None  # User always has to choose a file before tool works.
        
    def load_config(self, config_path):
        """ Load configuration parameters from JSON file. """
        with open(config_path, 'r', encoding="utf-8") as f:
            config = json.load(f)
        
        self.default_profile = config["default_profile"]
        self.files = config["files"]
        
        self.headerrows = {}
        self.index_start = {}
        self.columnnames = {}
        self.profiles = {}
        
        # Loop over all files in the config file.
        for filename, file_config in self.files.items():
            self.headerrows[filename] = {}
            self.index_start[filename] = {}
            self.columnnames[filename] = {}
            self.profiles[filename] = {}
            
            for sheetname, sheet_config in file_config["sheets"].items():
                self.headerrows[filename][sheetname] = sheet_config["header_row"]
                self.index_start[filename][sheetname] = sheet_config["index_start"]
                self.columnnames[filename][sheetname] = sheet_config["columns"]
                
            # Verwerk profielen
            for profile, access in file_config["profiles"].items():
                if access == "ALL":
                    # Als 'ALL', dan krijgt dit profiel toegang tot alle kolommen van alle sheets
                    self.profiles[filename][profile] = {
                        sheet: sheet_config["columns"] for sheet, sheet_config in file_config["sheets"].items()
                    }
                else:
                    self.profiles[filename][profile] = access
                
    def set_excel_filepath(self, excel_path):
        self.excel_path = excel_path
        try:
            self.load_excel()        
        except Exception as e:
            logging.exception(f"Error loading Excel file: {self.excel_path}")
    
    def get_config_sheets(self):
        """ Return list of available sheets for current profile. """
        if not self.excel_path:
            return []
        
        filename = os.path.basename(self.excel_path)
        if filename not in self.files:
            return []
        
        if self.profile not in self.profiles[filename]:
            return []
        
        if self.profiles[filename][self.profile] == "ALL":
            return list(self.headerrows[filename].keys())  # Admin kan alle sheets zien
        
        # available_sheets = []
        
        # for sheet in self.headerrows.keys():
        #     if sheet in self.profiles[self.profile]:
        #         available_sheets.append(sheet)
        # return available_sheets
        return list(self.profiles[filename][self.profile].keys())
        
    def set_profile(self, profile):
        # """ Set active profile (OR default profile). """
        # if profile not in self.profiles:
        #     self.profile = self.default_profile
        # else:
        #     self.profile = profile
        """ Set the active profile for data access. """
        self.profile = profile
                
    def load_excel(self):
        """ Load the Excel sheets into dataframes, with the specified header row per sheet and columns allowed depending on the indicated profile. """
        # Don't try to load an Excel file when the file has not been selected yet.
        if not self.excel_path:
            return
        
        filename = os.path.basename(self.excel_path)
        
        try:
            if filename not in self.files:
                print(f"Warning: {filename} is not in the config.")
            
            # Open the Excel-file.
            xls = pd.ExcelFile(self.excel_path)
            
            for sheet, sheet_config in self.files[filename]["sheets"].items():
                if sheet in xls.sheet_names:
                    header_row = self.headerrows[filename][sheet]
                    df = pd.read_excel(self.excel_path, sheet_name=sheet, header=header_row-1)
                    
                    # Only use profile-assigned columns.
                    # Controleer of het profiel 'ALL' mag zien (alle kolommen, dus admin profiel).
                    if self.profiles[self.profile] == "ALL":
                        valid_columns = df.columns.tolist()  # Neem alle kolommen, dus niet nodig om alle kolommen in config te zetten dan?
                    else:
                        # allowed_columns = self.profiles[self.profile][sheet]
                        allowed_columns = self.profiles[filename][self.profile].get(sheet, [])
                        valid_columns = [col for col in allowed_columns if col in df.columns]
                    
                    df = df[valid_columns]
                        
                    if sheet in self.index_start[filename]:
                        index_start = self.index_start[filename][sheet]
                        # Delete excess rows that are right below the header row(s).
                        if (index_start - header_row - 2) >= 0:
                            # We have to delete the first x rows in the dataframe.
                            # -2+1 because else range(0,0), then we get an empty list (row 0 doesn't get dropped from the dataframe)
                            df = df.drop([i for i in range(index_start-header_row-1)])
                        
                        # Adjust indices to match Excel row indeces.
                        df.index = range(index_start, index_start + len(df))                                 
                        
                    self.dataframes[sheet] = df
                else:
                    logging.warning(f"Warning: {sheet} not found in {self.excel_path}.")
            
        except Exception as e:
            logging.exception(f"Error loading Excel file: {self.excel_path}")
        
    def get_dataframe(self, sheetname):
        """ Get a specific sheet (DataFrame). """
        return self.dataframes.get(sheetname)
    


if __name__ == "__main__":
    config_file = "new_config_V2.json"
    current_profile = "operator"  # Will later be set through GUI dropdown list.
    processor = ExcelProcessor(config_file, profile=current_profile)
        
    try:
        app = QApplication(sys.argv)
        window = ExcelTableView(processor)
        window.show()
        sys.exit(app.exec_())

    except Exception as e:
        logging.exception("Unexpected error in the GUI.")

# I use these lines to print a list of all column names for a specific sheet in the right syntax, to be able to put it correctly in the config file.
# df_alarmlist = processor.get_dataframe("Alarmlist")
# print(list(df_alarmlist.columns.values))

# df_cp = processor.get_dataframe("Color Pictures")
# print(list(df_cp.columns.values))
