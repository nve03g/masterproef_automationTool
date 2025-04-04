import pandas as pd
import warnings
from collections import defaultdict

warnings.simplefilter("ignore", UserWarning) # we krijgen warning dat openpyxl geen dropdownlijsten in excel meer ondersteunt, maar dat is geen probleem want die controle ga ik via mijn python code uitvoeren, dus deze warning mag genegeerd worden

class ExcelProcessor:
    def __init__(self, filepath, headerrows, columnnames):
        """
        initialize Excel file processor
        - filepath : str
        - headerrows : {sheetname1: headerrow1, ...}, dictionary containing sheetnames and header-row-index
        - columnnames : {sheetname1: [columnname1, columnname2, ...], ...}, dictionary containing sheetnames and a list of column names to load
        """
        self.filepath = filepath
        self.headerrows = headerrows
        self.columnnames = columnnames
        self.dataframes = {} # dictionary containing all data {sheetname: DataFrame}
        
    def load_excel(self):
        # self.dataframes = pd.read_excel(self.filepath, sheet_name=self.sheetnames, header=self.headerrows)
        """ Laad de Excel-sheets in dataframes, met de opgegeven header-rij per sheet. """
        xls = pd.ExcelFile(self.filepath) # Open het Excel-bestand
        
        for sheet, header_row in self.headerrows.items():
            if sheet in xls.sheet_names:
                df = pd.read_excel(self.filepath, sheet_name=sheet, header=header_row-1)
                
                # Filter enkel de gewenste kolommen als ze bestaan in de DataFrame
                if sheet in self.columnnames:
                    valid_columns = [col for col in self.columnnames[sheet] if col in df.columns]
                    df = df[valid_columns]
                    
                self.dataframes[sheet] = df
            else:
                print(f"Waarschuwing: {sheet} niet gevonden in {self.filepath}")
        
    def get_dataframe(self, sheetname):
        """ get specific sheet (DataFrame) """
        return self.dataframes.get(sheetname)
    
    
class DataValidator:
    def __init__(self, dataframes):
        """
        initialize data validator with a dictionary of dataframes
        - dataframes : {sheetname: DataFrame, ...}
        """
        self.dataframes = dataframes
        self.errors = []
        
    def max_characters(self, sheetname, columnname, max_chars):
        """ check max amount of characters allowed in a column """
        if sheetname not in self.dataframes:
            print(f"Waarschuwing: Sheet '{sheetname}' niet gevonden.")
            return
        
        df = self.dataframes[sheetname]
        
        if columnname not in df.columns:
            print(f"Waarschuwing: Kolom '{columnname}' niet gevonden in '{sheetname}'.")
            return
        
        # perform check
        for index, value in df[columnname].dropna().items():
            # print((index, value))
            if isinstance(value, str) and len (value) > max_chars:
                error_msg = f"Rij {index+4}: '{value}' is te lang ({len(value)} > {max_chars})"
                self.errors.append((sheetname, columnname, index+4, error_msg))
    
    def alarm_exists(self, row, alarm_column_name="Alarmtext machine constructor (German)"):
        """ checks if alarm exists in given row, returns boolean """
        if alarm_column_name not in row.index:
            print(f"Waarschuwing: Kolom '{alarm_column_name}' niet gevonden.")
            return False
        
        value = str(row[alarm_column_name]).strip() if pd.notna(row[alarm_column_name]) else ""
        no_alarm_values = {"reserved", "gereserveerd", "spare", ""} # config
        
        return value.lower() not in no_alarm_values # True if alarm exists
    
    def file_type(self, sheetname, columnname, alarm_column_name="Alarmtext machine constructor (German)", file_extension=".pdl"):
        """ Controleert of waarden in een bepaalde kolom eindigen op een correct bestandstype. """
        if sheetname not in self.dataframes:
            print(f"Waarschuwing: Sheet '{sheetname}' niet gevonden.")
            return
        
        df = self.dataframes[sheetname]
        
        if columnname not in df.columns:
            print(f"Waarschuwing: Kolom '{columnname}' niet gevonden in '{sheetname}'.")
            return
        
        if alarm_column_name not in df.columns:
            print(f"Waarschuwing: Alarm-kolom '{alarm_column_name}' niet gevonden in '{sheetname}'.") # controle op bestaan van alarm kan niet worden uitgevoerd
            return

        # Controle per rij
        for index, row in df.iterrows():
            if self.alarm_exists(row, alarm_column_name):  # Controle: alleen als alarm bestaat
                cell_value = str(row[columnname]).strip() if pd.notna(row[columnname]) else ""
                
                if not cell_value.endswith(file_extension):
                    error_msg = f"Rij {index+4}: '{cell_value}' is geen geldig {file_extension}-bestand."
                    self.errors.append((sheetname, columnname, index+4, error_msg))
                
    def empty(self, sheetname, columnname, alarm_column_name="Alarmtext machine constructor (German)", must_be_empty=True):
        """ checks whether given column is empty or not, depending on parameter """
        if sheetname not in self.dataframes:
            print(f"Waarschuwing: Sheet '{sheetname}' niet gevonden.")
            return
        
        df = self.dataframes[sheetname]
        
        if columnname not in df.columns:
            print(f"Waarschuwing: Kolom '{columnname}' niet gevonden in '{sheetname}'.")
            return

        if alarm_column_name not in df.columns:
            print(f"Waarschuwing: Alarm-kolom '{alarm_column_name}' niet gevonden in '{sheetname}'.")
            return
        
        # controle per rij
        for index, row in df.iterrows():
            if self.alarm_exists(row, alarm_column_name):
                cell_value = row[columnname]
                
                if must_be_empty and pd.notna(cell_value):
                    error_msg = f"Rij {index+4}: '{cell_value}' moet leeg zijn, maar bevat een waarde!"
                    self.errors.append((sheetname, columnname, index+4, error_msg))
                elif not must_be_empty and pd.isna(cell_value):
                    error_msg = f"Rij {index+4}: Dit veld mag niet leeg zijn!"
                    self.errors.append((sheetname, columnname, index+4, error_msg))
    
    def log_errors(self, logfile="error_log.txt"): # variabele naam van maken, afh van Excel file
        if not self.errors:
            print("Geen fouten gevonden")
            return
        
        grouped_errors = defaultdict(list)
        
        # group errors per (sheet, column)
        for sheet, column, row, message in self.errors:
            grouped_errors[(sheet, column)].append(message)
        
        with open(logfile, "w") as log:
            for (sheet, column), messages in grouped_errors.items():
                log.write(f"### ERRORS IN SHEET '{sheet}', COLUMN '{column}' ###\n")
                for message in messages:
                    log.write(f"{message}\n")
                log.write("\n\n")
                
        print(f"Fouten opgeslagen in {logfile}")
        
file_path = "AlarmList_file_ingevuld.xlsx"
header_rows = { # ingeven via config
    "Alarmlist": 3,
    "Color Pictures": 3,
}
column_names = { # ingeven via config
    "Alarmlist": ['CRF / PCN', 'Version', 'PfizerNR', 'Alarmtext machine constructor (German)',
 'Alarmtext English', 'Dutch translation', 'Interlocks', 'Bypass', 'Stopmode',
 'Scada Alarmnr', 'Tagname', 'WORD number', 'bit in WORD', 'LAlm address',
 'PLC Data Type', 'PLC I/O', 'Class', 'PM67\nClass', 'VU-number', 'Picture',
 'Opkleuring\n(tags)', 'Color Picture', 'Lichtbalk\n(tekst)',
 'Lichtbalk (nummer)', 'Popup (tekst)', 'QSI', 'Alert\nmonitoring',
 'VQS reference', 'Hoorn / Buzzer', 'Special remarks', 'Pass / fail']
}

processor = ExcelProcessor(file_path, header_rows, column_names)
processor.load_excel()

df_alarmlist = processor.get_dataframe("Alarmlist").drop(0) # drop row index 0 ("VU X - VU Description")
# Index aanpassen zodat deze start bij 5
df_alarmlist.index = range(5, 5 + len(df_alarmlist))
# print(list(df_alarmlist.columns.values))
print(df_alarmlist.head())

# controle uitvoeren
validator = DataValidator(processor.dataframes)

validator.max_characters("Alarmlist", "Alarmtext English", 75)
validator.max_characters("Alarmlist", "Dutch translation", 75)

validator.file_type("Alarmlist", "Picture", "Alarmtext machine constructor (German)", ".pdl")

validator.empty("Alarmlist", "Pass / fail", must_be_empty=True)
validator.empty("Alarmlist", "Class", must_be_empty=False)

validator.log_errors()