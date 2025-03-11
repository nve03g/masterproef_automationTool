import pandas as pd
import warnings

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
            if isinstance(value, str) and len (value) > max_chars:
                error_msg = f"Rij {index+1}: '{value}' is te lang ({len(value)} > {max_chars})"
                self.errors.append((sheetname, columnname, index+1, error_msg))
                
    def log_errors(self, logfile="error_log.txt"): # variabele naam van maken, afh van Excel file
        if not self.errors:
            print("Geen fouten gevonden")
            return
        
        with open(logfile, "w") as log:
            for sheet, column, row, message in self.errors:
                log.write(f"[{sheet}] {message}\n")
                
        print(f"Fouten opgeslagen in {logfile}")
        
file_path = "AlarmList_file_ingevuld.xlsx"
header_rows = { # ingeven via config
    "Alarmlist": 3,
    "Color Pictures": 3,
}
column_names = { # ingeven via config
    "Alarmlist": ['CRF / PCN', 'Version', 'PfizerNR', 'Alarmtext machine constructor', 'Alarmtext English', 'Dutch translation', 'Interlocks', 'Bypass', 'Stopmode', 'Scada Alarmnr', 'Tagname', 'WORD number', 'bit in WORD', 'LAlm address', 'PLC Data Type', 'PLC I/O', 'Class', 'PM67\nClass', 'VU-number', 'Picture', 'Opkleuring Object tag\n(tags)', 'Colour Picture', 'Lichtbalk\n(tekst)', 'Lichtbalk (nummer)', 'Popup (tekst)', 'QSI', 'Alert\nmonitoring', 'VQS reference', 'Special remarks', 'Horn/Buzzer', 'Pass / fail']
}

processor = ExcelProcessor(file_path, header_rows, column_names)
processor.load_excel()

df_alarmlist = processor.get_dataframe("Alarmlist").drop(0) # drop row index 0 ("VU X - VU Description")
# print(list(df_alarmlist.columns.values))
# print(df_alarmlist.head())


# controle uitvoeren
validator = DataValidator(processor.dataframes)

validator.max_characters("Alarmlist", "Alarmtext English", 75)
validator.log_errors()
validator.max_characters("Alarmlist", "Dutch translation", 75)
validator.log_errors()