import pandas as pd
import warnings

warnings.simplefilter("ignore", UserWarning)

# sheet inlezen in een df
class ExcelProcessor:
    def __init__(self, filepath, headerrows, columnnames):
        self.filepath = filepath
        self.headerrows = headerrows
        self.columnnames = columnnames
        self.dataframes = {} # dictionary containing all data {sheetname: DataFrame}
        
    def load_excel(self):
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

file = "Visualisation&Commands_CU7000_CU7100_B4_FREEZE.xlsm"
header_rows = {
    "Version control": 3,
    "Bit info": 3,
    "Bit commands": 3,
    "Colour pictures Status": 3,
    "Interlock": 3,
    "Buttons": 3,
    "Motor (6 bytes)": 3,
    "Valve": 3,
    "Motor (48 bytes)": 3,
    "Input value": 3,
    "Measurement": 3,
    "Controller": 3,
    "Output values": 3,
    "Template history": 3
}
column_names = {}

processor = ExcelProcessor(file, header_rows, column_names)
processor.load_excel()

# sheet to test
sheetname = "Template history"
df = processor.get_dataframe(sheetname)

print(f"column names for sheet '{sheetname}':\n")
print(list(df.columns.values))