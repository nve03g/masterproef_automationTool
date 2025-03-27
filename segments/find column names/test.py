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

file = "AlarmList_CU7000_B5_FREEZE.xlsx"
header_rows = {
    "Version control": 3,
    "Alarmlist": 3,
    "Color Pictures": 3,
}
# column_names = {
#     "Alarmlist": ['CRF / PCN', 'Version', 'PfizerNR', 'Alarmtext machine constructor (German)',
#  'Alarmtext English', 'Dutch translation', 'Interlocks', 'Bypass', 'Stopmode',
#  'Scada Alarmnr', 'Tagname', 'WORD number', 'bit in WORD', 'LAlm address',
#  'PLC Data Type', 'PLC I/O', 'Class', 'PM67\nClass', 'VU-number', 'Picture',
#  'Opkleuring\n(tags)', 'Color Picture', 'Lichtbalk\n(tekst)',
#  'Lichtbalk (nummer)', 'Popup (tekst)', 'QSI', 'Alert\nmonitoring',
#  'VQS reference', 'Hoorn / Buzzer', 'Special remarks', 'Pass / fail'],
# }
column_names = {}

processor = ExcelProcessor(file, header_rows, column_names)
processor.load_excel()

sheetname = "Version control" # sheetname to test
df = processor.get_dataframe(sheetname).drop(0)
# Index aanpassen zodat deze start bij 5
df.index = range(5, 5 + len(df))
print(f"column names for sheet '{sheetname}':\n")
print(list(df.columns.values))
# print(df.head())