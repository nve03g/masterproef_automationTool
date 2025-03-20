import pandas as pd
import warnings
from collections import defaultdict
from openpyxl import load_workbook

warnings.simplefilter("ignore", UserWarning) # we krijgen warning dat openpyxl geen dropdownlijsten in excel meer ondersteunt, maar dat is geen probleem want die controle ga ik via mijn python code uitvoeren, dus deze warning mag genegeerd worden

class ExcelDataProcessor:
    def __init__(self, filepath):

        self.filepath = filepath
        self.data_dfs = {} # dictionary containing all data {sheetname: DataFrame}
        self.format_dfs = {} # dictionary containing all format data {sheetname: DataFrame}
        
    def load_all_sheets(self):
        """ Laad de Excel-sheets in dataframes. """
        try:
            dfs = pd.read_excel(self.filepath, sheet_name=None, header=None) # all sheets, without headers
            self.data_dfs = {sheet: df for sheet, df in dfs.items()}
        except Exception as e:
            print(f"Fout bij laden van de sheets: {e}")
                
    def load_formatting(self):
        """ Laad formattering per cel in dicts in dfs. """
        try:
            wb = load_workbook(self.filepath)
            for sheet in wb.sheetnames:
                ws = wb[sheet]
                
                format_data = []
                
                for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                    format_row = [
                        {
                            "cell color": cell.fill.fgColor.rgb if cell.fill.fgColor else None,
                            "text color": cell.font.color.rgb if cell.font.color else None, # doesn't return a string if no particular color was set (default: black text)
                            "font": cell.font.name,
                            "bold": cell.font.bold,
                            "italic": cell.font.italic,
                            "strikethrough": cell.font.strike
                        }
                        for cell in row
                    ]
                    format_data.append(format_row)

                self.format_dfs[sheet] = pd.DataFrame(format_data)

        except Exception as e:
            print(f"Fout bij laden van de formattering: {e}")
        
    def get_dataframe(self, sheetname):
        """ get specific sheet (DataFrame) """
        return self.data_dfs.get(sheetname)
    
    

file_path = "format_test.xlsx"
processor = ExcelDataProcessor(file_path)
processor.load_all_sheets()
processor.load_formatting()

data_dictionary = processor.data_dfs
format_dictionary = processor.format_dfs

print(data_dictionary)
print()
print(format_dictionary)

# print format data per cell in certain sheet
# sheetname = "Sheet2"
# first_row_format_df = format_dictionary.get("Sheet1").loc[0] # type: dict
# for i in range(format_dictionary.get(sheetname).shape[0]):  # Iterate through rows
#     for j in range(format_dictionary.get(sheetname).shape[1]): # iterate through columns
#         for key, values in format_dictionary.get(sheetname).iloc[i,j].items():
#             print(f"{key}: {values}")
#         print("\n")

