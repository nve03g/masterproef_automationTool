import tkinter as tk
from tkinter import ttk
import pandas as pd
import json
import warnings

## OPGEMERKT: laatste rij wordt niet juist berekend

warnings.simplefilter("ignore", UserWarning) # we krijgen warning dat openpyxl geen dropdownlijsten in excel meer ondersteunt, maar dat is geen probleem want die controle ga ik via mijn python code uitvoeren, dus deze warning mag genegeerd worden

class ExcelTreeview:
    def __init__(self, root, df):
        """ class that shows a dataframe in a treeview """
        self.root = root
        self.df = df
        self.root.title("Window title")
        self.root.geometry("1000x1000") # important to define window size!w
        
        # frame for treeview and scrollbars
        self.frame = ttk.Frame(self.root)
        self.frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        
        # scrollbars
        self.vsb = ttk.Scrollbar(self.frame, orient="vertical")
        self.hsb = ttk.Scrollbar(self.frame, orient="horizontal")
        
        # treeview widget
        self.tree = ttk.Treeview(
            self.frame,
            columns=list(self.df.columns),
            show="headings",
            yscrollcommand=self.vsb.set,
            xscrollcommand=self.hsb.set
        )
        
        # koppel scrollbars
        self.vsb.config(command=self.tree.yview)
        self.hsb.config(command=self.tree.xview)

        # grid positioning
        self.tree.grid(row=0, column=0, sticky='nsew')
        self.vsb.grid(row=0, column=1, sticky='ns')
        self.hsb.grid(row=1, column=0, sticky='ew')

        # frame layout configureren
        self.frame.grid_rowconfigure(0, weight=1)
        self.frame.grid_columnconfigure(0, weight=1)
        
        # kolomheaders instellen
        for col in self.df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor="w") # standaardbreedte instellen

        # voeg rijen toe aan treeview
        self.populate_treeview()
        
    def populate_treeview(self):
        """ Voeg data uit dataframe toe aan de treeview."""
        for _, row in self.df.iterrows():
            self.tree.insert("", "end", values=row.tolist())
 
 
class ExcelProcessor:
    def __init__(self, config_path):
        """
        initialize Excel file processor
        - config_path : str, path to JSON config file
        """
        self.load_config(config_path)
        self.dataframes = {} # dictionary containing all data {sheetname: DataFrame}
        
    def load_config(self, config_path):
        """ Load configuration parameters from JSON file. """
        with open(config_path, 'r', encoding="utf-8") as f:
            config = json.load(f)
        
        self.filepath = config["file_path"]
        self.headerrows = config["header_rows"]
        self.columnnames = config["column_names"]
        self.index_start = config["index_start"]
        
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
                    
                if sheet in self.index_start:
                    # verwijder overbodige rijen direct onder header
                    if (self.index_start[sheet] - self.headerrows[sheet] - 2) >= 0: # dan moeten we x aantal eerste rijen in df verwijderen
                        df = df.drop([i for i in range(self.index_start[sheet]-self.headerrows[sheet]-1)]) # -2+1 want anders range(0,0), dan krijg je lege lijst (rij 0 wordt niet gedropt)
                    
                    # pas index aan conform Excel lijst
                    df.index = range(self.index_start[sheet], self.index_start[sheet] + len(df))                                 
                    
                self.dataframes[sheet] = df
            else:
                print(f"Waarschuwing: {sheet} niet gevonden in {self.filepath}")
        
    def get_dataframe(self, sheetname):
        """ get specific sheet (DataFrame) """
        return self.dataframes.get(sheetname)
    


config_file = "config.json"
processor = ExcelProcessor(config_file)
processor.load_excel()


## dit gebruik ik om lijst te printen van alle kolomnamen, juiste syntax
df_alarmlist = processor.get_dataframe("Alarmlist")
# print(list(df_alarmlist.columns.values))

df_cp = processor.get_dataframe("Color Pictures")
# print(list(df_cp.columns.values))


root = tk.Tk()
app = ExcelTreeview(root, df_alarmlist)
root.mainloop()
