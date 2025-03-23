import tkinter as tk
from tkinter import ttk
import pandas as pd
import json
import warnings

## OPGEMERKT: laatste rij wordt niet juist berekend

warnings.simplefilter("ignore", UserWarning) # we krijgen warning dat openpyxl geen dropdownlijsten in excel meer ondersteunt, maar dat is geen probleem want die controle ga ik via mijn python code uitvoeren, dus deze warning mag genegeerd worden

class ExcelTreeview:
    def __init__(self, root, processor):
        """ GUI class showing dropdown for profile selection and treeview to display dataframe. """
        self.root = root
        self.processor = processor
        self.root.title("Pfizer Automation Tool")
        self.root.geometry("1000x800") # important to define window size!
        
        # frame for dropdown and treeview
        self.frame = ttk.Frame(self.root)
        self.frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        
        # dropdown menu for profile selection
        self.profile_var = tk.StringVar(value=self.processor.profile)
        self.profile_dropdown = ttk.Combobox(self.frame, textvariable=self.profile_var, values=list(self.processor.profiles.keys()), state="readonly") # processor.profiles.keys() geeft alle mogelijke profielen aangegeven in config file
        self.profile_dropdown.pack(pady=5)
        self.profile_dropdown.bind("<<ComboboxSelected>>", self.update_profile)
        
        # scrollbars
        self.vsb = ttk.Scrollbar(self.frame, orient="vertical")
        self.hsb = ttk.Scrollbar(self.frame, orient="horizontal")
        
        # treeview widget
        self.tree = ttk.Treeview(
            self.frame,
            show="headings",
            yscrollcommand=self.vsb.set,
            xscrollcommand=self.hsb.set
        )
        # koppel scrollbars
        self.vsb.config(command=self.tree.yview)
        self.hsb.config(command=self.tree.xview)

        # set grid layout
        self.tree.pack(fill=tk.BOTH, expand=True)
        self.vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.hsb.pack(side=tk.BOTTOM, fill=tk.X)

        # show initial data
        self.load_treeview()
        
    def update_profile(self, event=None):
        """ Update user profile and load correct data into treeview. """
        new_profile = self.profile_var.get()
        self.processor.set_profile(new_profile)
        self.processor.load_excel() # load excel file with new profile
        self.load_treeview()
        
    def load_treeview(self):
        """ Reload treeview with correct columns and data. """
        self.tree.delete(*self.tree.get_children()) # delete current data
        
        df = self.processor.get_dataframe("Color Pictures") # toon standaard deze sheet
        if df is not None:
            self.tree["columns"] = list(df.columns)
            for col in df.columns:
                self.tree.heading(col, text=col)
                self.tree.column(col, width=100, anchor="w")
                
            # populate treeview
            for _, row in df.iterrows():
                self.tree.insert("", "end", values=row.tolist())
        
        
class ExcelProcessor:
    def __init__(self, config_path, profile):
        """
        initialize Excel file processor
        - config_path : str, path to JSON config file
        - profile : str, huidige hardcoded profielaanwijzing (supplier, developer, operator)
        """
        self.load_config(config_path)
        self.dataframes = {} # dictionary containing all data {sheetname: DataFrame}
        self.profile = profile
        self.load_excel()
        
    def load_config(self, config_path):
        """ Load configuration parameters from JSON file. """
        with open(config_path, 'r', encoding="utf-8") as f:
            config = json.load(f)
        
        self.filepath = config["file_path"]
        self.default_profile = config["default_profile"]
        self.headerrows = config["header_rows"]
        self.index_start = config["index_start"]
        self.columnnames = config["column_names"]
        self.profiles = config["profiles"]
        
    def set_profile(self, profile):
        """ Set active profile (OR default profile). """
        if profile not in self.profiles:
            self.profile = self.default_profile
        else:
            self.profile = profile
                
    def load_excel(self):
        """ Laad de Excel-sheets in dataframes, met de opgegeven header-rij per sheet en toegestane kolommen afhankelijk van het aangeduide profiel. """
        xls = pd.ExcelFile(self.filepath) # open het Excel-bestand
        
        for sheet, header_row in self.headerrows.items():
            if sheet in xls.sheet_names:
                df = pd.read_excel(self.filepath, sheet_name=sheet, header=header_row-1)
                
                # gebruik enkel toegewezen kolommen (afhankelijk van profiel)
                allowed_columns = self.profiles[self.profile][sheet]
                valid_columns = [col for col in allowed_columns if col in df.columns]
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
current_profile = "operator" # wordt later ingesteld via GUI dropdown list
processor = ExcelProcessor(config_file, profile=current_profile)


## dit gebruik ik om lijst te printen van alle kolomnamen, juiste syntax
# df_alarmlist = processor.get_dataframe("Alarmlist")
# print(list(df_alarmlist.columns.values))

# df_cp = processor.get_dataframe("Color Pictures")
# print(list(df_cp.columns.values))


root = tk.Tk()
app = ExcelTreeview(root, processor)
root.mainloop()
