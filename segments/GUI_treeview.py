import tkinter as tk
from tkinter import ttk
import pandas as pd
import warnings

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
# print(df_alarmlist.head())

root = tk.Tk()
app = ExcelTreeview(root, df_alarmlist)
root.mainloop()