import tkinter as tk
from tkinter import ttk, filedialog
import json
import warnings
import pandas as pd

# REMARK: last data row isn't correctly calculated

# We get a warning that openpyxl no longer supports dropdown lists in Excel, 
# but that's not a problem because we’re performing that check 
# through the python code, so this warning may be ignored.
warnings.simplefilter("ignore", UserWarning) 


class ExcelTreeview:
    """ 
    GUI class showing dropdown for profile selection, 
    sheet selection and treeview to display dataframe. 
    """
    def __init__(self, root, processor):
        self.root = root
        self.processor = processor
        self.root.title("Pfizer Automation Tool")
        self.root.geometry("1000x800")  # It's important to define window size!

        # Frame for dropdown and treeview.
        self.frame = ttk.Frame(self.root)
        self.frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        # Dropdown menu for profile selection.
        self.profile_var = tk.StringVar()  # no start-value
        self.profile_dropdown = ttk.Combobox(
            self.frame, 
            textvariable=self.profile_var, 
            values=list(self.processor.profiles.keys()),
            # processor.profiles.keys() gives all possible profiles specified in the config file
            state="readonly"
            )
        self.profile_dropdown.pack(pady=5)
        self.profile_dropdown.bind("<<ComboboxSelected>>", self.update_profile)

        # Dropdown menu for sheet selection.
        self.sheet_var = tk.StringVar()  # no start-value
        self.sheet_dropdown = ttk.Combobox(
            self.frame, 
            textvariable=self.sheet_var, 
            state="readonly"
            )
        self.sheet_dropdown.pack(pady=5)
        self.sheet_dropdown.bind("<<ComboboxSelected>>", self.update_sheet)
        
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
        # Link the scrollbars to the treeview.
        self.vsb.config(command=self.tree.yview)
        self.hsb.config(command=self.tree.xview)

        # Set grid layout.
        self.tree.pack(fill=tk.BOTH, expand=True)
        self.vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.hsb.pack(side=tk.BOTTOM, fill=tk.X)

        # Show initial data in the treeview.
        self.load_treeview()
        
    def update_profile(self, event=None):
        """ Update user profile and load correct data into treeview. """
        new_profile = self.profile_var.get()
        # Store currently selected sheet.
        current_sheet = self.sheet_var.get()
        self.processor.set_profile(new_profile)
        
        # Update sheet dropdown options based on selected profile.
        self.update_sheet_options(current_sheet)
        
        # Load excel file with new profile.
        self.processor.load_excel()
        # Load the treeview for the current or restored sheet.
        self.load_treeview(self.sheet_var.get())
        
    def update_sheet_options(self, previous_sheet=None):
        """ Update the available sheet options in dropdown list according to current profile and try to keep the previous selection. """
        # Get the sheets for current user profile out of config file.
        available_sheets = self.processor.get_config_sheets()
        self.sheet_dropdown['values'] = available_sheets
        
        # Keep previously selected sheet open if it's still available, otherwise fall back to first available sheet.
        if previous_sheet in available_sheets:
            self.sheet_var.set(previous_sheet)
        elif available_sheets:
            self.sheet_var.set(available_sheets[0])
        # else:
        #     self.sheet_var.set("")
            
        # Automatically load the selected sheet.
        self.update_sheet()
        
    def update_sheet(self, event=None):
        """ Update the treeview with data from selected sheet. """
        sheet_name = self.sheet_var.get()
        self.load_treeview(sheet_name)
        
    def load_treeview(self, sheet_name=None):
        """ Reload treeview with correct sheet, columns and data. """
        # Delete the current treeview data.
        self.tree.delete(*self.tree.get_children())
        
        if sheet_name is None:
            # By default take the current selected sheet.
            sheet_name = self.sheet_var.get()
        
        # Get dataframe for selected sheet.
        df = self.processor.get_dataframe(sheet_name)
        if df is not None:
            self.tree["columns"] = list(df.columns)
            for col in df.columns:
                self.tree.heading(col, text=col)
                self.tree.column(col, width=100, anchor="w")
                
            # populate treeview
            for _, row in df.iterrows():
                self.tree.insert("", "end", values=row.tolist())
        
        
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
    


config_file = "config.json"
current_profile = "operator"  # Will later be set through GUI dropdown list.
processor = ExcelProcessor(config_file, profile=current_profile)


# I use these lines to print a list of all column names for a specific sheet in the right syntax, to be able to put it correctly in the config file.
# df_alarmlist = processor.get_dataframe("Alarmlist")
# print(list(df_alarmlist.columns.values))

# df_cp = processor.get_dataframe("Color Pictures")
# print(list(df_cp.columns.values))


root = tk.Tk()
app = ExcelTreeview(root, processor)
root.mainloop()
