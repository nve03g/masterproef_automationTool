import tkinter as tk
from tkinter import ttk
import pandas as pd


def load_excel(file_path):
    df = pd.read_excel(file_path)
    return df


def create_gui(df):
    root = tk.Tk()
    root.title("Excel Data in Treeview")

    # Maak een frame voor de Treeview en scrollbars
    frame = ttk.Frame(root)
    frame.pack(padx=10, pady=10)

    # Voeg een verticale scrollbar toe
    vsb = ttk.Scrollbar(frame, orient="vertical")
    vsb.grid(row=0, column=1, sticky='ns')

    # Voeg een horizontale scrollbar toe
    hsb = ttk.Scrollbar(frame, orient="horizontal")
    hsb.grid(row=1, column=0, sticky='ew')

    # Maak de Treeview widget
    tree = ttk.Treeview(frame, columns=list(df.columns), show="headings", yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    tree.grid(row=0, column=0, sticky='nsew')

    # Verbind de scrollbars aan de Treeview
    vsb.config(command=tree.yview)
    hsb.config(command=tree.xview)

    # Stel de kolomheaders in
    for col in df.columns:
        tree.heading(col, text=col)

    # Voeg de rijen toe aan de Treeview
    for _, row in df.iterrows():
        tree.insert("", "end", values=row.tolist())

    # Maak de layout flexibel
    frame.grid_rowconfigure(0, weight=1)
    frame.grid_columnconfigure(0, weight=1)

    root.mainloop()


file_path = "AlarmList_file_ingevuld.xlsx"
df = load_excel(file_path)

create_gui(df)