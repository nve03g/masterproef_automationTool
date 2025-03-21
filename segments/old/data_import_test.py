import xlwings as xw
import pandas as pd
import unittest

sheet = xw.Book("data_import_test.xlsx").sheets[0]

def count_consecutive_cells(sheet, start_cel="A1"):
    # aantal OPEENVOLGENDE niet-lege cellen in een kolom vanaf bepaalde cel
    column = sheet.range(f"{start_cel}").expand("down")  # Bepaalt het bereik van niet-lege cellen
    return column.rows.count

def get_last_filled_row(sheet, col_letter="A"):
    # lengte van kolom: laatst gevulde cel in de kolom
    return sheet.range(f"{col_letter}{sheet.cells.last_cell.row}").end("up").row

def get_full_used_column(sheet, column_letter="A"):
    # volledige kolom
    return sheet[f"{column_letter}1:{column_letter}{get_last_filled_row(sheet, column_letter)}"].value

def get_used_range(sheet):
    # alle data (tot en met laatst gebruikte cel)
    all_data = sheet.used_range.value
    return pd.DataFrame(all_data)

def get_initial_range_size(sheet):
    start_cell = sheet.used_range[0,0].address  # Bovenste linkse cel
    end_cell = sheet.used_range.rows[-1].columns[-1].address  # Rechtsonderste cel
    initial_range = f"{start_cell.replace("$","")}:{end_cell.replace("$","")}"
    initial_size = f"({sheet.used_range.rows.count}, {sheet.used_range.columns.count})"
    return initial_range, initial_size

# print in tabel, stop bij eerste lege cel
# print(sheet["A1"].expand().value)
# print(sheet.range("A1").options(expand='table').value) # options: table, right, down
df1 = pd.DataFrame(sheet["A1"].expand().value)

def calc_new_range_size(sheet):
    # nieuwe range -> get begin- and start row- and column
    range_start_row = sheet.used_range.row + 2
    range_start_column = sheet.used_range.column + 2
    range_end_row = range_start_row + sheet.used_range.rows.count - 1
    range_end_column = range_start_column + sheet.used_range.columns.count - 1
    # convert to Excel letters
    range_start_column_letter = xw.utils.col_name(range_start_column)
    range_end_column_letter = xw.utils.col_name(range_end_column)
    # new range
    new_range = f"{range_start_column_letter}{range_start_row}:{range_end_column_letter}{range_end_row}"
    new_size = f"({range_end_row - range_start_row +1}, {range_end_column - range_start_column +1})"
    return new_range, new_size

# kies een range die je dan in df plaatst (+ meerdere header rijen mogelijk)
df2 = sheet[get_initial_range_size(sheet)[0]].options(pd.DataFrame, header=2, index=False).value
df3 = sheet[calc_new_range_size(sheet)[0]].options(pd.DataFrame, header=False, index=False).value
# print(df3)

df_write_data = pd.DataFrame([[1.1, 2.2], [3.3, None]], columns=['one', 'two'])
def write_df_to_sheet(df, sheet, start_cell="E1", header=True):
    # schrijf df terug naar bestaande excel file
    sheet[f"{start_cell}"].options(index=False, header=header).value = df
    return


n=0
assert 1<n, 'The assertion is false!'
# assert <condition being tested>, <error message to be displayed>

