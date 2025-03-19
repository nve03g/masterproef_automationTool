import xlwings as xw
import pandas as pd

sheet = xw.Book("format_test.xlsx").sheets[0]

# aantal OPEENVOLGENDE niet-lege cellen in een kolom
column = sheet.range("A1").expand("down")  # Bepaalt het bereik van niet-lege cellen
amount = column.rows.count

# lengte van kolom: laatst gevulde cel in de kolom
last_row_index = sheet.range("A" + str(sheet.cells.last_cell.row)).end("up").row

# volledige kolom
# print(sheet[f"A1:A{last_row_index}"].value)

# alle data (tot en met laatst gebruikte cel)
all_data = sheet.used_range.value
df_all_data = pd.DataFrame(all_data)

start_cell = sheet.used_range[0,0].address  # Bovenste linkse cel
end_cell = sheet.used_range.rows[-1].columns[-1].address  # Rechtsonderste cel
initial_range = f"{start_cell.replace("$","")}:{end_cell.replace("$","")}"
# print(f"initieel bereik: {initial_range}")
# print(f"initial size: ({sheet.used_range.rows.count}, {sheet.used_range.columns.count})")

# print in tabel, stop bij eerste lege cel
# print(sheet["A1"].expand().value)
# print(sheet.range("A1").options(expand='table').value) # options: table, right, down
df1 = pd.DataFrame(sheet["A1"].expand().value)

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
# print(f"nieuw bereik: {new_range}")
# print(f"new size: ({range_end_row - range_start_row +1}, {range_end_column - range_start_column +1})")

# kies een range die je dan in df plaatst (+ meerdere header rijen mogelijk)
df2 = sheet[initial_range].options(pd.DataFrame, header=2, index=False).value
df3 = sheet[new_range].options(pd.DataFrame, header=False, index=False).value
# print(df3)

# schrijf df terug naar bestaande excel file
df_write_data = pd.DataFrame([[1.1, 2.2], [3.3, None]], columns=['one', 'two'])
# sheet["E1"].options(index=False, header=True).value = df_write_data


## FORMATTERING
"""
Op macOS werkt xlwings anders dan op Windows, omdat het via AppleScript met Excel communiceert in plaats van via COM (zoals op Windows). Hierdoor zijn sommige functies zoals .get_format() niet beschikbaar op Mac.

een oplossing kan zijn: 
schrijf een functie die, afhankelijk van OS op een andere manier de formattering ophaalt, en deze in een gelijke dictionary returnt, zodat je over heel de code uniform naar de formattering kan verwijzen
"""
## ik test hier voor MacOS

for i in range(1,last_row_index+1):
    print(f"FORMAT DATA CELL 'A{i}':")
    # celkleur (achtergrond) van specifieke cel
    cell_color = sheet.range(f'A{i}').color  # Geeft een tuple (R, G, B) of None als geen kleur ingesteld
    print(f"celkleur (rgb): {cell_color}")

    # tekstkleur van specifieke cel
    font_color = sheet.range(f'A{i}').font.color
    print(f"tekstkleur (rgb): {font_color}")

    print(f"Lettertype: {sheet.range(f'A{i}').font.name}")
    print(f"Vet: {sheet.range(f'A{i}').font.bold}")
    print(f"Cursief: {sheet.range(f'A{i}').font.italic}")
    # print(f"Doorhaling: {sheet.range(f'A{i}').api.Font.Strikethrough}") # werkt op Windows, niet op Mac
    print("\n")