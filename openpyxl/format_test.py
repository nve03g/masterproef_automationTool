from openpyxl import load_workbook

wb = load_workbook("format_test.xlsx")
ws = wb.active

last_row_index = ws.max_row
while last_row_index > 0 and ws[f"A{last_row_index}"].value is None:
    last_row_index-=1

for i in range(1,last_row_index+1):
    cell = ws[f"A{i}"]
    
    print(f"FORMAT DATA CELL 'A{i}':")
    # celkleur (achtergrond) van specifieke cel
    cell_color = cell.fill.fgColor.rgb if cell.fill.fgColor else None
    print(f"celkleur (rgb): {cell_color}")

    # tekstkleur van specifieke cel
    font_color = cell.font.color.rgb if cell.font.color else None # doesn't return a string if no particular color was set (default: black text)
    print(f"tekstkleur (rgb): {font_color}")

    print(f"Lettertype: {cell.font.name}")
    print(f"Vet: {cell.font.bold}")
    print(f"Cursief: {cell.font.italic}")
    print(f"Doorhaling: {cell.font.strike}")
    print("\n")