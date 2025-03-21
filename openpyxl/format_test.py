from openpyxl import load_workbook
import xlwings as xw # voor uitlezen van cel- en tekstkleuren (openpyxl alleen voor basiskleuren)
# NOG UITZOEKEN: export format to excel dan met xlwings of openpyxl?

wb = load_workbook("format_test.xlsx")
ws = wb.active

# last_row_index_wb = ws.max_row
# while last_row_index_wb > 0 and ws[f"A{last_row_index_wb}"].value is None:
#     last_row_index_wb-=1
    
    
# is dit nodig?? JA, want sheet en wb hebben ander type en anders kan formattering niet correct uitgelezen worden
sheet = xw.Book("format_test.xlsx").sheets[0]

last_row_index_sheet = sheet.range("A" + str(sheet.cells.last_cell.row)).end("up").row

# print(f"wb: {last_row_index_wb}, sheet: {last_row_index_sheet}") # komen overeen

# print(f"wb: {type(wb)}, sheet: {type(sheet)}") # niet hetzelfde, obviously

for i in range(1,last_row_index_sheet+1):
    cell = ws[f"A{i}"]
    
    print(f"FORMAT DATA CELL 'A{i}':")
    # PAS OP: niet alle kleuren worden correct uitgelezen! Hiervoor ga ik xlwings gebruiken
    
    # met xlwings
    # celkleur (achtergrond) van specifieke cel
    cell_color = sheet.range(f'A{i}').color # geeft een tuple (R, G, B) of None als geen kleur ingesteld
    print(f"celkleur (rgb): {cell_color}")

    # tekstkleur van specifieke cel
    font_color = sheet.range(f'A{i}').font.color # geeft een tuple (R, G, B), (0, 0, 0) als geen kleur ingesteld (default zwarte tekst)
    print(f"tekstkleur (rgb): {font_color}")

    # met openpyxl
    print(f"Lettertype: {cell.font.name}")
    print(f"Lettergrootte: {cell.font.size}")
    print(f"Vet: {cell.font.bold}")
    print(f"Cursief: {cell.font.italic}")
    print(f"Doorhaling: {cell.font.strike}")
    print("\n")