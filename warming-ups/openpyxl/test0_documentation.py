import openpyxl as xl

# wb = xl.Workbook()

# ws1 = wb.active # sheet in file, default 0 (so unless you change it, always the first sheet)

# ws2 = wb.create_sheet("worksheet 2") # insert at end (default)
# ws0 = wb.create_sheet("worksheet 100",0) # insert in beginning
# ws_voorlaatste = wb.create_sheet("voorlaatste worksheet",-1) # insert op voorlaatste positie

# ws1.title = "New Title" # verander naam van bestaande sheet
# # once a sheet is named, you can access it as a key of the workbook (Excel file)
# ws2 = wb["New Title"]

# print(wb.sheetnames)

# # loop through sheets
# for sheet in wb:
#     print(sheet.title)

# wb.save("testfile.xlsx")




wb = xl.load_workbook("testfile.xlsx")

for cell in wb["Sheet"]["1"]:
    print(cell.value)

wb["Sheet"]["D1"] = "Hallo"
wb.save("testfile.xlsx")