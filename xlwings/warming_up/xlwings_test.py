import xlwings as xw

wb = xw.Book("voorbeeld.xlsx") # open existing file
sheet = wb.sheets["Sheet1"] # select a sheet

# Lees een cel
waarde = sheet.range("A1").value
# print(waarde)

sheet.range("B2").value = "Nellie :)" # write to Excel (real-time)
# je kan live de aanpassingen zien in de Excel lijst, zonder file te moeten sluiten en opnieuw openen

# Leest een kolom
waarden = sheet.range("A1:A4").value
# print(waarden)

sheet.range("C1:C4").value = ["easy", "peasy", "lemon", "squeezy"] # simple list -> row
# XLWings behandelt lijsten als RIJEN ipv kolommen, use nested list voor kolom!
# sheet.range("C1:C4").value = [["easy"], ["peasy"], ["lemon"], ["squeezy"]] # nested list -> kolom
sheet.range("C1:C4").options(transpose=True).value = ["easy", "peasy", "lemon", "squeezy"] # transposed list = nested list -> kolom

# save and close Excel file
wb.save("voorbeeld_edited.xlsx") # als je save() gebruikt (zonder filename) dan wordt originele file aangepast, anders wordt nieuwe file aangemaakt met alle wijzigingen die je net hebt uitgevoerd
wb.close() # free memory, avoid memory leaks

"""
ofwel doe je "file open, ..., file close", 
ofwel doe je "with file open as ..., do ..., save" (sluit automatisch)
"""