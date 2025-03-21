import openpyxl as xl
import datetime

### Excel bestand openen en lezen
wb = xl.load_workbook("testfile.xlsx") # load file

print(f"sheet names: {wb.sheetnames}")

ws_last_active = wb.active # sheet dat open stond wanneer file voor het laatst opgeslagen werd
print(f"last active sheet: {ws_last_active.title}")

ws_first_sheet = wb[wb.sheetnames[0]]
print(f"first sheet: {ws_first_sheet.title}\n")

ws_feb = wb["februari 25"]
print(f"rijen in {ws_feb}:")
for row in ws_feb.values:
    print(row)

print(f"\nwaarde in cel C3, januari 2025: {wb["januari 25"]["C3"].value}") # specifieke cel

print("\nrij 6, januari 2025:") # hele rij
for cel in wb["januari 25"][6]:
    print(cel.value)

print("\nkolom C, februari 2025:") # hele kolom
for cel in wb["februari 25"]["C"]:
    print(cel.value)


### data schrijven en opslaan
wb["maart 25"]["A2"] = datetime.datetime(2025, 3, 1) # specifieke cel

wb["maart 25"].append([datetime.datetime(2025, 3, 1), 210, "Basic Fit jaarabonnement"]) # hele nieuwe rij
wb["maart 25"].append([datetime.datetime(2025, 3, 1), 43, "XXL Nutrition proteïne & creatine"])

wb.save("updated_testfile.xlsx") # opslaan als nieuwe file


### alle rijen en kolommen doorlopen in een sheet
# alleen inhoud
print(f"\ndata in sheet {ws_last_active}:")
for rij in ws_last_active.iter_rows(values_only=True):
    print(rij)

# inhoud en coördinaten
print(f"\ndata en coördinaten in sheet {ws_last_active}:")
for rij in ws_last_active.iter_rows():
    for cel in rij:
        print(cel.coordinate, cel.value) # geeft bv. "A1 100"
