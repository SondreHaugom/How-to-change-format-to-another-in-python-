import openpyxl as op
# henter datane fra excel formatet
data = op.load_workbook(r"C:\Users\SondreHaugom\OneDrive - Telemark fylkeskommune\Sondre\Prakis oppgaver høst 2024\endreFromat\innbyggerliste-skien-test-20.11.2024.xlsx")
sheets =  data.sheetnames
sheet = data.active

#oppretter ny Excel-fil for export
ny_fil = op.Workbook()
ny_sheet = ny_fil.active
ny_sheet.tittle = "Nytt format"

# henter kolonner
ny_sheet.append(["Fornavn","Etternavn","Adresse","Post nummer","Post sted"])

# finner max antal rader 
rows = sheet.max_row

# henter innformasjonen fra det første excel arket
for i in range(2, rows + 1):
    fornavn = sheet[f"D{i}"].value
    etternavn = sheet[f"E{i}"].value
    adresse = sheet[f"F{i}"].value
    post_nummer = sheet[f"G{i}"].valu
    post_sted = sheet[f"H{i}"].value

    if fornavn and etternavn and adresse and post_nummer and post_sted:

        ny_sheet.append([fornavn,etternavn,adresse,post_nummer,post_sted])

# lager den nye excel filen
output_path = r"C:\Users\SondreHaugom\OneDrive - Telemark fylkeskommune\Sondre\Prakis oppgaver høst 2024\endreFromat\Nytt_format.xlsx"
ny_fil.save(output_path)

print(f"Data ekspoterer til {output_path}")