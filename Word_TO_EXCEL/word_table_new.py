from docx import Document
import openpyxl
import re

# Word-Dokument laden
doc = Document("TEST20250617 Entwurf Arbeitshilfe Prüfung DORA clean.docx")

# Erwartete Header (diese sollen vollständig übernommen werden!)
erwartete_header = ["Themenkomplex", "Beispielhafte Prüfungshandlungen", "DORA"]

# Tabellen mit diesen Headern suchen
def finde_alle_passenden_tabellen(doc, erwartete_header):
    passende_tabellen = []
    for table in doc.tables:
        if len(table.rows) == 0:
            continue
        erste_zeile = [cell.text.strip() for cell in table.rows[0].cells]
        if all(any(h.lower() in z.lower() for z in erste_zeile) for h in erwartete_header):
            passende_tabellen.append(table)
    return passende_tabellen

# Excel vorbereiten
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "DORA-Rohdaten"

# Tabellen extrahieren
tabellen = finde_alle_passenden_tabellen(doc, erwartete_header)
if not tabellen:
    print("❌ Keine passenden Tabellen gefunden.")
    exit()

# Spaltenüberschriften aus der ersten passenden Tabelle übernehmen
erste_tabelle = tabellen[0]
header_zeile = [cell.text.strip() for cell in erste_tabelle.rows[0].cells]
ws.append(header_zeile)

# Daten aus allen Tabellen übernehmen
for tabelle in tabellen:
    for row in tabelle.rows[1:]:
        zellen = [cell.text.strip() for cell in row.cells]
        # Nur komplette Zeilen übernehmen
        if len(zellen) == len(header_zeile):
            ws.append(zellen)

# Formatierung: Zeilenumbruch aktivieren
for row in ws.iter_rows(min_row=2):
    for cell in row:
        cell.alignment = openpyxl.styles.Alignment(wrap_text=True)

# Spaltenbreiten automatisch anpassen
for col in ws.columns:
    max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col)
    col_letter = col[0].column_letter
    ws.column_dimensions[col_letter].width = min(max_len + 5, 80)

# Excel speichern
wb.save("NEW_DORA_Rohdaten.xlsx")
print("✅ Fertig! Datei gespeichert als 'DORA_Rohdaten.xlsx'")
