from docx import Document
import openpyxl
import re

# 📄 Word-Dokument laden
doc = Document("TEST20250617 Entwurf Arbeitshilfe Prüfung DORA clean.docx")

# ✅ Gewünschte Spaltennamen (sollen enthalten sein)
erforderliche_spalten = ["Themenkomplex", "Beispielhafte Prüfungshandlungen", "DORA"]

# 📊 Excel vorbereiten
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Alle Tabellen"

# 🧠 Tabellen mit den gesuchten Spalten extrahieren
def enthaelt_erforderliche_spalten(headerzeile, erforderlich):
    header_normiert = [h.strip().lower() for h in headerzeile]
    return all(any(erw.lower() in h for h in header_normiert) for erw in erforderlich)

zeilenzahl = 1

for table in doc.tables:
    if len(table.rows) < 2:
        continue  # überspringen, wenn zu klein

    # Header prüfen
    headerzellen = [cell.text.strip() for cell in table.rows[0].cells]
    if not enthaelt_erforderliche_spalten(headerzellen, erforderliche_spalten):
        continue

    # Tabelle übernehmen – mit Überschrift zur Trennung
    ws.append( headerzellen)
    zeilenzahl += 1

    for row in table.rows[1:]:
        zellen = [cell.text.strip() for cell in row.cells]
        # Leere oder zu kurze Zeilen überspringen
        if all(z == "" for z in zellen):
            continue
        # ggf. mit leeren Zellen auffüllen
        while len(zellen) < len(headerzellen):
            zellen.append("")
        ws.append(zellen)
        zeilenzahl += 1

    # Leerzeile zwischen Tabellen
    ws.append([])
    zeilenzahl += 1

# ✨ Zeilenumbruch aktivieren
for row in ws.iter_rows(min_row=1, max_row=zeilenzahl):
    for cell in row:
        cell.alignment = openpyxl.styles.Alignment(wrap_text=True)

# 📐 Spaltenbreite anpassen
for col in ws.columns:
    max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col)
    col_letter = col[0].column_letter
    ws.column_dimensions[col_letter].width = min(max_len + 5, 80)

# 💾 Speichern
wb.save("DORA_Alle_Tabellen_Vollständig.xlsx")
print("✅ Excel-Datei 'DORA_Alle_Tabellen_Vollständig.xlsx' wurde erfolgreich erstellt.")
