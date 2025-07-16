import openpyxl
from openpyxl import Workbook
from datetime import datetime, timedelta
import locale


# Loclae für Deutsch Wochentage und Datum 

try: 
    locale.setlocale(locale.LC_TIME, 'de_DE.UTF-8')
except:
    try: 
        locale.setlocale(locale.LC_TIME, 'deu')
    except: 
        print("Warnung: Konnte deutsches Locale nicht setzen.")


# Neues Workbook 

wb = Workbook()

# BLATT 1 Werktage 1 

ws1 = wb.active
ws1.title = "Blatt 1"
ws1.append(["lfd. Nummer", "Wochentag", "Datum"])

start_datum = datetime.strptime("01.11.2024", "%d.%m.%Y")
end_datum = datetime.strptime("30.07.2025", "%d.%m.%Y")

aktuelles_datum = start_datum
laufende_nummer = 1 

while aktuelles_datum <= end_datum:
    if aktuelles_datum.weekday() < 5: # Montag - Freitag 
        wochentag = aktuelles_datum.strftime("%A")
        datum_str = aktuelles_datum.strftime("%d.%m.%Y")
        ws1.append([laufende_nummer, wochentag, datum_str])
        laufende_nummer += 1
    aktuelles_datum += timedelta(days=1)

# Blatt 2 KW 

ws2 = wb.create_sheet("Kalenderwochen")
ws2.append(["Kalenderwochen", "Stratdatum (Mo)", "Enddatum (SO)"])

# Zurück zum ersten Montag vor dem Startdatum 
kw_start = start_datum
while kw_start.weekday() != 0:
    kw_start -= timedelta(days=1)


while kw_start <= end_datum:
    kw_ende = kw_start + timedelta(days=6)
    if kw_ende < start_datum:
        kw_start += timedelta(weeks=1)
        continue
    if kw_start > end_datum:
        break

    kw_nummer = kw_start.isocalendar()[1]
    jahr = kw_start.isocalendar()[0]
    start_str = kw_start.strftime("%d.%m.%Y")
    ende_str = kw_ende.strftime("%d.%m.%Y")

    ws2.append([f"KW {kw_nummer}/{jahr}", start_str, ende_str])
    kw_start += timedelta(weeks=1)

# end of while 


wb.save("Werktage_und_Kalenderwochen.xlsx")
