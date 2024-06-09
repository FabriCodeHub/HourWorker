import datetime
import pandas as pd
from openpyxl import Workbook
import pyfiglet
import os

def normalizza_orario(orario):
    return orario.replace('.', ':')

def inserisci_settimana():
    while True:
        try:
            data_inizio = input("Inserisci la data di inizio della settimana (dd/mm/yyyy): ")
            data_fine = input("Inserisci la data di fine della settimana (dd/mm/yyyy): ")
            data_inizio = datetime.datetime.strptime(data_inizio, "%d/%m/%Y")
            data_fine = datetime.datetime.strptime(data_fine, "%d/%m/%Y")
            if data_inizio > data_fine:
                print("La data di inizio deve essere precedente alla data di fine. Riprova.")
                continue
            return data_inizio, data_fine
        except ValueError:
            print("Formato data non valido. Riprova.")

def inserisci_orario_lavorativo(giorno_settimana):
    while True:
        orario_inizio = input(f"Inserisci Orario di Inizio Lavoro per {giorno_settimana} (HH:MM o HH.MM) o '/' se non lavori: ")
        if orario_inizio.strip() == '/':
            return None, None
        try:
            orario_fine = input(f"Inserisci Orario di Fine Lavoro per {giorno_settimana} (HH:MM o HH.MM): ")
            
            orario_inizio = datetime.datetime.strptime(normalizza_orario(orario_inizio), "%H:%M")
            orario_fine = datetime.datetime.strptime(normalizza_orario(orario_fine), "%H:%M")
            
            if orario_fine < orario_inizio:
                print("L'orario di fine deve essere successivo all'orario di inizio. Riprova.")
                continue
            
            return orario_inizio, orario_fine
        except ValueError:
            print("Formato orario non valido. Riprova.")

def main():
    # Titolo
    titolo = pyfiglet.figlet_format("HourWorker", font="slant")
    print(titolo)
    print("_" * 60)

    # Cartella di destinazione per il file Excel
    cartella_destinazione = r"C:\Users\batte\Desktop\HourWorker"
    if not os.path.exists(cartella_destinazione):
        os.makedirs(cartella_destinazione)

    data_inizio, data_fine = inserisci_settimana()
    giorni_settimana = ["Lunedì", "Martedì", "Mercoledì", "Giovedì", "Venerdì", "Sabato", "Domenica"]
    dati_settimanali = []

    for giorno_settimana in giorni_settimana:
        orario_inizio, orario_fine = inserisci_orario_lavorativo(giorno_settimana)
        if orario_inizio is None:
            print(f"Non lavori il {giorno_settimana}.")
        else:
            ore_lavorate = (orario_fine - orario_inizio).seconds / 3600
            dati_settimanali.append((giorno_settimana, data_inizio.strftime("%d/%m/%Y"), ore_lavorate))
            print(f"Hai lavorato {ore_lavorate} ore il {giorno_settimana} {data_inizio.strftime('%d/%m/%Y')}")
        data_inizio += datetime.timedelta(days=1)

    df = pd.DataFrame(dati_settimanali, columns=['Giorno', 'Data', 'Ore Lavorate'])
    somma_ore = df['Ore Lavorate'].sum()
    print(f"Totale ore lavorate nella settimana: {somma_ore}")

    df.loc[len(df.index)] = ['Totale', '', somma_ore]  # Aggiunge la riga del totale

    titolo_settimana = f"Settimana dal {data_inizio.strftime('%d/%m/%Y')} al {data_fine.strftime('%d/%m/%Y')}"

    # Crea il file Excel nella cartella specificata
    file_excel_path = os.path.join(cartella_destinazione, 'ore_lavorate_settimana.xlsx')
    writer = pd.ExcelWriter(file_excel_path, engine='openpyxl')
    df.to_excel(writer, index=False, sheet_name='Ore Lavorative')

    # Aggiunge il titolo della settimana
    workbook = writer.book
    worksheet = writer.sheets['Ore Lavorative']
    worksheet.merge_cells('A1:C1')
    worksheet['A1'] = titolo_settimana
    worksheet['A1'].style = 'Title'

    writer._save()
    print(f"File Excel creato: {file_excel_path}")

if __name__ == "__main__":
    main()
