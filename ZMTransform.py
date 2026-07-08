#!/usr/bin/env python3
"""
Zinc Project - Measurement Analysis

A project for zinc measurement and analysis.
"""

import sys, datetime, os, re, time, configparser, statistics, threading
import tkinter as tk
from tkinter import scrolledtext, filedialog, ttk
from unittest import result
from pathlib import Path
import pandas as pd
import glob
from typing import Any
import openpyxl
from ctypes import windll

_message: str = "Zinc Project - Measurement Analysis"
_data_ora: str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

config = configparser.ConfigParser()
config.read('config.ini')
path = config['Impostazioni']['path']
print(f"Path from config.ini: {path}")

try: 
    windll.shcore.SetProcessDpiAwareness(1)  # Imposta la consapevolezza DPI per migliorare la resa grafica su schermi ad alta risoluzione
except Exception as e:
    print(f"Errore durante l'impostazione della consapevolezza DPI: {e}")

### Main Application Window ###
root = tk.Tk()
root.title("Zinc Project")  
root.geometry("1000x800")

# --- Area di Log --- #
log_area = scrolledtext.ScrolledText(root, wrap="word") # in questo modo l'area di log avrà una barra di scorrimento verticale
log_area.pack(expand=True, fill="both", padx=10, pady=10)

def write_log(message:str): 
    # Scrive un messaggio nell'area di log
    log_area.insert(tk.END, message + "\n")
    log_area.see(tk.END)
    root.update_idletasks()


def create_excel_file(file_path: str) -> None: #creo l'intestazione del file excel di destinazione
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Dati"
    ws["A1"] = "CoilID"
    ws["B1"] = "DateTime"
    ws["C1"] = "BS total length"
    ws["D1"] = "BS Nominal"
    ws["E1"] = "BS Avg"
    ws["F1"] = "BS St.Dev"
    ws["G1"] = "BS min"
    ws["H1"] = "BS max"
    ws["I1"] = "TS total length"
    ws["J1"] = "TS Nominal"
    ws["K1"] = "TS Avg"
    ws["L1"] = "TS St.Dev"
    ws["M1"] = "TS min"
    ws["N1"] = "TS max  "
    percorso = file_path / "Misurazioni Zinco.xlsx"
    wb.save(percorso)

def calculate_excel(data: list[list[Any]]) -> None:                                     #funzione che esegue i calcoli
    #excel_document = openpyxl.load_workbook(data[0])                                   #qui c'è la lista di tutti i file excel che passo alla funzione di calcolo
    result_file = os.path.join(path, "Misurazioni Zinco.xlsx")
    result_excel = openpyxl.load_workbook(result_file)                                  #apro il workbook per i risultati
    re = result_excel.active

    for i,file in enumerate(data, start=2):              #ciclo sui file excel 
        excel_document = openpyxl.load_workbook(file)
        
        values = excel_document["Values"]
        lengthprofiles = excel_document["LengthProfiles"]

        CoildID = values["B5"].value
        DateTime = values["B2"].value
        Nominal = values["B4"].value
        total_length = lengthprofiles.cell(row=lengthprofiles.max_row, column=1).value

        try: 
            #Calcoli
            rows_avg = lengthprofiles.iter_rows(min_row=2, max_row=lengthprofiles.max_row, min_col=2, max_col=2, values_only=True)
            values_avg  = [row[0] for row in rows_avg]
            avg_bs      = sum(values_avg) / len(values_avg)
            massimo     = max(values_avg)
            minimo      = min(values_avg)
            dev_std     = statistics.stdev(values_avg)
            
            re.cell(column=1, row=i).value = CoildID
            re.cell(column=2, row=i).value = DateTime
            re.cell(column=3, row=i).value = total_length
            re.cell(column=4, row=i).value = Nominal
            re.cell(column=5, row=i).value = avg_bs
            re.cell(column=6, row=i).value = dev_std
            re.cell(column=7, row=i).value = minimo
            re.cell(column=8, row=i).value = massimo
            chargin_bar['value'] = (i / len(data)) * 100  # Aggiorna la barra di caricamento
            root.update_idletasks()
        except Exception as e: 
            write_log(f"{file}: NOT OK \n")

    # Permetto di definire dall'utente il percorso di destinazione
    result_file = filedialog.asksaveasfilename(title="Salva il file con nome",initialfile="Misuazioni_Zinco.xlsx", defaultextension=".xlsx", filetypes=[("File Excel", "*.xlsx")])
    if result_file:
        result_excel.save(result_file)

    write_log("Elaborazione dei file in BS completata")

def main() -> None:

    ##### Technical sets (X BARRA DI CARICAMENTO) #####
    chargin_bar["value"] = 0 #min value
    chargin_bar["maximum"] = 100 #max value

    cartella = filedialog.askdirectory(title="Select a directory")  #apre una finestra di dialogo
    write_log(f"Cartella selezionata: {cartella}")

    ##### SE LA CARTELLA VIENE SELEZIONATA #####
    if cartella: 

    ##### Selezione della cartella BS e dei file Excel associati #####

        file_excelBS = glob.glob(cartella + "/BS/*.xlsx")  # ottengo una lista di tutti i file Excel nella cartella selezionata 
        numero_file = len(file_excelBS) # ottengo il numero dei file excel dentro alla cartella BS

        write_log(f"Numero di file excel trovati in cartella BS: {numero_file}")

        print(f"Il file verrà salvato in: {path}")

        create_excel_file(Path(path)) #Creo il file excel (in seguito dovrò controllare se c'è gia o no)

        if file_excelBS: #se il file excel è stato trovato
            #I calcoli vengono fatti in un thread separato in modo da non bloccare la GUI
            thread = threading.Thread(target=calculate_excel, args=(file_excelBS,), daemon=True)
            thread.start()
        else: 
            write_log("file excel non trovato in cartella BS")  # Stampa un messaggio se non sono stati trovati file Excel


    ##### Selezione della cartella TS e ricerca del file Excel associato #####

        file_excelTS = glob.glob(cartella + "/TS/*.xlsx")  # ottengo una lista di tutti i file Excel nella cartella selezionata 
        numero_file = len(file_excelTS) # ottengo il numero dei file excel dentro alla cartella TS
        write_log(f"Numero di file excel trovati in cartella TS: {numero_file}")

        if file_excelTS: #se il file excel è stato trovato
            write_log(f"File excel selezionato in cartella TS: {file_excelTS[0]}")
        else:
            write_log("file excel non trovato in cartella TS")  # Stampa un messaggio se non sono stati trovati file Excel
            
        
    else:
        print("No directory selected.")  # Print a message if no directory was selected

    print("Zinc Project Started")


frame_bottoni = tk.Frame(root)
frame_bottoni.pack(pady=10)

#creazione barra di caricamento
chargin_bar = ttk.Progressbar(frame_bottoni, orient = "horizontal", length = 200, mode = "determinate")
chargin_bar.grid(row=0, column=2, padx=5)

btn1 = tk.Button(frame_bottoni, text="Seleziona la cartella", command=main)
btn1.grid(row=0, column=0, sticky="w", padx=5)
btn2 = tk.Button(frame_bottoni, text="Esci", command=root.quit)
btn2.grid(row=0, column=1, sticky="e", padx=5)

root.mainloop()  # Avvia il loop principale dell'applicazione Tkinter
