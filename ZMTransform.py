#!/usr/bin/env python3
"""
Zinc Project - Measurement Analysis

A project for zinc measurement and analysis.
"""

import sys, datetime, os, re, time, configparser
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
    #excel_document = openpyxl.load_workbook(data[0])                                    #qui c'è la lista di tutti i file excel che passo alla funzione di calcolo
    result_file = os.path.join(path, "Misurazioni Zinco.xlsx")
    result_excel = openpyxl.load_workbook(result_file)             #apro il workbook per i risultati
    re = result_excel.active

    #print(excel_document.sheetnames)                    #mostro i fogli disponibili
    
    #Values = excel_document["Values"]                   #Foglio Values 
    #LengthProfiles = excel_document["LengthProfiles"]   #Foglio LengthProfiles

    for i,file in enumerate(data, start=2):
        excel_document = openpyxl.load_workbook(file)
        
        values = excel_document["Values"]
        lengthprofiles = excel_document["LengthProfiles"]

        CoildID = values["B5"].value
        DateTime = values["B2"].value
        total_length = lengthprofiles.cell(row=lengthprofiles.max_row, column=1).value

        re.cell(column=1, row=i).value = CoildID
        re.cell(column=2, row=i).value = DateTime
        re.cell(column=3, row=i).value = total_length
        chargin_bar['value'] = (i / len(data)) * 100  # Aggiorna la barra di caricamento
        root.update_idletasks()

    result_excel.save(result_file)


    

    #provo a fare una versione generica in cui non so a priori il numero di righe da leggere
    '''
    righe = tuple(ws.iter_rows(min_row = 2, max_row = ws.max_row, min_col = 1, max_col = 4, values_only = True))
    i = 1
    for riga in righe: 
        coil_id = riga[0]
        max_value = riga[1] 
        min_value = riga[2]
        result = round((max_value + min_value) / 2, 3)
        ws.cell(column = 4, row = i+1).value = result
        excel_document.save(data[0])
        chargin_bar['value'] = (i / len(righe)) * 100  # Aggiorna la barra di caricamento
        write_log(f"{_data_ora} - Coil ID: {coil_id}, Max Value: {max_value}, Min Value: {min_value}, Calculated Value: {result} \n")
        i = i + 1
    '''

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
        #write_log(f"File excel trovati in cartella BS: {file_excelBS}")

        print(f"Il file verrà salvato in: {path}")

        create_excel_file(Path(path)) #Creo il file excel (in seguito dovrò controllare se c'è gia o no)

        if file_excelBS: #se il file excel è stato trovato
            #write_log(f"File excel selezionato in cartella BS: {file_excelBS[0]}")
            calculate_excel(file_excelBS)

        else: 
            write_log("file excel non trovato in cartella BS")  # Stampa un messaggio se non sono stati trovati file Excel


    ##### Selezione della cartella TS e ricerca del file Excel associato #####

        file_excelTS = glob.glob(cartella + "/TS/*.xlsx")  # ottengo una lista di tutti i file Excel nella cartella selezionata 
        numero_file = len(file_excelTS) # ottengo il numero dei file excel dentro alla cartella TS
        write_log(f"Numero di file excel trovati in cartella TS: {numero_file}")
        #write_log(f"File excel trovati in cartella TS: {file_excelTS}")

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
