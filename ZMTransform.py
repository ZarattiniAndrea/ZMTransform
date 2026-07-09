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

def set_TS(data: list[list[Any]]) -> dict:
    ts_files = {}
    for file in data:
        wb = openpyxl.load_workbook(file)
        coil = wb["Values"]["B5"].value
        ts_files[coil] = file
    return ts_files

#FUNZIONE CHE ESEGUE I CALCOLI 
def calculate_excel(data_BS: list[list[Any]], data_TS: list[list[Any]], result_file: str) -> None:
    print(len(data_BS), len(data_TS))
    #inserisco a dizionario tutti i valori di TS per leggerli UNA SOLA VOLTA
    ts_files = set_TS(data_TS)

    #se il file excel che contiene i risultati esiste, lo apro
    if os.path.exists(result_file):
        result_excel = openpyxl.load_workbook(result_file) #apro il workbook per i risultati    
        re = result_excel.active
    #altrimenti lo creo
    else: 
        result_excel = openpyxl.Workbook()
        re = result_excel.active
        re.title = "Dati"
        re["A1"] = "CoilID"
        re["B1"] = "DateTime"
        re["C1"] = "BS total length"
        re["D1"] = "BS Nominal"
        re["E1"] = "BS Avg"
        re["F1"] = "BS St.Dev"
        re["G1"] = "BS min"
        re["H1"] = "BS max"
        re["I1"] = "TS total length"
        re["J1"] = "TS Nominal"
        re["K1"] = "TS Avg"
        re["L1"] = "TS St.Dev"
        re["M1"] = "TS min"
        re["N1"] = "TS max"

    for i,file in enumerate(data_BS, start=2):         #ciclo sui file excel 
        excel_document_bs  = openpyxl.load_workbook(file)

        values          = excel_document_bs["Values"]
        lengthprofiles  = excel_document_bs["LengthProfiles"]

        CoildID_BS      = values["B5"].value
        DateTime        = values["B2"].value
        Nominal         = values["B4"].value
        total_length    = lengthprofiles.cell(row=lengthprofiles.max_row, column=1).value

        try: 
            #Calcoli PER BS
            rows_avg    = lengthprofiles.iter_rows(min_row=2, max_row=lengthprofiles.max_row, min_col=2, max_col=2, values_only=True)
            values_avg  = [row[0] for row in rows_avg]
            avg_bs      = sum(values_avg) / len(values_avg)
            massimo_bs  = max(values_avg)
            minimo_bs   = min(values_avg)
            dev_std_bs  = statistics.stdev(values_avg)
            
            re.cell(column=1, row=i).value = CoildID_BS
            re.cell(column=2, row=i).value = DateTime
            re.cell(column=3, row=i).value = total_length
            re.cell(column=4, row=i).value = Nominal
            re.cell(column=5, row=i).value = avg_bs
            re.cell(column=6, row=i).value = dev_std_bs
            re.cell(column=7, row=i).value = minimo_bs
            re.cell(column=8, row=i).value = massimo_bs
            chargin_bar['value'] = (i / len(data_BS)) * 100  # Aggiorna la barra di caricamento
            root.update_idletasks()
        except Exception as e: 
            write_log(f"{file}: NOT OK \n")

        if CoildID_BS in ts_files: 
            try: 
                #Calcoli PER TS
                file_ts             = ts_files[CoildID_BS]
                excel_document_ts   = openpyxl.load_workbook(file_ts)
                lengthprofiles_ts   = excel_document_ts["LengthProfiles"]
                rows_avg_ts         = lengthprofiles_ts.iter_rows(min_row=2, max_row=lengthprofiles_ts.max_row, min_col=2, max_col=2, values_only=True)
                values_avg_ts       = [row[0] for row in rows_avg_ts]
                avg_ts              = sum(values_avg_ts) / len(values_avg_ts)
                massimo_ts          = max(values_avg_ts)
                minimo_ts           = min(values_avg_ts)
                dev_std_ts          = statistics.stdev(values_avg_ts)
                total_length_ts     = lengthprofiles_ts.cell(row=lengthprofiles_ts.max_row, column=1).value

                re.cell(column=9, row=i).value  = total_length_ts
                re.cell(column=11, row=i).value = avg_ts
                re.cell(column=12, row=i).value = dev_std_ts
                re.cell(column=13, row=i).value = minimo_ts
                re.cell(column=14, row=i).value = massimo_ts
                print(f"Corrispetivo del CoilID: {CoildID_BS} è presente nel file: {file_ts}")

            except Exception as e: 
                write_log(f"{ts_files}: NOT OK FOR TS - {e} \n")
    #Chiudo i documenti
    excel_document_bs.close()
    excel_document_ts.close()
        
    # Permetto di definire dall'utente il percorso di destinazione
    result_file = filedialog.asksaveasfilename(title="Salva il file con nome",initialfile="Misurazioni_Zinco.xlsx", defaultextension=".xlsx", filetypes=[("File Excel", "*.xlsx")])
    if result_file:
        result_excel.save(result_file)


def main() -> None:

    ##### Technical sets (X BARRA DI CARICAMENTO) #####
    chargin_bar["value"] = 0 #min value
    chargin_bar["maximum"] = 100 #max value

    cartella = filedialog.askdirectory(title="Select a directory")  #apre una finestra di dialogo
    write_log(f"Cartella selezionata: {cartella}")

    ##### SE LA CARTELLA VIENE SELEZIONATA #####
    if cartella: 

    ##### Selezione della cartella BS e dei file Excel associati #####

        file_excelBS = glob.glob(cartella + "/BS/*.xlsx")  # ottengo una lista di tutti i file Excel nella cartella BS
        file_excelTS = glob.glob(cartella + "/TS/*.xlsx")  # ottengo una lista di tutti i file Excel nella cartella TS 

        numero_fileBS = len(file_excelBS) # ottengo il numero dei file excel dentro alla cartella BS
        numero_fileTS = len(file_excelTS) # ottengo il numero dei file excel dentro alla cartella TS

        write_log(f"Numero di file excel trovati in cartella BS: {numero_fileBS}")
        write_log(f"Numero di file excel trovati in cartella TS: {numero_fileTS}")

        print(f"Il file verrà salvato in: {path}")

        result_file = os.path.join(path, "Misurazioni Zinco.xlsx")

        if file_excelBS: #se il file excel è stato trovato
            if file_excelTS: 
                calculate_excel(file_excelBS, file_excelTS, result_file)
            #I calcoli vengono fatti in un thread separato in modo da non bloccare la GUI
            #thread = threading.Thread(target=calculate_excel, args=(file_excelBS,), daemon=True)
            #thread.start()
        else: 
            write_log("file excel non trovato in cartella BS")  # Stampa un messaggio se non sono stati trovati file Excel


    ##### Selezione della cartella TS e ricerca del file Excel associato #####
        '''
            if file_excelTS: #se il file excel è stato trovato
                calculate_excel(file_excelTS, result_file)
                write_log("File excel TS trovato")
            else:
                write_log("file excel non trovato in cartella TS")  # Stampa un messaggio se non sono stati trovati file Excel
        '''
        
    else:
        print("No directory selected.")  # Print a message if no directory was selected

    print("Zinc Project Started")
    write_log("ELABORAZIONE TERMINATA")


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
