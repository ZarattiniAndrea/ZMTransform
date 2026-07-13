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
import messagebox as mb
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
# ------------------- #

def write_log(message:str): 
    # Scrive un messaggio nell'area di log
    log_area.insert(tk.END, message + "\n")
    log_area.see(tk.END)
    root.update_idletasks()

def setting_TS(data: list[list[Any]]) -> dict:
    ts_files = {}
    for file in data:
        wb = openpyxl.load_workbook(file)
        values = wb["Values"]
        #Il valore del CoilID non è detto che sia sempre nella stessa posizione
        for row in values.iter_rows(min_col=1, max_col=1): 
            cella_ID = row[0]
            if cella_ID.value == "CoilID":
                coil = values.cell(row = cella_ID.row, column= cella_ID.column + 1).value
        ts_files[coil] = file
        wb.close()
    return ts_files

def user_save(result_file: str, result_excel) -> None: 
    result_file = filedialog.asksaveasfilename(title="Salva il file con nome",initialfile="Misurazioni_Zinco.xlsx", defaultextension=".xlsx", filetypes=[("File Excel", "*.xlsx")])
    try: 
        if result_file:
            result_excel.save(result_file)
            write_log("ELABORAZIONE TERMINATA")
    except PermissionError as e:
        write_log(f"Errore di permesso durante il salvataggio del file: {e}")
        mb.showerror("Errore di permesso", f"Non è possibile salvare il file. Controlla se il file è aperto in un'altra applicazione o se hai i permessi necessari. \nDettagli: {e}")
    except Exception as e:
        write_log(f"Errore durante il salvataggio del file: {e}")


#FUNZIONE CHE ESEGUE I CALCOLI 
def calculate_excel(data_BS: list[list[Any]], data_TS: list[list[Any]], result_file: str) -> None:
    print(len(data_BS), len(data_TS))
    #inserisco a dizionario tutti i valori di TS per leggerli UNA SOLA VOLTA
    ts_files = setting_TS(data_TS)

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
        try:
            excel_document_bs  = openpyxl.load_workbook(file)

            values          = excel_document_bs["Values"]
            lengthprofiles  = excel_document_bs["LengthProfiles"]


            for row in values.iter_rows(min_col=1, max_col=1): 
                cella_ID = row[0]
                if cella_ID.value == "CoilID":
                    CoilID_BS = values.cell(row = cella_ID.row, column= cella_ID.column + 1).value
            #CoilID_BS      = values["B5"].value
            DateTime        = values["B2"].value
            
            
            #Come con CoilID, ma con nominal
            Nominal = None
            for row in values.iter_rows(min_col=1, max_col=1): 
                cella_nominal = row[0]
                if cella_nominal.value == "Coating_BS_Nominal":
                    Nominal = values.cell(row = cella_nominal.row, column= cella_nominal.column + 1).value
            
            total_length    = lengthprofiles.cell(row=lengthprofiles.max_row, column=1).value

            try: 
                #Calcoli PER BS
                rows_avg    = lengthprofiles.iter_rows(min_row=2, max_row=lengthprofiles.max_row, min_col=2, max_col=2, values_only=True)
                values_avg  = [row[0] for row in rows_avg]
                avg_bs      = sum(values_avg) / len(values_avg)
                massimo_bs  = max(values_avg)
                minimo_bs   = min(values_avg)
                dev_std_bs  = statistics.stdev(values_avg)
                
                re.cell(column=1, row=i).value = CoilID_BS
                re.cell(column=2, row=i).value = DateTime
                re.cell(column=3, row=i).value = total_length
                re.cell(column=4, row=i).value = Nominal
                re.cell(column=5, row=i).value = avg_bs
                re.cell(column=6, row=i).value = dev_std_bs
                re.cell(column=7, row=i).value = minimo_bs
                re.cell(column=8, row=i).value = massimo_bs
            except Exception as e: 
                write_log(f"{file}: NOT OK BS - {e} \n")
                continue
            finally:
                chargin_bar['value'] = (i / len(data_BS)) * 100  # Aggiorna la barra di caricamento
                root.update_idletasks()

            if CoilID_BS in ts_files: 
                try: 
                    #Calcoli PER TS
                    file_ts             = ts_files[CoilID_BS]
                    excel_document_ts   = openpyxl.load_workbook(file_ts)
                    values_ts           = excel_document_ts["Values"]

                    for row in values_ts.iter_rows(min_col=1, max_col=1): 
                        cella_nominal_ts = row[0]
                        if cella_nominal_ts.value == "Coating_TS_Nominal":
                            nominal_ts = values_ts.cell(row = cella_nominal_ts.row, column= cella_nominal_ts.column + 1).value
                    
                    lengthprofiles_ts   = excel_document_ts["LengthProfiles"]
                    rows_avg_ts         = lengthprofiles_ts.iter_rows(min_row=2, max_row=lengthprofiles_ts.max_row, min_col=2, max_col=2, values_only=True)
                    values_avg_ts       = [row[0] for row in rows_avg_ts]
                    avg_ts              = sum(values_avg_ts) / len(values_avg_ts)
                    massimo_ts          = max(values_avg_ts)
                    minimo_ts           = min(values_avg_ts)
                    dev_std_ts          = statistics.stdev(values_avg_ts)
                    total_length_ts     = lengthprofiles_ts.cell(row=lengthprofiles_ts.max_row, column=1).value

                    re.cell(column=9, row=i).value  = total_length_ts
                    re.cell(column=10, row=i).value = nominal_ts
                    re.cell(column=11, row=i).value = avg_ts
                    re.cell(column=12, row=i).value = dev_std_ts
                    re.cell(column=13, row=i).value = minimo_ts
                    re.cell(column=14, row=i).value = massimo_ts
                    print(f"Corrispetivo del CoilID: {CoilID_BS} è presente nel file: {file_ts}")

                except Exception as e: 
                    write_log(f"{file_ts}: NOT OK FOR TS - {e} \n")
                    continue
                finally: 
                    excel_document_ts.close()
        except PermissionError as e:
            write_log(f"Errore di permesso durante il salvataggio del file: {e}")
            mb.showerror("Errore di permesso", f"Non è possibile salvare il file. Controlla se il file è aperto in un'altra applicazione o se hai i permessi necessari. \nDettagli: {e}")
                
    excel_document_bs.close()

    user_save(result_file, result_excel)  # Chiamo la funzione per salvare il file con il percorso scelto dall'utente

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
                #calculate_excel(file_excelBS, file_excelTS, result_file)
                #I calcoli vengono fatti in un thread separato in modo da non bloccare la GUI
                thread = threading.Thread(target=calculate_excel, args=(file_excelBS, file_excelTS, result_file), daemon=True)
                thread.start()
        else: 
            write_log("file excel non trovato in cartella BS")  # Stampa un messaggio se non sono stati trovati file Excel

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