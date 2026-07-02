#!/usr/bin/env python3
"""
Zinc Project - Measurement Analysis

A project for zinc measurement and analysis.
"""

import sys, datetime, os, re
import tkinter as tk
from tkinter import scrolledtext, filedialog
from unittest import result
import pandas as pd
import glob
from typing import Any
import openpyxl

_message: str = "Zinc Project - Measurement Analysis"
_data_ora: str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

root = tk.Tk()  # Create the main application window
root.title("Zinc Project")  
root.geometry("800x500")

# --- Area di Log --- #
log_area = scrolledtext.ScrolledText(root, wrap="word") # in questo modo l'area di log avrà una barra di scorrimento verticale
log_area.pack(expand=True, fill="both", padx=10, pady=10)

def write_log(message:str): 
    # Scrive un messaggio nell'area di log
    log_area.insert(tk.END, message + "\n")
    log_area.see(tk.END)
    root.update_idletasks()



def main() -> None:
    cartella = filedialog.askdirectory(title="Select a directory")  #apre una finestra di dialogo
    write_log(f"Cartella selezionata: {cartella}")

    ##### SE LA CARTELLA VIENE SELEZIONATA #####
    if cartella: 

    ##### Selezione della cartella BS e ricerca del file Excel associato #####

        file_excelBS = glob.glob(cartella + "/BS/*.xlsx")  # ottengo una lista di tutti i file Excel nella cartella selezionata 
        numero_file = len(file_excelBS) # ottengo il numero dei file excel dentro alla cartella BS
        write_log(f"Numero di file excel trovati in cartella BS: {numero_file}")
        write_log(f"File excel trovati in cartella BS: {file_excelBS}")
        

        if file_excelBS: #se il file excel è stato trovato
            write_log(f"File excel selezionato in cartella BS: {file_excelBS[0]}")
            excel_document = openpyxl.load_workbook(file_excelBS[0]) 
            ws = excel_document.active
            '''
            write_log(f"Valore della cella A1 nel file Excel BS: {ws['A1'].value}")
            valA2 = ws.cell(column = 1, row = 2).value
            valB2 = ws.cell(column = 2, row = 2).value
            valC2 = ws.cell(column = 3, row = 2).value
            valD2 = valB2 + valC2 / 2 
            ws.cell(column = 4, row = 2).value = valD2
            excel_document.save(file_excelBS[0])
            write_log(f"{_data_ora} - Valore calcolato per coil ID {valA2} : {valD2}")
            '''

            #provo a fare una versione generica in cui non so a priori il numero di righe da leggere
            righe = tuple(ws.iter_rows(min_row = 2, max_row = ws.max_row, min_col = 1, max_col = 4, values_only = True))
            i = 1
            for riga in righe:
                coil_id = riga[0]
                max_value = riga[1]
                print(max_value)
                min_value = riga[2]
                print(min_value)
                result = round((max_value + min_value) / 2, 3)
                print(result)
                ws.cell(column = 4, row = i+1).value = result
                excel_document.save(file_excelBS[0])
                write_log(f"{_data_ora} - Coil ID: {coil_id}, Max Value: {max_value}, Min Value: {min_value}, Calculated Value: {result} \n")
                i = i + 1

        else: 
            write_log("file excel non trovato in cartella BS")  # Stampa un messaggio se non sono stati trovati file Excel


    ##### Selezione della cartella TS e ricerca del file Excel associato #####

        file_excelTS = glob.glob(cartella + "/TS/*.xlsx")  # ottengo una lista di tutti i file Excel nella cartella selezionata 
        numero_file = len(file_excelTS) # ottengo il numero dei file excel dentro alla cartella TS
        write_log(f"Numero di file excel trovati in cartella TS: {numero_file}")
        write_log(f"File excel trovati in cartella TS: {file_excelTS}")

        if file_excelTS: #se il file excel è stato trovato
            write_log(f"File excel selezionato in cartella TS: {file_excelTS[0]}")
        else:
            write_log("file excel non trovato in cartella TS")  # Stampa un messaggio se non sono stati trovati file Excel
            
        
    else:
        print("No directory selected.")  # Print a message if no directory was selected

    print("Zinc Project Started")


frame_bottoni = tk.Frame(root)
frame_bottoni.pack(pady=10)

btn1 = tk.Button(frame_bottoni, text="Seleziona la cartella", command=main)
btn1.grid(row=0, column=0, sticky="w", padx=5)
btn2 = tk.Button(frame_bottoni, text="Esci", command=root.quit)
btn2.grid(row=0, column=1, sticky="e", padx=5)

root.mainloop()  # Avvia il loop principale dell'applicazione Tkinter
