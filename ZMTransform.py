#!/usr/bin/env python3
"""
Zinc Project - Measurement Analysis

A project for zinc measurement and analysis.
"""

import sys
import tkinter as tk
from tkinter import scrolledtext, filedialog
import pandas as pd
import glob
from typing import Any
import openpyxl

_message: str = "Zinc Project - Measurement Analysis"

root = tk.Tk()  # Create the main application window
root.title("Zinc Project")  
root.geometry("600x500")

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
            sheet = excel_document['Sheet1']
            write_log(f"Valore della cella A1 nel file Excel BS: {sheet.cell(column = 1, row = 1).value}")
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
