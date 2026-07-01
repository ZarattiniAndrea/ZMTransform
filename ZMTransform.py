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

_message: str = "Zinc Project - Measurement Analysis"

root = tk.Tk()  # Create the main application window
root.title("Zinc Project")  
root.geometry("400x300")

# --- Area di Log --- #
log_area = scrolledtext.ScrolledText(root, width=50, height=10)
log_area.pack(padx=10, pady=10)

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

        if file_excelBS: #se il file excel è stato trovato
            print("Excel files found:") 
            print(file_excelBS)
            print(f"il nome del file excel è: {file_excelBS[0]}")
        else: 
            print("file excel non trovato in cartella BS")  # Stampa un messaggio se non sono stati trovati file Excel

        print(f"Selected directory: {cartella}")
        print(f"Excel files found: {len(file_excelBS)}")  

    ##### Selezione della cartella TS e ricerca del file Excel associato #####

        file_excelTS = glob.glob(cartella + "/TS/*.xlsx")  # ottengo una lista di tutti i file Excel nella cartella selezionata 

        if file_excelTS: #se il file excel è stato trovato
            print("Excel files found:")  
            print(file_excelTS)  
            print(f"il nome del file excel è: {file_excelTS[0]}") 
        else:
            print("file excel non trovato in cartella TS")  # Stampa un messaggio se non sono stati trovati file Excel
            
        print(f"Selected directory: {cartella}")
        
        print(f"Excel files found: {len(file_excelTS)}")
    else:
        print("No directory selected.")  # Print a message if no directory was selected

    print("Zinc Project Started")



btn = tk.Button(root, text="Seleziona la cartella", command=main)
btn.pack(pady=10)
root.mainloop()  # Avvia il loop principale dell'applicazione Tkinter
