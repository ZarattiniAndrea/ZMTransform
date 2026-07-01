#!/usr/bin/env python3
"""
Zinc Project - Measurement Analysis

A project for zinc measurement and analysis.
"""

import sys
import tkinter as tk
import tkinter.filedialog
import pandas as pd
import glob
from typing import Any


def main() -> None:
    root = tk.Tk()  # Create the main application window
    root.title("Zinc Project")  
    root.geometry("400x300") 

    cartella = tk.filedialog.askdirectory(title="Select a directory")  #apre una finestra di dialogo


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


if __name__ == "__main__":
    sys.exit(main())
