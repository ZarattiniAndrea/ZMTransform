# Project Requirements

## Overview
1)	Avremmo necessità che venisse realizzato un programma in grado di estrarre dati aggregati da file CSV
2)	Ho appena ricevuto i file CSV che come ci aspettavamo sono diversi dall’excel. Purtroppo i file i vengono duplicati:
-uno con i dati di tracciabilità 
-uno con le misurazioni di zinco che però rimane privo di associazione al numero coi. 
Ho l’impressione che questi CSV siano molto più complessi da gestire che il file Excel che abbiamo visto oggi. Mi confermate che è meglio lavorare sull’excel? 

3)	Il programma dovrà interrogare i file presenti nelle cartelle che verranno create con le misurazioni mensile. Per ogni rotolo è presente:
Un file nella cartella TS
Un file nella cartella BS

4)	Il programma dovrà popolare le righe del file Misurazioni Zinco aggiungendo una riga per ogni identificativo coil. Se possibile sarebbe gradito evitare le duplicazioni di coil ID non aggiungendo la riga nel caso il coil id fosse già presente. Se però allunga i tempi di elaborazione non è un problema anche accettare la duplicazione. 


## Functional Requirements
1)	Run del processo
2)	Seleziono la cartella del mese di calcolo
3)	All’interno della cartella del mese dovranno esserci obbligatoriamente le sottocartelle BS e TS (altrimenti fermo il run)
4)	Ciascuna cartella BS e TS conterrà i file CSV da analizzare (oppure XLSX)
5)	Ogni file xlsx  ha 2 fogli (nel caso di CSV molto probabilmente tutte le info saranno su un unico foglio)
        a.	La prima “….” Contiene l’ID del coil e la data di produzione
        b.	La seconda “…”  contiene due colonne: Lunghezza e Spessore rivestimento
6)	Il programma partendo dalla cartella BS aprirà ciascun file e
        a.	Salverà il COIL ID e la DATA ORA di produzione in un file dei risultati (un coil una riga) Il file sarà un XLSX
        b.	Calcolerà la media dello spessore di rivestimento, il min, il max e la dev std, la lunghezza massima e il numero dei campioni e li salverà nella stessa riga del punto precedente
        c.	Chiuderà il file e aprirà un nuovo file sempre dentro la cartella BS
        d.	Terminati i file della cartella BS passerà ad esaminare i file della cartella TS ripetendo le operazioni da A a C. 
            Attenzione ci aspettiamo che per ciascun file della cartella TS ci sia già una riga nel file dei risultati. I valori di TS dovranno essere aggiunti alla riga già presente.
            Nel caso non ci sia una riga allora aggiungerne una nuova.
        e.	Ogni volta che apro il file dei risultati in caso di informazioni già presenti per BS o TS allora skip al prossimo file.
        f.	LOG FILE: creare un log file txt per ciascun lancio del programma oppure un log file per giorno nel quale trascrivere man mano le operazioni svolte
                    i.	Avvio calcolo
                    ii.	Selezionata cartella ….\maggio\
                    iii.	ROTOLO XYZ creata riga 
                    iv.	ROTOLO xyz aggiunti dati BS
                    v.	ROTOLO xyz aggiunti dati TS
                    vi.	ROTOLO XYZ riga già presente skip BS
                    vii.	ROTOLO XYZ riga già presente skip TS
                    viii.	…
                    ix.	Esaminati 100 file BS e 98 file TS
                    x.	Skippati 3 file BS e 4 file TS
                    xi.	Fine calcolo

## Non-functional Requirements
- Ensure data integrity and accuracy.
- Maintain responsiveness and reliability.
- Support easy deployment and maintenance.

## Stakeholders
- Project owner
- Engineering team
- Quality assurance
- End users
