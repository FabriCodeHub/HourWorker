# HourWorker

Questo script in Python permette di registrare e calcolare le ore lavorative settimanali, generando un file Excel con i risultati. L'utente inserisce gli orari di inizio e fine lavoro per ciascun giorno della settimana, e lo script calcola le ore totali lavorate e salva i dati in un file Excel.

## Funzionalità

1. **Inserimento settimana**: La funzione `inserisci_settimana` richiede all'utente di inserire la data di inizio e fine della settimana.
2. **Inserimento orario lavorativo**: La funzione `inserisci_orario_lavorativo` richiede all'utente di inserire l'orario di inizio e fine lavoro per ciascun giorno della settimana.
3. **Calcolo ore lavorate**: Lo script calcola le ore lavorate per ciascun giorno e le somma per ottenere il totale settimanale.
4. **Generazione file Excel**: I dati settimanali vengono salvati in un file Excel nella cartella specificata.

## Installazione

1. Clonare il repository:
    ```bash
    git clone https://github.com/tuo-username/hourworker.git
    cd hourworker
    ```

2. Creare un ambiente virtuale:
    ```bash
    python -m venv env
    ```

3. Attivare l'ambiente virtuale:

    - Su Windows:
        ```bash
        .\env\Scripts\activate
        ```

    - Su macOS/Linux:
        ```bash
        source env/bin/activate
        ```

4. Installare le dipendenze:
    ```bash
    pip install -r requirements.txt
    ```

## Utilizzo

1. Eseguire lo script:
    ```bash
    python orelavorate.py
    ```

2. Seguire le istruzioni a schermo per inserire la data di inizio e fine della settimana, e gli orari di inizio e fine lavoro per ciascun giorno della settimana.

3. Lo script genererà un file Excel con i dati delle ore lavorate nella cartella `C:\Users\<tuo-username>\Desktop\HourWorker`.

## Dipendenze

- `datetime`
- `pandas`
- `openpyxl`
- `pyfiglet`
- `os`

## Dettagli delle Funzioni

### normalizza_orario(orario)

Questa funzione normalizza il formato dell'orario sostituendo il punto con i due punti.

### inserisci_settimana()

Richiede all'utente di inserire la data di inizio e fine della settimana e restituisce queste date.

### inserisci_orario_lavorativo(giorno_settimana)

Richiede all'utente di inserire l'orario di inizio e fine lavoro per un dato giorno della settimana. Se l'utente non lavora quel giorno, restituisce `None`.

### main()

La funzione principale che coordina l'intero processo: visualizza il titolo, richiede l'inserimento delle date e degli orari, calcola le ore lavorate, genera il file Excel e stampa il percorso del file creato.

## Contributi

I contributi sono benvenuti! Sentitevi liberi di aprire un problema o una pull request.

## Licenza

Questo progetto è sotto licenza MIT. Vedi il file LICENSE per maggiori dettagli.

