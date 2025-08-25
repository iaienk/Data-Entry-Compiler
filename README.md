[![Python](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://www.python.org/downloads/)
[![Status](https://img.shields.io/badge/status-demo-orange.svg)]()
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Privacy](https://img.shields.io/badge/data-sensitive%20files%20excluded-critical.svg)]()

🇮🇹 Come funziona

Questo repository contiene solo lo script (data_entry_compiler.py).
I file Excel necessari contengono dati sensibili e non sono inclusi:

Sequenza_Giornaliera.xlsx (input)

MoQ_10_Data entry.xlsx (template con il foglio Data_Entry)

Passi (copia & incolla)

Prepara la cartella di lavoro
Metti nella stessa cartella dello script questi due file (non nel repo pubblico):

Sequenza_Giornaliera.xlsx

MoQ_10_Data entry.xlsx

Esegui lo script (dalla cartella):

python data_entry_compiler.py


Lo script:

legge le colonne richieste da Sequenza_Giornaliera.xlsx
(es.: Tipologia in salita, Progressivo, Data, Linee, Tipologia in discesa, Bus, Direzione, Fermata di salita, Fermata di discesa, Meteo, ID, Rilevatore, Tipo di giorno)

apre il template MoQ_10_Data entry.xlsx (foglio Data_Entry)

compila le righe, copia bordi/bold dal template, pulisce eventuali righe in eccesso

salva un file di output con nome tipo: 01_YYYYMMDD.xlsx
(regola predefinita: 01_ se l’ID inizia per “1”, altrimenti 02_)

File output
L’output viene salvato nella stessa cartella dello script (o nella cartella indicata da riga di comando).

Opzioni utili (riga di comando)
# specifica percorsi diversi
python data_entry_compiler.py \
  --sequence "path/to/Sequenza_Giornaliera.xlsx" \
  --template "path/to/MoQ_10_Data entry.xlsx" \
  --outdir "output/"

# formati data e prefissi
python data_entry_compiler.py \
  --date-format "%d/%m/%Y" \
  --filedate-format "%Y%m%d" \
  --prefix-if-id-startswith "1" \
  --prefix-true "01_" \
  --prefix-false "02_"

Note e requisiti

Requisiti: pandas, openpyxl

Le immagini/screenshot non servono per questo script.

Se mancano colonne richieste o fogli nel template, lo script segnala l’errore in console.

Troubleshooting

“Mancano colonne richieste” → verifica i nomi colonna nell’input (Sequenza_Giornaliera.xlsx).

“Template non trovato / foglio non trovato” → controlla percorso e nome foglio (Data_Entry).

Bordi/Grassetto → lo script replica bordi/bold dal template (stesse posizioni); se il template cambia, aggiornare di conseguenza.

🇬🇧 How it works

This repository includes only the script (data_entry_compiler.py).
The Excel files contain sensitive data and are not included:

Sequenza_Giornaliera.xlsx (input)

MoQ_10_Data entry.xlsx (template with Data_Entry sheet)

Steps (copy & paste)

Prepare the working folder
Place these files in the same directory as the script (do not add them to the public repo):

Sequenza_Giornaliera.xlsx

MoQ_10_Data entry.xlsx

Run the script:

python data_entry_compiler.py


The script will:

read required columns from Sequenza_Giornaliera.xlsx

open the template MoQ_10_Data entry.xlsx (Data_Entry sheet)

fill rows, copy borders/bold from the template, and clean leftover rows

save an output file named like 01_YYYYMMDD.xlsx
(default rule: 01_ if ID starts with “1”, else 02_)

Output file
Saved in the same folder (or in --outdir if provided).

Useful CLI options
python data_entry_compiler.py \
  --sequence "path/to/Sequenza_Giornaliera.xlsx" \
  --template "path/to/MoQ_10_Data entry.xlsx" \
  --outdir "output/"

python data_entry_compiler.py \
  --date-format "%d/%m/%Y" \
  --filedate-format "%Y%m%d" \
  --prefix-if-id-startswith "1" \
  --prefix-true "01_" \
  --prefix-false "02_"

Notes & requirements

Requirements: pandas, openpyxl

No images needed for this script.

If columns/sheets are missing, the script will print a clear error.

Project structure (suggested)
.
├─ data_entry_compiler.py
├─ README.md
├─ .gitignore
└─ (place locally, do NOT commit)
   ├─ Sequenza_Giornaliera.xlsx
   └─ MoQ_10_Data entry.xlsx
