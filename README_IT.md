# DICOM Metadata Extractor - Archivio Esami RX

Estrae i metadati DICOM da una struttura di cartelle OsiriX e produce un foglio Excel formattato ("Archivio Esami RX") con una riga per studio (ordine esame paziente).

## Funzionalita

- **Raggruppamento per studio** - Raggruppa i file DICOM per `StudyInstanceUID`, in modo che ogni riga rappresenti un esame, non una singola immagine.
- **Rilevamento intelligente del tipo di esame** - Determina il `TIPO ESAME` da diversi tag DICOM in ordine di priorita, partendo da `AcquisitionDeviceProcessingDescription (0018,1400)`. Aggiunge automaticamente il prefisso `RX` se mancante.
- **Conversione dose** - Legge `ImageAreaDoseProduct (0018,115E)` e converte da dGy-cm^2 a Gy-m^2, con rilevamento automatico per i produttori che salvano gia i valori in Gy-m^2.
- **Categorizzazione per eta** - Analizza `PatientAge` e classifica in quattro gruppi: `0-1`, `1-16`, `16-60`, `>60`.
- **Filtro per anno** - Prompt interattivo per selezionare quale/i anno/i esportare quando sono presenti piu anni.
- **Filtro per modalita** - Elabora solo le modalita radiografiche (DX, CR, RF, XA, MG, DR, PX, IO, OP) e ignora SR/KO/PR/SC.
- **Output Excel formattato** - Intestazioni stilizzate, filtro automatico, blocco riquadri e larghezze colonne appropriate.

## Requisiti

- Python 3.8+
- Dipendenze:
  ```
  pip install pydicom pandas openpyxl
  ```

## Utilizzo

**Base** (produce `Archivio_Esami_RX.xlsx`):
```bash
python3 dicom_to_excel.py /percorso/cartella_OsiriX
```

**Nome file di output personalizzato:**
```bash
python3 dicom_to_excel.py /percorso/cartella_OsiriX --output MioArchivio.xlsx
```

**Modalita debug** (ispeziona i tag DICOM grezzi senza generare l'Excel):
```bash
python3 dicom_to_excel.py /percorso/cartella_OsiriX --debug
python3 dicom_to_excel.py /percorso/cartella_OsiriX --debug --debug-count 20
```

## Colonne di Output

| Colonna | Descrizione |
|---|---|
| DATA ESAME | Data dell'esame (GG/MM/AA) |
| TIPO ESAME | Tipo di esame (es. `RX TORACE`, `RX MANO`) |
| 0 - 1 | Fascia d'eta: da 0 a 1 anno |
| 1 - 16 | Fascia d'eta: da 1 a 16 anni |
| 16 - 60 | Fascia d'eta: da 16 a 60 anni |
| > 60 | Fascia d'eta: oltre 60 anni |
| DOSE | Prodotto Dose-Area in Gy-m^2 (notazione scientifica) |

## Selezione Anno

Quando la cartella analizzata contiene studi di piu anni, lo script chiede di scegliere quale/i anno/i esportare:

```
Anni trovati nei dati:
  [1] 2023  (142 studi)
  [2] 2024  (208 studi)

Inserisci anno/i da esportare (es. 2024 o 2023,2024), oppure premi Invio per tutti:
```

## Ordine di Risoluzione del Tipo di Esame

Lo script controlla questi tag DICOM in ordine e utilizza il primo valore non vuoto:

1. `AcquisitionDeviceProcessingDescription (0018,1400)`
2. `SeriesDescription (0008,103E)`
3. `StudyDescription (0008,1030)`
4. `RequestedProcedureDescription (0032,1060)`
5. `PerformedProcedureStepDescription (0040,0254)`
6. `ProtocolName (0018,1030)`
7. `BodyPartExamined (0018,0015)`
8. Nome della cartella genitore (fallback OsiriX)
