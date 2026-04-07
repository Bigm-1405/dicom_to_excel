# DICOM Metadata Extractor - Archivio Esami RX

Extracts DICOM metadata from an OsiriX folder structure and produces a formatted Excel spreadsheet ("Archivio Esami RX") with one row per study (patient exam order).

## Features

- **Study-level grouping** - Groups DICOM files by `StudyInstanceUID`, so each row represents one exam, not one image.
- **Smart exam-type detection** - Resolves `TIPO ESAME` from multiple DICOM tags in priority order, starting with `AcquisitionDeviceProcessingDescription (0018,1400)`. Automatically prepends `RX` when missing.
- **Dose conversion** - Reads `ImageAreaDoseProduct (0018,115E)` and converts from dGy-cm^2 to Gy-m^2, with auto-detection for vendors that already store values in Gy-m^2.
- **Age categorization** - Parses `PatientAge` and buckets into four groups: `0-1`, `1-16`, `16-60`, `>60`.
- **Year filter** - Interactive prompt to select which year(s) to export when multiple years are present.
- **Modality filtering** - Processes only X-ray modalities (DX, CR, RF, XA, MG, DR, PX, IO, OP) and skips SR/KO/PR/SC.
- **Formatted Excel output** - Styled headers, auto-filter, freeze panes, and proper column widths.

## Requirements

- Python 3.8+
- Dependencies:
  ```
  pip install pydicom pandas openpyxl
  ```

## Usage

**Basic** (outputs `Archivio_Esami_RX.xlsx`):
```bash
python3 dicom_to_excel.py /path/to/OsiriX_folder
```

**Custom output filename:**
```bash
python3 dicom_to_excel.py /path/to/OsiriX_folder --output MyArchive.xlsx
```

**Debug mode** (inspect raw DICOM tags without generating Excel):
```bash
python3 dicom_to_excel.py /path/to/OsiriX_folder --debug
python3 dicom_to_excel.py /path/to/OsiriX_folder --debug --debug-count 20
```

## Output Columns

| Column | Description |
|---|---|
| DATA ESAME | Exam date (DD/MM/YY) |
| TIPO ESAME | Exam type (e.g. `RX TORACE`, `RX MANO`) |
| 0 - 1 | Age bracket: 0 to 1 year |
| 1 - 16 | Age bracket: 1 to 16 years |
| 16 - 60 | Age bracket: 16 to 60 years |
| > 60 | Age bracket: over 60 years |
| DOSE | Dose Area Product in Gy-m^2 (scientific notation) |

## Year Selection

When the scanned folder contains studies from multiple years, the script prompts you to choose which year(s) to export:

```
Years found in data:
  [1] 2023  (142 studies)
  [2] 2024  (208 studies)

Enter year(s) to export (e.g. 2024 or 2023,2024), or press Enter for all:
```

## Exam-Type Resolution Order

The script checks these DICOM tags in order and uses the first non-empty value:

1. `AcquisitionDeviceProcessingDescription (0018,1400)`
2. `SeriesDescription (0008,103E)`
3. `StudyDescription (0008,1030)`
4. `RequestedProcedureDescription (0032,1060)`
5. `PerformedProcedureStepDescription (0040,0254)`
6. `ProtocolName (0018,1030)`
7. `BodyPartExamined (0018,0015)`
8. Parent folder name (OsiriX fallback)

## WARNING

### 1. Not Medically Approved
This program is **not approved** by any medical body, regulatory agency, or health authority. It is **not** a certified medical program and must **not** be used for clinical diagnosis, treatment, or any medical decision-making. Use it for informational, educational, or research purposes only.

### 2. No Liability
Bigm assumes **no responsibility** and **no liability** for any use, misuse, damages, errors, or consequences arising from the use of this program. This software is provided **"as is"**, without warranty of any kind, express or implied. Use it at your own risk.

### 3. Credit Required for Reposting
If you repost, share, redistribute, or build upon this program, you **must give proper credit** to Bigm. Reposting **without credit is not allowed**.
