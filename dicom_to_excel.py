"""
DICOM Metadata Extractor → Archivio Esami RX
=============================================
One Excel row = one DICOM Study (one patient exam order).
Multiple series / projections within the same study are merged:
  - TIPO ESAME comes from the first non-empty description tag found
  - DOSE collects every distinct value across all images, joined with ' / '

Fixes applied after inspecting real DICOM files:
  1. Group by StudyInstanceUID (not SeriesInstanceUID) → no more blank half
  2. Use AcquisitionDeviceProcessingDescription (0018,1400) as the primary
     TIPO ESAME source; prepend 'RX ' if the value doesn't already start with it
  3. Divide dose (0018,115E) by 100 000 to convert dGy·cm² → Gy·m²
  4. Skip SR / KO / PR / SC modalities entirely

Usage:
    python dicom_to_excel.py /path/to/OsiriX_folder
    python dicom_to_excel.py /path/to/OsiriX_folder --output MyArchive.xlsx
    python dicom_to_excel.py /path/to/OsiriX_folder --debug   ← inspect tags

Dependencies:
    pip install pydicom pandas openpyxl
"""

import os, sys, re, argparse, struct
from pathlib import Path
from collections import defaultdict
from typing import Optional, Dict, List

import pydicom
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ─────────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────────

# Only real X-ray image modalities produce data we want
XRAY_MODALITIES = {"DX", "CR", "RF", "XA", "MG", "DR", "PX", "IO", "OP", ""}

# dGy·cm² → Gy·m²  (DICOM tag 0018,115E is stored in dGy·cm²)
# 1 dGy·cm² = 0.1 Gy × (0.01 m)² = 1e-5 Gy·m²
DOSE_CONVERSION = 1e-5


# ─────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────

def _tag(ds, *tags, default="") -> str:
    """Try each tag in order; return first non-empty string value."""
    for t in tags:
        try:
            v = str(ds[t].value).strip()
            if v and v.upper() not in ("", "NONE"):
                return v
        except (KeyError, AttributeError, TypeError):
            pass
    return default


def _build_tipo(ds, fallback_path: Path) -> str:
    """
    Return the best exam-type label for this dataset.

    Priority order (reflects what this scanner actually populates):
      1. AcquisitionDeviceProcessingDescription (0018,1400)  ← KEY for this scanner
      2. SeriesDescription         (0008,103E)
      3. StudyDescription          (0008,1030)
      4. RequestedProcedureDescription  (0032,1060)
      5. PerformedProcedureStepDescription (0040,0254)
      6. ProtocolName              (0018,1030)
      7. BodyPartExamined          (0018,0015)               ← last resort
      8. Parent folder name        (OsiriX sometimes stores it there)

    After obtaining the raw value, the function:
      - Strips leading numbers/underscores OsiriX sometimes prefixes
      - Prepends 'RX ' if the value doesn't already start with 'RX'
      - Returns '' if nothing found
    """
    # Skip SR / annotation objects – they have their own description
    sop = _tag(ds, (0x0008, 0x0016))
    if "88.11.1" in sop or "88.67" in sop or "88.59" in sop:
        return ""

    raw = _tag(ds,
               (0x0018, 0x1400),   # AcquisitionDeviceProcessingDescription ← main one for this scanner
               (0x0008, 0x103E),   # SeriesDescription
               (0x0008, 0x1030),   # StudyDescription
               (0x0032, 0x1060),   # RequestedProcedureDescription
               (0x0040, 0x0254),   # PerformedProcedureStepDescription
               (0x0018, 0x1030),   # ProtocolName
               (0x0018, 0x0015),   # BodyPartExamined
               )

    # Folder-name fallback (OsiriX sometimes puts exam name in folder)
    if not raw:
        for part in reversed(fallback_path.parts):
            p = part.strip()
            if re.match(r"^(RX|ECO|TC|RM|MX)\b", p, re.IGNORECASE):
                raw = re.sub(r"^\d+[-_\s]*", "", p).strip()
                break

    if not raw:
        return ""

    # Normalise
    raw = raw.strip()
    # Remove OsiriX numeric prefixes like "0012_MANO"
    raw = re.sub(r"^\d+[-_\s]+", "", raw).strip()
    # Upper-case for consistency
    raw = raw.upper()
    # Prepend 'RX ' if missing
    if not raw.startswith("RX"):
        raw = "RX " + raw

    return raw


def parse_age(age_str: str) -> Optional[int]:
    if not age_str:
        return None
    m = re.match(r"(\d+)([YMWDymwd])", age_str.strip())
    if not m:
        return None
    value, unit = int(m.group(1)), m.group(2).upper()
    if unit == "Y":  return value
    if unit == "M":  return max(0, value // 12)
    return 0


def age_to_category(age: Optional[int]) -> Dict[str, str]:
    cols = {"0 - 1": "", "1 - 16": "", "16 - 60": "", "> 60": ""}
    if age is None:
        return cols
    if   age <= 1:  cols["0 - 1"]   = "X"
    elif age <= 16: cols["1 - 16"]  = "X"
    elif age <= 60: cols["16 - 60"] = "X"
    else:           cols["> 60"]    = "X"
    return cols


def format_date(raw: str) -> str:
    raw = str(raw).strip()
    if len(raw) == 8 and raw.isdigit():
        return f"{raw[6:8]}/{raw[4:6]}/{raw[2:4]}"
    return raw


def format_dose(raw_str: str) -> str:
    """
    Convert raw DICOM dose to 'X.XXe-XX Gy.m2'.

    DICOM tag (0018,115E) should be in dGy·cm², but some vendors (e.g. Fujifilm
    CALNEO) write it already in Gy·m².  Distinguish by magnitude:
      - value >= 0.1  -> stored in dGy·cm²  -> multiply by 1e-5 to get Gy·m²
      - value <  0.1  -> already in Gy·m²   -> use as-is
    """
    try:
        value = float(raw_str)
        if value >= 0.1:        # stored in dGy·cm², convert
            value *= DOSE_CONVERSION
        return f"{value:.2e} Gy.m2"
    except (ValueError, TypeError):
        return ""


# ─────────────────────────────────────────────────────────────────
# DEBUG MODE
# ─────────────────────────────────────────────────────────────────

def run_debug(root: str, max_files: int = 10):
    print(f"\n{'='*65}")
    print(f"DEBUG – first {max_files} valid DICOM files in: {root}")
    print(f"{'='*65}\n")
    TAGS = [
        ((0x0008, 0x0016), "SOPClassUID"),
        ((0x0008, 0x0060), "Modality"),
        ((0x0008, 0x0020), "StudyDate"),
        ((0x0008, 0x1030), "StudyDescription"),
        ((0x0008, 0x103E), "SeriesDescription"),
        ((0x0032, 0x1060), "RequestedProcedureDescription"),
        ((0x0040, 0x0254), "PerformedProcedureStepDescription"),
        ((0x0018, 0x1030), "ProtocolName"),
        ((0x0018, 0x1400), "AcquisitionDeviceProcessingDescription"),
        ((0x0018, 0x0015), "BodyPartExamined"),
        ((0x0010, 0x1010), "PatientAge"),
        ((0x0018, 0x115E), "ImageAreaDoseProduct (raw dGy·cm²)"),
        ((0x0020, 0x000D), "StudyInstanceUID"),
        ((0x0020, 0x000E), "SeriesInstanceUID"),
    ]
    shown = 0
    for dirpath, _, files in os.walk(root):
        for fname in files:
            fpath = Path(dirpath) / fname
            try:
                ds = pydicom.dcmread(str(fpath), stop_before_pixels=True)
            except Exception:
                continue
            shown += 1
            print(f"File #{shown}: {fpath.name}  (folder: {fpath.parent.name})")
            for tag, name in TAGS:
                try:
                    v = ds[tag].value
                    extra = ""
                    if tag == (0x0018, 0x115E):
                        try:
                            extra = f"  →  {float(v)*DOSE_CONVERSION:.2e} Gy.m2"
                        except Exception:
                            pass
                    print(f"  {name:48s} = {v!r}{extra}")
                except KeyError:
                    print(f"  {name:48s} = <missing>")
            print()
            if shown >= max_files:
                return


# ─────────────────────────────────────────────────────────────────
# SCAN
# ─────────────────────────────────────────────────────────────────

def scan_folder(root: str) -> pd.DataFrame:
    """
    Two-pass scan.
    Pass 1 – group files by StudyInstanceUID, skip non-X-ray modalities.
             Also caches tipo and raw dose per study so Pass 2 needs no
             additional file reads.
    Pass 2 – per study: assemble the record from cached values.
    """

    # study_files[uid] = list of (date_raw, fpath)
    study_files:    Dict[str, List]            = defaultdict(list)
    study_meta:     Dict[str, pydicom.Dataset] = {}
    study_tipo:     Dict[str, str]             = {}   # first non-empty tipo per study
    study_dose_raw: Dict[str, str]             = {}   # first non-empty raw dose per study
    scanned = skipped = 0

    for dirpath, _, files in os.walk(root):
        for fname in files:
            fpath = Path(dirpath) / fname
            try:
                ds = pydicom.dcmread(str(fpath), stop_before_pixels=True)
            except Exception:
                scanned += 1
                continue

            modality = _tag(ds, (0x0008, 0x0060)).upper()
            if modality and modality not in XRAY_MODALITIES:
                skipped += 1
                scanned += 1
                continue

            date_raw  = _tag(ds, (0x0008, 0x0020))
            study_uid = _tag(ds, (0x0020, 0x000D))

            # Fallback key if StudyInstanceUID is absent
            if not study_uid:
                study_uid = f"{date_raw}|{_tag(ds,(0x0008,0x0070))}|{fname}"

            study_files[study_uid].append((date_raw, fpath))
            if study_uid not in study_meta:
                study_meta[study_uid] = ds

            # Cache tipo (first non-empty result wins across all files)
            if study_uid not in study_tipo:
                t = _build_tipo(ds, fpath)
                if t:
                    study_tipo[study_uid] = t

            # Cache raw dose string (first valid reading wins)
            if study_uid not in study_dose_raw:
                raw_d = _tag(ds, (0x0018, 0x115E))
                if not raw_d:
                    raw_d = _tag(ds, (0x0040, 0x8302))
                if not raw_d:
                    raw_d = _tag(ds, (0x0040, 0x0302))
                if raw_d:
                    study_dose_raw[study_uid] = raw_d

            scanned += 1
            if scanned % 200 == 0:
                print(f"  … {scanned} files, {len(study_files)} studies, "
                      f"{skipped} non-X-ray skipped", flush=True)

    print(f"\n  Scan done: {scanned} files | {len(study_files)} studies "
          f"| {skipped} non-X-ray skipped\n", flush=True)

    records: List[Dict] = []
    skipped_studies = 0

    for study_uid, file_list in study_files.items():
        if not file_list:
            continue
        file_list.sort(key=lambda t: (t[0], str(t[1])))

        # ── Metadata from the best available file ───────────────────
        meta_ds = study_meta.get(study_uid)
        if meta_ds is None:
            for _, fp in file_list:
                try:
                    meta_ds = pydicom.dcmread(str(fp), stop_before_pixels=True)
                    break
                except Exception:
                    continue
        if meta_ds is None:
            continue

        date_raw = _tag(meta_ds, (0x0008, 0x0020))

        # ── TIPO ESAME (from Pass 1 cache) ───────────────────────────
        tipo = study_tipo.get(study_uid, "")

        # ── Patient age ──────────────────────────────────────────────
        age_raw  = _tag(meta_ds, (0x0010, 0x1010))
        age_cols = age_to_category(parse_age(age_raw))

        # ── DOSE (from Pass 1 cache) ──────────────────────────────────
        raw_dose = study_dose_raw.get(study_uid, "")
        dose_str = format_dose(raw_dose) if raw_dose else ""

        # Skip studies with no identifiable date
        if not date_raw:
            skipped_studies += 1
            continue

        records.append({
            "_sort_date": date_raw,
            "DATA ESAME": format_date(date_raw) if date_raw else "",
            "TIPO ESAME": tipo,
            **age_cols,
            "DOSE": dose_str,
        })

    print(f"  Studies skipped (no date): {skipped_studies}", flush=True)

    if not records:
        return pd.DataFrame()

    df = pd.DataFrame(records)
    df.sort_values("_sort_date", inplace=True, ignore_index=True)
    df.fillna("", inplace=True)
    return df


# ─────────────────────────────────────────────────────────────────
# EXCEL BUILDER
# ─────────────────────────────────────────────────────────────────

_THIN = Side(style="thin")

def _border(*sides) -> Border:
    return Border(**{s: _THIN for s in sides})

_ALL = _border("top", "left", "bottom", "right")

def _hdr(ws, row, col, value, *, bg, bold=True, size=10):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font      = Font(name="Arial", bold=bold, size=size, color="000000")
    cell.fill      = PatternFill("solid", start_color=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border    = _ALL
    return cell

def _clean(v) -> Optional[str]:
    if v is None:
        return None
    if isinstance(v, float) and pd.isna(v):
        return None
    s = str(v)
    return s if s.strip() else None


def build_excel(df: pd.DataFrame, output_path: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Archivio Esami RX"

    C_DATE, C_TYPE           = 1, 2
    C_A1, C_A2, C_A3, C_A4  = 3, 4, 5, 6
    C_DOSE                   = 7

    for col, w in {C_DATE: 13, C_TYPE: 28, C_A1: 7, C_A2: 7,
                   C_A3: 8, C_A4: 7, C_DOSE: 40}.items():
        ws.column_dimensions[get_column_letter(col)].width = w

    BLUE, GREY = "BDD7EE", "D9D9D9"

    # Row 1 – group headers
    _hdr(ws, 1, C_DATE, "ARCHIVIO ESAMI RX", bg=BLUE, size=11)
    ws.merge_cells(start_row=1, start_column=C_DATE, end_row=1, end_column=C_TYPE)
    _hdr(ws, 1, C_A1, "ETA'", bg=BLUE, size=11)
    ws.merge_cells(start_row=1, start_column=C_A1, end_row=1, end_column=C_A4)
    _hdr(ws, 1, C_DOSE, "DOSE DAP", bg=BLUE, size=11)

    # Row 2 – sub-headers
    for col, label in [(C_DATE, "DATA ESAME"), (C_TYPE, "TIPO ESAME"),
                       (C_A1, "0 - 1"), (C_A2, "1 - 16"),
                       (C_A3, "16 - 60"), (C_A4, "> 60"),
                       (C_DOSE, "Gy·m²")]:
        _hdr(ws, 2, col, label, bg=GREY, size=9)

    ws.row_dimensions[1].height = 20
    ws.row_dimensions[2].height = 18

    DATA_COLS = ["DATA ESAME", "TIPO ESAME",
                 "0 - 1", "1 - 16", "16 - 60", "> 60", "DOSE"]

    for r_idx, row in df.iterrows():
        xl_row = r_idx + 3
        for c_idx, col_name in enumerate(DATA_COLS, start=1):
            val  = _clean(row.get(col_name, ""))
            cell = ws.cell(row=xl_row, column=c_idx, value=val)
            cell.font      = Font(name="Arial", size=9)
            cell.alignment = Alignment(
                horizontal="center" if c_idx >= C_A1 else "left",
                vertical="center",
                wrap_text=(c_idx == C_DOSE),
            )
            cell.border = _border("left", "right", "top", "bottom")
        ws.row_dimensions[xl_row].height = 15

    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{get_column_letter(C_DOSE)}{len(df) + 2}"
    wb.save(output_path)
    print(f"✅  Saved → {output_path}  ({len(df)} rows)")


# ─────────────────────────────────────────────────────────────────
# YEAR FILTER
# ─────────────────────────────────────────────────────────────────

def prompt_year_filter(df: pd.DataFrame) -> pd.DataFrame:
    """
    Inspect the _sort_date column to find available years, then prompt
    the user to select which year(s) to export.  Returns the filtered
    DataFrame (still containing _sort_date; caller is responsible for
    dropping it).  If only one year is present the prompt is skipped.
    """
    years = sorted(df["_sort_date"].str[:4].unique().tolist())

    if len(years) <= 1:
        return df

    print("\n📅  Years found in data:")
    for i, yr in enumerate(years, 1):
        count = (df["_sort_date"].str[:4] == yr).sum()
        print(f"  [{i}] {yr}  ({count} studies)")

    while True:
        raw = input(
            "\nEnter year(s) to export "
            f"(e.g. {years[-1]} or {','.join(years[:2])}"
            "), or press Enter for all: "
        ).strip()

        if not raw:
            print(f"  → Exporting all {len(df)} studies.")
            return df

        tokens = [t.strip() for t in raw.split(",") if t.strip()]
        invalid = [t for t in tokens if t not in years]
        if invalid:
            print(f"  ⚠️  Not found: {', '.join(invalid)}. "
                  f"Available years: {', '.join(years)}")
            continue

        filtered = df[df["_sort_date"].str[:4].isin(tokens)].reset_index(drop=True)
        print(f"  → {len(filtered)} studies selected for year(s): {', '.join(tokens)}.")
        return filtered


# ─────────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser(
        description="Extract DICOM metadata from OsiriX folder → Archivio Esami RX Excel"
    )
    ap.add_argument("folder", help="Root path of the OsiriX MD study folder")
    ap.add_argument("--output", "-o", default="Archivio_Esami_RX.xlsx")
    ap.add_argument("--debug", action="store_true",
                    help="Print tag values of first N files and exit")
    ap.add_argument("--debug-count", type=int, default=10, metavar="N")
    args = ap.parse_args()

    if not os.path.isdir(args.folder):
        print(f"❌  '{args.folder}' is not a valid directory.", file=sys.stderr)
        sys.exit(1)

    if args.debug:
        run_debug(args.folder, args.debug_count)
        return

    print(f"🔍  Scanning: {args.folder}")
    df = scan_folder(args.folder)

    if df.empty:
        print("⚠️  No valid DICOM studies found.")
        print("    Tip: run with --debug to inspect tag contents.")
        sys.exit(0)

    df = prompt_year_filter(df)
    df.drop(columns=["_sort_date"], inplace=True)

    if df.empty:
        print("⚠️  No studies remaining after year filter.")
        sys.exit(0)

    print("📊  Building Excel file …")
    build_excel(df, args.output)


if __name__ == "__main__":
    main()
