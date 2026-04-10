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
    python dicom_to_excel.py /path/to/OsiriX_folder --year 2024
    python dicom_to_excel.py /path/to/OsiriX_folder --year 2023,2024
    python dicom_to_excel.py /path/to/OsiriX_folder --debug   ← inspect tags

Dependencies:
    pip install pydicom pandas openpyxl
"""

import os, sys, re, argparse, time
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

# English → Italian body-part translation for scanner descriptions
_BODY_PART_IT = {
    # Multi-word phrases (matched first, longest wins)
    "CERVICAL SPINE":     "COLONNA CERVICALE",
    "THORACIC SPINE":     "COLONNA DORSALE",
    "LUMBAR SPINE":       "COLONNA LOMBARE",
    "LUMBOSACRAL SPINE":  "COLONNA LOMBOSACRALE",
    "WHOLE SPINE":        "COLONNA IN TOTO",
    "FULL SPINE":         "COLONNA IN TOTO",
    "C SPINE":            "COLONNA CERVICALE",
    "T SPINE":            "COLONNA DORSALE",
    "L SPINE":            "COLONNA LOMBARE",
    "FACIAL BONES":       "OSSA FACCIALI",
    "NASAL BONES":        "OSSA NASALI",
    # Single words
    "FOOT":       "PIEDE",
    "FEET":       "PIEDI",
    "HAND":       "MANO",
    "HANDS":      "MANI",
    "CHEST":      "TORACE",
    "RIB":        "COSTOLA",
    "RIBS":       "COSTOLE",
    "SKULL":      "CRANIO",
    "HEAD":       "CRANIO",
    "KNEE":       "GINOCCHIO",
    "HIP":        "ANCA",
    "SHOULDER":   "SPALLA",
    "ELBOW":      "GOMITO",
    "WRIST":      "POLSO",
    "ANKLE":      "CAVIGLIA",
    "PELVIS":     "BACINO",
    "ABDOMEN":    "ADDOME",
    "FINGER":     "DITO",
    "FINGERS":    "DITA",
    "THUMB":      "POLLICE",
    "TOE":        "DITO PIEDE",
    "TOES":       "DITA PIEDE",
    "FEMUR":      "FEMORE",
    "TIBIA":      "TIBIA",
    "FIBULA":     "PERONE",
    "HUMERUS":    "OMERO",
    "RADIUS":     "RADIO",
    "ULNA":       "ULNA",
    "CLAVICLE":   "CLAVICOLA",
    "SCAPULA":    "SCAPOLA",
    "SPINE":      "COLONNA",
    "FOREARM":    "AVAMBRACCIO",
    "LEG":        "GAMBA",
    "ARM":        "BRACCIO",
    "NECK":       "COLLO",
    "CALCANEUS":  "CALCAGNO",
    "PATELLA":    "ROTULA",
    "STERNUM":    "STERNO",
    "MANDIBLE":   "MANDIBOLA",
    "JAW":        "MANDIBOLA",
    "SACRUM":     "SACRO",
    "COCCYX":     "COCCIGE",
    "THIGH":      "COSCIA",
    "ORBIT":      "ORBITA",
}

# Projection abbreviations (kept as-is, not translated)
_PROJECTIONS = {"AP", "PA", "LAT", "LL", "RL", "OBL", "AX", "AXIAL",
                "OBLIQUE", "LATERAL", "TANGENTIAL", "SKYLINE"}

# Scanner processing / positioning words to strip from exam descriptions
_STRIP_WORDS = {"ORTO", "ORTHO", "CLINO", "ERECT", "SUPINE", "PRONE",
                "GRID", "TABLE", "WALL", "BUCKY", "UPRIGHT",
                "STANDING", "RECUMBENT", "DECUBITUS"}


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


def _clean_description(raw: str) -> str:
    """
    Clean scanner descriptions like 'FOOT/FOOT AP/STYLE_M':
      1. Split on '/' and discard STYLE_* / processing segments
      2. Pick the most descriptive segment (prefers body-part + projection)
      3. Translate English body parts to Italian
    Already-Italian values pass through unchanged.
    """
    segments = [s.strip() for s in raw.split("/") if s.strip()]

    # Drop processing-style segments
    segments = [s for s in segments
                if not re.match(r"^(STYLE|PROC|MENU|MODE|FILTER)[_\s\-]", s)]

    if not segments:
        return raw

    # Prefer the segment that includes a projection and a body part
    best = segments[0]
    for s in segments:
        words = s.split()
        if len(words) > 1 and any(w in _PROJECTIONS for w in words):
            best = s
            break

    # Translate English body parts → Italian (longest phrases first)
    result = best
    for en, it in sorted(_BODY_PART_IT.items(), key=lambda x: -len(x[0])):
        result = re.sub(r"\b" + re.escape(en) + r"\b", it, result)

    # Strip scanner positioning / processing junk words
    result = " ".join(w for w in result.split() if w not in _STRIP_WORDS)

    # Normalise projection: LAT → LL
    result = re.sub(r"\bLAT\b", "LL", result)

    return result


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
    # Skip all Structured Report / annotation objects (UIDs under ...1.1.88.*)
    sop = _tag(ds, (0x0008, 0x0016))
    if ".1.1.88." in sop:
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
    # Clean scanner descriptions and translate body parts to Italian
    raw = _clean_description(raw)
    # Prepend 'RX ' if missing
    if not re.match(r"^RX\b", raw):
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
    if   age < 1:   cols["0 - 1"]   = "X"
    elif age <= 16: cols["1 - 16"]  = "X"
    elif age <= 60: cols["16 - 60"] = "X"
    else:           cols["> 60"]    = "X"
    return cols


def format_date(raw: str) -> str:
    raw = str(raw).strip()
    if len(raw) == 8 and raw.isdigit():
        return f"{raw[6:8]}/{raw[4:6]}/{raw[0:4]}"
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
    t0 = time.monotonic()

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

    elapsed = time.monotonic() - t0
    print(f"\n  Scan done: {scanned} files | {len(study_files)} studies "
          f"| {skipped} non-X-ray skipped | {elapsed:.1f}s\n", flush=True)

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


def _add_summary_sheet(wb: Workbook, df: pd.DataFrame) -> None:
    """Add a 'Riepilogo Mensile' sheet with exam-type counts by month."""
    if "_sort_date" not in df.columns or "TIPO ESAME" not in df.columns:
        return

    month_label = df["_sort_date"].apply(
        lambda x: f"{x[4:6]}/{x[0:4]}" if len(x) >= 6 else "??/????"
    )
    mask = df["TIPO ESAME"] != ""
    if not mask.any():
        return

    ct = pd.crosstab(
        df.loc[mask, "TIPO ESAME"],
        month_label[mask],
        margins=True,
        margins_name="TOTALE",
    )

    # Sort month columns chronologically, TOTALE last
    month_cols = sorted(
        [c for c in ct.columns if c != "TOTALE"],
        key=lambda x: x[3:7] + x[0:2],
    )
    col_order = month_cols + ["TOTALE"]
    ct = ct[col_order]

    # Sort rows by total descending, TOTALE row last
    if "TOTALE" in ct.index:
        total_row = ct.loc[["TOTALE"]]
        ct_body = ct.drop("TOTALE").sort_values("TOTALE", ascending=False)
        ct = pd.concat([ct_body, total_row])

    ws = wb.create_sheet("Riepilogo Mensile")
    BLUE, GREY, LIGHT_BLUE = "BDD7EE", "D9D9D9", "DAEEF3"
    n_cols = 1 + len(col_order)

    # Column widths
    ws.column_dimensions["A"].width = max(22, max(len(str(i)) for i in ct.index) + 4)
    for ci, col_name in enumerate(col_order, start=2):
        ws.column_dimensions[get_column_letter(ci)].width = max(9, len(col_name) + 2)

    # Row 1 – title
    _hdr(ws, 1, 1, "RIEPILOGO MENSILE ESAMI", bg=BLUE, size=11)
    if n_cols > 1:
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    ws.row_dimensions[1].height = 22

    # Row 2 – column headers
    _hdr(ws, 2, 1, "TIPO ESAME", bg=GREY, size=9)
    for ci, col_name in enumerate(col_order, start=2):
        _hdr(ws, 2, ci, col_name, bg=GREY, size=9)
    ws.row_dimensions[2].height = 18

    # Data rows
    for ri, idx in enumerate(ct.index):
        xl_row = ri + 3
        is_total = (idx == "TOTALE")
        bg = LIGHT_BLUE if is_total else None
        bold = is_total

        cell = ws.cell(row=xl_row, column=1, value=str(idx))
        cell.font      = Font(name="Arial", size=9, bold=bold)
        cell.alignment = Alignment(horizontal="left", vertical="center")
        cell.border    = _ALL
        if bg:
            cell.fill = PatternFill("solid", start_color=bg)

        for ci, col_name in enumerate(col_order, start=2):
            val = int(ct.loc[idx, col_name])
            cell = ws.cell(row=xl_row, column=ci, value=val if val else None)
            cell.font      = Font(name="Arial", size=9, bold=bold)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border    = _ALL
            if bg:
                cell.fill = PatternFill("solid", start_color=bg)
            # TOTALE row always shows values, data rows blank out zeroes
            if is_total and val == 0:
                cell.value = 0

        ws.row_dimensions[xl_row].height = 15

    ws.freeze_panes = "B3"


def build_excel(df: pd.DataFrame, output_path: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Archivio Esami RX"

    C_DATE, C_TYPE           = 1, 2
    C_A1, C_A2, C_A3, C_A4  = 3, 4, 5, 6
    C_DOSE                   = 7

    for col, w in {C_DATE: 14, C_TYPE: 28,
                   C_A1: 7, C_A2: 7, C_A3: 8, C_A4: 7, C_DOSE: 40}.items():
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

    _add_summary_sheet(wb, df)

    try:
        wb.save(output_path)
    except PermissionError:
        print(f"❌  Cannot write '{output_path}' – file is open in another program.",
              file=sys.stderr)
        base, ext = os.path.splitext(output_path)
        alt_path = f"{base}_{int(time.time())}{ext}"
        wb.save(alt_path)
        print(f"✅  Saved to alternate path → {alt_path}  ({len(df)} rows)")
        return

    print(f"✅  Saved → {output_path}  ({len(df)} rows)")


# ─────────────────────────────────────────────────────────────────
# YEAR FILTER
# ─────────────────────────────────────────────────────────────────

def prompt_year_filter(df: pd.DataFrame, cli_years: Optional[str] = None) -> pd.DataFrame:
    """
    Filter by year.  If *cli_years* is provided (from --year), use it
    directly without prompting.  Otherwise prompt interactively.
    """
    years = sorted(df["_sort_date"].str[:4].unique().tolist())

    # Non-interactive: --year flag
    if cli_years:
        tokens = [t.strip() for t in cli_years.split(",") if t.strip()]
        invalid = [t for t in tokens if t not in years]
        if invalid:
            print(f"⚠️  Year(s) not in data: {', '.join(invalid)}. "
                  f"Available: {', '.join(years)}", file=sys.stderr)
            tokens = [t for t in tokens if t in years]
        if not tokens:
            print(f"  → No matching years. Exporting all {len(df)} studies.")
            return df
        filtered = df[df["_sort_date"].str[:4].isin(tokens)].reset_index(drop=True)
        print(f"  → {len(filtered)} studies for year(s): {', '.join(tokens)}.")
        return filtered

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
# SUMMARY STATISTICS
# ─────────────────────────────────────────────────────────────────

def print_summary(df: pd.DataFrame) -> None:
    """Print a brief summary of the exported data."""
    print(f"\n{'─'*50}")
    print(f"  Summary: {len(df)} studies exported")
    print(f"{'─'*50}")

    has_dates = "_sort_date" in df.columns
    has_tipos = "TIPO ESAME" in df.columns

    # ── Monthly breakdown by exam type ───────────────────────────
    if has_dates and has_tipos:
        # Extract MM/YYYY from _sort_date (YYYYMMDD)
        month_label = df["_sort_date"].apply(
            lambda x: f"{x[4:6]}/{x[0:4]}" if len(x) >= 6 else "??/????"
        )
        mask = df["TIPO ESAME"] != ""
        if mask.any():
            ct = pd.crosstab(
                df.loc[mask, "TIPO ESAME"],
                month_label[mask],
                margins=True,
                margins_name="TOTALE",
            )
            # Sort month columns chronologically, TOTALE last
            month_cols = sorted(
                [c for c in ct.columns if c != "TOTALE"],
                key=lambda x: x[3:7] + x[0:2],
            )
            ct = ct[month_cols + ["TOTALE"]]

            # Sort rows by total descending, TOTALE row last
            if "TOTALE" in ct.index:
                total_row = ct.loc[["TOTALE"]]
                ct_body = ct.drop("TOTALE").sort_values("TOTALE", ascending=False)
                ct = pd.concat([ct_body, total_row])

            # Column widths
            tipo_w = max(20, max(len(str(i)) for i in ct.index) + 2)
            col_w = max(7, max(len(c) for c in ct.columns) + 1)
            sep = "    " + "─" * (tipo_w + (col_w + 1) * len(ct.columns))

            print(f"\n  Exam count by month:\n")

            # Header
            hdr = f"    {'TIPO ESAME':<{tipo_w}}"
            for c in ct.columns:
                hdr += f" {c:>{col_w}}"
            print(hdr)
            print(sep)

            # Data rows
            for idx in ct.index:
                if idx == "TOTALE":
                    print(sep)
                line = f"    {str(idx):<{tipo_w}}"
                for c in ct.columns:
                    line += f" {int(ct.loc[idx, c]):>{col_w}}"
                print(line)

            print()

    elif has_tipos:
        # Fallback when no date column is available
        tipo_counts = df["TIPO ESAME"].replace("", pd.NA).dropna().value_counts()
        if not tipo_counts.empty:
            print("\n  Top exam types:")
            for tipo, count in tipo_counts.head(10).items():
                print(f"    {tipo:30s}  {count:>4d}")

    # ── By age category ──────────────────────────────────────────
    age_cols = ["0 - 1", "1 - 16", "16 - 60", "> 60"]
    present = [c for c in age_cols if c in df.columns]
    if present:
        total_marked = 0
        print("\n  By age group:")
        for col in present:
            count = (df[col] == "X").sum()
            total_marked += count
            print(f"    {col:30s}  {count:>4d}")
        unknown = len(df) - total_marked
        if unknown > 0:
            print(f"    {'Unknown':30s}  {unknown:>4d}")

    print()


# ─────────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser(
        description="Extract DICOM metadata from OsiriX folder → Archivio Esami RX Excel"
    )
    ap.add_argument("folder", help="Root path of the OsiriX MD study folder")
    ap.add_argument("--output", "-o", default="Archivio_Esami_RX.xlsx")
    ap.add_argument("--year", "-y", default=None, metavar="YYYY",
                    help="Export only these year(s), comma-separated (e.g. 2024 or 2023,2024)")
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

    df = prompt_year_filter(df, cli_years=args.year)

    if df.empty:
        print("⚠️  No studies remaining after year filter.")
        sys.exit(0)

    print("📊  Building Excel file …")
    build_excel(df, args.output)
    print_summary(df)


if __name__ == "__main__":
    main()
