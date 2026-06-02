import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import shutil, glob, os

INPUT_DIR = "input/"
OUTPUT_DIR = "output/"
SHEET_NAME = "Resum RA"

os.makedirs(OUTPUT_DIR, exist_ok=True)
FILES = glob.glob(os.path.join(INPUT_DIR, "*.xlsx"))

def is_evaluated(val):
    """Avaluat = número o 'NA' (suspès). PDT i buit = no avaluat."""
    if pd.isna(val):
        return False
    if isinstance(val, str):
        return val.strip().upper() == "NA"
    return True

def is_approved(val):
    """Aprovat = nota >= 5. 'NA' = avaluat però suspès."""
    if pd.isna(val):
        return False
    if isinstance(val, str):
        return False
    try:
        return float(val) >= 5
    except (ValueError, TypeError):
        return False

def process_file(filepath):
    df = pd.read_excel(filepath, header=1, keep_default_na=False, na_values=[""])
    ra_cols = [c for c in df.columns if str(c).endswith("RA")]

    results = []
    for _, row in df.iterrows():
        student_id = row.iloc[0]
        student_name = row.iloc[1]
        evaluats = sum(is_evaluated(row[c]) for c in ra_cols)
        aprovats = sum(is_approved(row[c]) for c in ra_cols)
        results.append({
            "ID Alumne": student_id,
            "Nom": student_name,
            "RA Avaluats": evaluats,
            "RA Aprovats": aprovats,
        })

    filename = os.path.basename(filepath)
    outpath = os.path.join(OUTPUT_DIR, filename)
    shutil.copy(filepath, outpath)

    wb = load_workbook(outpath)
    if SHEET_NAME in wb.sheetnames:
        del wb[SHEET_NAME]
    ws = wb.create_sheet(SHEET_NAME)

    header_font = Font(name="Arial", bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", start_color="2E4057")
    center = Alignment(horizontal="center", vertical="center")
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    alt_fill = PatternFill("solid", start_color="EBF0F7")

    headers = ["ID Alumne", "Nom", "RA Avaluats", "RA Aprovats"]
    col_widths = [12, 30, 16, 16]

    for col_idx, (header, width) in enumerate(zip(headers, col_widths), start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.row_dimensions[1].height = 22

    for row_idx, record in enumerate(results, start=2):
        fill = alt_fill if row_idx % 2 == 0 else None
        for col_idx, key in enumerate(headers, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=record[key])
            cell.border = border
            cell.alignment = center if col_idx != 2 else Alignment(vertical="center")
            if fill:
                cell.fill = fill
            cell.font = Font(name="Arial")

    wb.save(outpath)
    print(f"Fitxer guardat: {outpath}")
    print(pd.DataFrame(results).to_string(index=False))
    print()

if not FILES:
    print(f"No s'han trobat fitxers .xlsx a '{INPUT_DIR}'")
else:
    for f in sorted(FILES):
        print(f"=== Processant: {os.path.basename(f)} ===")
        process_file(f)
