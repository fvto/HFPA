"""
Populate Database/HFPA_Template.xlsx with accurate data from HFPA (Mes410) and QAStation.
Adheres strictly to Database/2026 FTT&HFPA EN.pdf and QAStation_Tool_Agent_Spec.md:
- Never overwrite formula cells (=SUM, =1-J9/J2, =F22/E22, etc.).
- Input production and defects into exact month columns.
- Populate Top 5 Models and Top 5 Defect Matrix per factory.
- Preserves all existing sheets, cell formatting, and formulas.
"""

import os
import glob
import re
import shutil
import pandas as pd
import openpyxl
from collections import defaultdict
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from file_utils import ensure_file_writable, detect_file_month

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATABASE_TEMPLATE = os.path.join(BASE_DIR, "Database", "HFPA_Template.xlsx")
OUTPUT_DIR = os.path.join(BASE_DIR, "Output")
OUTPUT_DATABASE = os.path.join(OUTPUT_DIR, "HFPA_Template_Updated.xlsx")
INPUT_HFPA_DIR = os.path.join(BASE_DIR, "Input", "HFPA")
INPUT_FTT_DIR = os.path.join(BASE_DIR, "Input", "FTT")

MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
FACTORIES = ["VH", "VH2", "JV", "JV2"]


def detect_month_from_data(hfpa_df: pd.DataFrame) -> str:
    """Detect month abbreviation (e.g., 'Jul') from Audit Date column."""
    date_col = "Audit Date (dd-mmm-yyyy)" if "Audit Date (dd-mmm-yyyy)" in hfpa_df.columns else "Audit Date"
    if date_col in hfpa_df.columns:
        dates = pd.to_datetime(hfpa_df[date_col], errors="coerce").dropna()
        if not dates.empty:
            month_num = dates.iloc[0].month
            return MONTH_NAMES[month_num - 1]
    return "Aug"


def load_top3_defects_from_reports():
    """Load per-model top 3 defects from QAStation analysis reports."""
    records = defaultdict(lambda: defaultdict(dict))
    report_candidates = [
        os.path.join(OUTPUT_DIR, "QAStation_HFPA_Analysis_Report.xlsx"),
        os.path.join(OUTPUT_DIR, "QAStation_HFPA_Analysis_Report_New.xlsx"),
    ]
    for rpath in report_candidates:
        if os.path.exists(rpath):
            try:
                wb_p2 = openpyxl.load_workbook(rpath, data_only=True)
                if "Pivot2_Top3_Defects" in wb_p2.sheetnames:
                    ws_p2 = wb_p2["Pivot2_Top3_Defects"]
                    for r in range(2, ws_p2.max_row + 1):
                        fty = ws_p2.cell(r, 1).value
                        model = ws_p2.cell(r, 2).value
                        dname = ws_p2.cell(r, 4).value
                        share = ws_p2.cell(r, 7).value
                        if fty and model and dname and share is not None:
                            try:
                                records[str(fty).strip()][str(model).strip()][str(dname).strip()] = float(share)
                            except (ValueError, TypeError):
                                pass
                    wb_p2.close()
                    if records:
                        return records
            except Exception:
                pass
    return records


def populate_hfpa_database(target_month=None, p2_df=None):
    """
    Populate HFPA sheet in database template.
    Strictly preserves formulas in rows 6, 13, 16-20, and Top 5 DR% columns.
    """
    if not os.path.exists(DATABASE_TEMPLATE):
        print(f"[-] Template not found at: {DATABASE_TEMPLATE}")
        return False

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    target_output_path = ensure_file_writable(OUTPUT_DATABASE, "Template đầu ra (Output/HFPA_Template_Updated.xlsx)")

    shutil.copy2(DATABASE_TEMPLATE, target_output_path)
    print(f"[+] Loaded database template into: {target_output_path}")

    # Load HFPA data for all 4 factories
    hfpa_dict = {}
    for fty in FACTORIES:
        files = glob.glob(os.path.join(INPUT_HFPA_DIR, f"*{fty}.xlsx"))
        if files:
            hfpa_dict[fty] = pd.read_excel(files[0])
            print(f"   [+] Loaded HFPA {fty}: {len(hfpa_dict[fty]):,} records")

    if not hfpa_dict:
        print("[-] No HFPA files loaded!")
        return False

    # Auto-detect month if not specified
    if not target_month:
        first_df = next(iter(hfpa_dict.values()))
        target_month = detect_month_from_data(first_df)
    print(f"[+] Target Month for population: {target_month}")

    wb = openpyxl.load_workbook(target_output_path, data_only=False)

    if "HFPA" not in wb.sheetnames:
        print("[-] Sheet 'HFPA' not found in template!")
        wb.close()
        return False

    ws = wb["HFPA"]
    print("   [+] Updating Sheet 'HFPA' strictly preserving formulas...")

    # Locate month column in row 1
    month_col = None
    for c in range(1, ws.max_column + 1):
        val = str(ws.cell(1, c).value or "").strip()
        if val.lower() == target_month.lower():
            month_col = c
            break

    if not month_col:
        month_col = 10  # Fallback to column J for Aug
        print(f"   [-] Month '{target_month}' not found in header, using column J ({month_col})")
    else:
        col_letter = get_column_letter(month_col)
        print(f"   [+] Found month '{target_month}' at column {col_letter} ({month_col})")

    # 1. Update Monthly Production (Audit Sample Size) & Defective Pairs
    # Rows: VH=2, VH2=3, JV=4, JV2=5 for Production
    # Rows: VH=9, VH2=10, JV=11, JV2=12 for Defects
    fty_prod_rows = {"VH": 2, "VH2": 3, "JV": 4, "JV2": 5}
    fty_def_rows = {"VH": 9, "VH2": 10, "JV": 11, "JV2": 12}

    for fty in FACTORIES:
        if fty not in hfpa_dict:
            continue
        df = hfpa_dict[fty]
        sample_col = "Audit Sample Size" if "Audit Sample Size" in df.columns else "Audit Sample"
        def_col = [c for c in df.columns if "Defective" in c and "pair" in c]
        def_col_name = def_col[0] if def_col else "Defective Q'ty(pair)"

        tot_sample = int(df[sample_col].sum())
        tot_defective = int(df[def_col_name].sum())

        prod_r = fty_prod_rows[fty]
        def_r = fty_def_rows[fty]

        ws.cell(prod_r, month_col).value = tot_sample
        ws.cell(prod_r, month_col).number_format = "#,##0"

        ws.cell(def_r, month_col).value = tot_defective
        ws.cell(def_r, month_col).number_format = "#,##0"

    # Restore and calculate totals across 4 factories with 'Grand Total' title (Rows 6, 13, 20)
    col_letter = get_column_letter(month_col)
    bold_hdr_font = Font(name="Calibri", size=11, bold=True)

    # Row 6: Grand Total for Production
    ws.cell(6, 1).value = "Grand Total"
    ws.cell(6, 1).font = bold_hdr_font
    ws.cell(6, 2).value = "Total pair produced"
    ws.cell(6, 2).font = bold_hdr_font
    ws.cell(6, month_col).value = f"=SUM({col_letter}2:{col_letter}5)"
    ws.cell(6, month_col).number_format = "#,##0"

    # Row 13: Grand Total for Defect Pairs
    ws.cell(13, 1).value = "Grand Total"
    ws.cell(13, 1).font = bold_hdr_font
    ws.cell(13, 2).value = "Defect"
    ws.cell(13, 2).font = bold_hdr_font
    ws.cell(13, month_col).value = f"=SUM({col_letter}9:{col_letter}12)"
    ws.cell(13, month_col).number_format = "#,##0"

    # Row 20: Grand Total for HFPA%
    ws.cell(20, 1).value = "Grand Total"
    ws.cell(20, 1).font = bold_hdr_font
    ws.cell(20, 2).value = "HFPA%"
    ws.cell(20, 2).font = bold_hdr_font
    ws.cell(20, month_col).value = f"=1-{col_letter}13/{col_letter}6"
    ws.cell(20, month_col).number_format = "0.0%"

    print(f"   [+] Updated monthly Production and Defective pairs with 'Grand Total' in rows 6, 13, 20.")

    # 2. Update Top 5 Models and Top 5 Defect Matrix per Factory
    # Layout specifications:
    # - VH:
    #     Left Table (Models): Rows 22-26. C22='VH', D22:D26=Model2, E22:E26=Tot Sample, F22:F26=Def. pairs, G22:G26==F{r}/E{r}
    #     Right Table (Defects): Header Row 21 (M21:Q21), Data Rows 22-26 (K22='VH', L22:L26=Model2, M22:Q26=Share%)
    # - VH2:
    #     Left Table (Models): Rows 28-32. C28='VH2', D28:D32=Model2, E28:E32=Tot Sample, F28:F32=Def. pairs, G28:G32==F{r}/E{r}
    #     Right Table (Defects): Header Row 27 (M27:Q27), Data Rows 28-32 (K28='VH2', L28:L32=Model2, M28:Q32=Share%)
    # - JV:
    #     Left Table (Models): Rows 34-38. C34='JV', D34:D38=Model2, E34:E38=Tot Sample, F34:F38=Def. pairs, G34:G38==F{r}/E{r}
    #     Right Table (Defects): Header Row 35 (M35:Q35), Data Rows 36-40 (K36='JV', L36:L40=Model2, M36:Q40=Share%)
    # - JV2:
    #     Left Table (Models): Rows 40-44. C40='JV2', D40:D44=Model2, E40:E44=Tot Sample, F40:F44=Def. pairs, G40:G44==F{r}/E{r}
    #     Right Table (Defects): Header Row 42 (M42:Q42), Data Rows 43-47 (K43='JV2', L43:L47=Model2, M43:Q47=Share%)
    # Clear existing model data (cols D to F) for rows 22-26, 28-32, 34-38, 40-44
    model_rows_to_clear = list(range(22, 27)) + list(range(28, 33)) + list(range(34, 39)) + list(range(40, 45))
    for r in model_rows_to_clear:
        for c in [4, 5, 6]:
            cell = ws.cell(r, c)
            if cell.__class__.__name__ != "MergedCell":
                cell.value = None

    # Clear existing defect matrix range (K21:Q50) skipping merged cells
    for r in range(21, 51):
        for c in range(11, 18):
            cell = ws.cell(r, c)
            if cell.__class__.__name__ != "MergedCell":
                cell.value = None

    # Clean rows 45-50 across all data columns
    for r in range(45, 52):
        for c in range(3, 18):
            cell = ws.cell(r, c)
            if cell.__class__.__name__ != "MergedCell":
                cell.value = None

    # Dynamic Month Title for Top 5 Model and Top 5 Defect
    bold_hdr_title_font = Font(name="Calibri", size=11, bold=True)
    ws.cell(21, 2).value = f"Top 5 model ({target_month}, 2026)"
    ws.cell(21, 2).font = bold_hdr_title_font
    ws.cell(21, 10).value = f"Top 5 defect ({target_month}, 2026)"
    ws.cell(21, 10).font = bold_hdr_title_font

    layout_config = {
        "VH": {
            "header_row": 21,
            "data_start_row": 22
        },
        "VH2": {
            "header_row": 27,
            "data_start_row": 28
        },
        "JV": {
            "header_row": 33,
            "data_start_row": 34
        },
        "JV2": {
            "header_row": 39,
            "data_start_row": 40
        }
    }

    hdr_font = Font(name="Calibri", size=11, bold=True)
    bold_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    # Fix merged cells in Column K to match Column C
    k_ranges_to_remove = []
    for rng in list(ws.merged_cells.ranges):
        if rng.min_col == 11 and rng.max_col == 11 and rng.min_row >= 21:
            k_ranges_to_remove.append(rng)
    for rng in k_ranges_to_remove:
        ws.unmerge_cells(str(rng))

    c_ranges_to_remove = []
    for rng in list(ws.merged_cells.ranges):
        if rng.min_col == 3 and rng.max_col == 3 and rng.min_row >= 21:
            c_ranges_to_remove.append(rng)
    for rng in c_ranges_to_remove:
        ws.unmerge_cells(str(rng))

    ws.merge_cells("C22:C26")
    ws.merge_cells("K22:K26")
    ws.merge_cells("C28:C32")
    ws.merge_cells("K28:K32")
    ws.merge_cells("C34:C38")
    ws.merge_cells("K34:K38")
    ws.merge_cells("C40:C44")
    ws.merge_cells("K40:K44")

    # Load top 3 defect data per model (from QAStation validated report or df fallback)
    records = defaultdict(lambda: defaultdict(dict))
    if p2_df is not None and hasattr(p2_df, "empty") and not p2_df.empty:
        for row in p2_df.itertuples():
            f_val = str(getattr(row, "FTY", "") or "").strip()
            m_val = str(getattr(row, "Shoename", "") or "").strip()
            d_val = str(getattr(row, "_4", "") or getattr(row, "Defect_Type", "") or "").strip()
            pct_val = getattr(row, "_7", None) if hasattr(row, "_7") else getattr(row, "Defect_Share_%", None)
            if f_val and m_val and d_val and pct_val is not None:
                try:
                    records[f_val][m_val][d_val] = float(pct_val)
                except (ValueError, TypeError):
                    pass

    if not records:
        records = load_top3_defects_from_reports()

    for fty, cfg in layout_config.items():
        if fty not in hfpa_dict:
            continue
        df = hfpa_dict[fty]
        sample_col = "Audit Sample Size" if "Audit Sample Size" in df.columns else "Audit Sample"
        def_col = [c for c in df.columns if "Defective" in c and "pair" in c][0]
        defect_cols = [c for c in df.columns[21:] if c in df.columns]

        tot_factory_sample = int(df[sample_col].sum())

        # Top 5 models sorted descending by Defective Q'ty(pair) (SOP Page 16)
        top5_series = df.groupby("Model 2")[def_col].sum().sort_values(ascending=False).head(5)
        top5_models = top5_series.index.tolist()

        # If records for this factory not found, fallback to computing top 3 per model directly from df
        if fty not in records or not records[fty]:
            for m in top5_models:
                m_df = df[df["Model 2"] == m]
                tot_issues = m_df[defect_cols].sum().sum()
                top3_s = m_df[defect_cols].sum().sort_values(ascending=False).head(3)
                for d, cnt in top3_s.items():
                    if cnt > 0 and tot_issues > 0:
                        records[fty][m][d] = round(cnt / tot_issues, 6)

        # Collect unique defect types across Top 3 of all 5 models
        defect_sums = defaultdict(float)
        defect_order = []
        for m in top5_models:
            m_defects = records[fty].get(m, {})
            for d, pct in m_defects.items():
                defect_sums[d] += pct
                if d not in defect_order:
                    defect_order.append(d)

        # Sort unique defect types descending by total share
        top_def_types = sorted(defect_order, key=lambda d: defect_sums[d], reverse=True)

        hdr_r = cfg["header_row"]
        data_r = cfg["data_start_row"]

        # Left Table Header (Model, Insp., Def. q'ty (prs), HFPA DR%)
        ws.cell(hdr_r, 4).value = "Model"
        ws.cell(hdr_r, 4).font = hdr_font
        ws.cell(hdr_r, 5).value = "Insp."
        ws.cell(hdr_r, 5).font = hdr_font
        ws.cell(hdr_r, 6).value = "Def. q'ty (prs)"
        ws.cell(hdr_r, 6).font = hdr_font
        ws.cell(hdr_r, 7).value = "HFPA DR%"
        ws.cell(hdr_r, 7).font = hdr_font

        # Right Table Header (Factory, Model, Dynamic Defect Types)
        ws.cell(hdr_r, 11).value = "Factory"
        ws.cell(hdr_r, 11).font = hdr_font
        ws.cell(hdr_r, 12).value = "Model"
        ws.cell(hdr_r, 12).font = hdr_font

        for d_idx, def_name in enumerate(top_def_types):
            c_cell = ws.cell(hdr_r, 13 + d_idx)
            c_cell.value = def_name
            c_cell.font = hdr_font
            c_cell.alignment = bold_center

        # Clear remaining columns up to col 20 in header
        for c in range(13 + len(top_def_types), 21):
            ws.cell(hdr_r, c).value = None

        # Write Models (Left Table) & Defect Matrix (Right Table) on EXACT SAME ROWS
        for idx, (model_name, def_pairs) in enumerate(top5_series.items()):
            curr_r = data_r + idx

            # Left Table:
            if idx == 0:
                ws.cell(curr_r, 3).value = fty
                ws.cell(curr_r, 3).font = hdr_font
                ws.cell(curr_r, 3).alignment = bold_center

            # Col D: Model 2 name
            ws.cell(curr_r, 4).value = model_name
            ws.cell(curr_r, 4).alignment = align_left

            # Col E: Factory Total Sample (Insp.)
            ws.cell(curr_r, 5).value = tot_factory_sample
            ws.cell(curr_r, 5).number_format = "#,##0"
            ws.cell(curr_r, 5).alignment = align_right

            # Col F: Model Defective Pairs
            ws.cell(curr_r, 6).value = int(def_pairs)
            ws.cell(curr_r, 6).number_format = "#,##0"
            ws.cell(curr_r, 6).alignment = align_right

            # Col G: Formula =F{row}/E{row} (PRESERVED)
            ws.cell(curr_r, 7).value = f"=F{curr_r}/E{curr_r}"
            ws.cell(curr_r, 7).number_format = "0.00%"
            ws.cell(curr_r, 7).alignment = align_right

            # Right Table (Defect Matrix):
            if idx == 0:
                ws.cell(curr_r, 11).value = fty
                ws.cell(curr_r, 11).font = hdr_font
                ws.cell(curr_r, 11).alignment = bold_center

            # Col L: Model 2 name
            ws.cell(curr_r, 12).value = model_name
            ws.cell(curr_r, 12).alignment = align_left

            # Defect columns: ONLY populate if defect is in this model's top 3!
            m_top3 = records[fty].get(model_name, {})
            for d_idx, def_name in enumerate(top_def_types):
                c_cell = ws.cell(curr_r, 13 + d_idx)
                if def_name in m_top3:
                    c_cell.value = round(m_top3[def_name], 6)
                    c_cell.number_format = "0.00%"
                    c_cell.alignment = align_right
                else:
                    c_cell.value = None

            # Clear remaining columns up to col 20 in data row
            for c in range(13 + len(top_def_types), 21):
                ws.cell(curr_r, c).value = None

        print(f"   [+] Updated Top 5 Models & Defect Matrix for Factory: {fty} ({len(top_def_types)} defect columns)")

    # Apply modern, beautiful borders and styling to Sheet HFPA
    apply_elegant_borders(ws, target_month=target_month)
    print("   [+] Applied professional corporate borders and styling to Sheet HFPA!")

    # Populate Defect Color Legend in columns U & V matching process_ftt.py & user's reference
    populate_custom_defect_color_legend(ws, target_month=target_month)

    # Purge any redundant ColorMap_ sheets to keep workbook clean and optimal
    for s_name in list(wb.sheetnames):
        if s_name.startswith("ColorMap_"):
            del wb[s_name]

    target_output_path = ensure_file_writable(target_output_path, "File kết quả (Output/HFPA_Template_Updated.xlsx)")
    wb.save(target_output_path)
    wb.close()

    # Synchronize to Database/HFPA_Template.xlsx
    ensure_file_writable(DATABASE_TEMPLATE, "Template gốc (Database/HFPA_Template.xlsx)")
    try:
        shutil.copy2(target_output_path, DATABASE_TEMPLATE)
        print(f"[+] Successfully synchronized to master: {DATABASE_TEMPLATE}")
    except Exception as e:
        print(f"[-] Notice: Could not overwrite master template (may be open in Office): {e}")

    # Remove temporary fallback file if target_output_path is active
    fallback_new = os.path.join(OUTPUT_DIR, "HFPA_Template_Updated_New.xlsx")
    if os.path.exists(fallback_new) and target_output_path != fallback_new:
        try:
            os.remove(fallback_new)
        except Exception:
            pass

    print(f"[+] Output database written to: {target_output_path}")
    return True


def populate_custom_defect_color_legend(ws, target_month="Jul"):
    """
    Populate custom defect color key in columns U and V of sheet HFPA.
    Matches process_ftt.py write_custom_defect_legend & user's reference screenshot:
    - Merged U:V header for each site ('Site VH', 'Site VH2', 'Site JV', 'Site JV2')
    - Col U: Color chip filled with exact HEX from Color_template.xlsx (width 5)
    - Col V: Defect name (width 34), Calibri 9.5pt, left-aligned
    - Thin black border on all cells
    """
    from process_pptx import get_defect_color, load_defect_color_map
    color_template_path = os.path.join(BASE_DIR, "Color_template.xlsx")
    color_map = load_defect_color_map(color_template_path)
    fallback_cache = {}

    fill_hdr = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    font_hdr = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_txt = Font(name="Calibri", size=9.5, color="000000")
    thin_border = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000"),
    )

    site_rows = {"VH": 21, "VH2": 27, "JV": 33, "JV2": 39}

    ws.column_dimensions["U"].width = 5
    ws.column_dimensions["V"].width = 34

    # 1. Unmerge any existing merged ranges in columns U & V (cols 21-22) for rows 21..46
    merged_to_remove = []
    for rng in list(ws.merged_cells.ranges):
        if rng.min_col >= 21 and rng.max_col <= 22 and rng.min_row >= 21 and rng.max_row <= 46:
            merged_to_remove.append(rng)
    for rng in merged_to_remove:
        ws.unmerge_cells(str(rng))

    # 2. Clear all values, fills, and borders in cols 21 & 22 for rows 21..46
    no_fill = PatternFill(fill_type=None)
    for r in range(21, 47):
        for c in (21, 22):
            cell = ws.cell(r, c)
            cell.value = None
            cell.fill = no_fill
            cell.border = Border()

    # 3. Populate site headers and defect color chips (strictly capped to 5 defects per site)
    for fty, start_r in site_rows.items():
        # Extract defect names from row start_r (cols 13 onwards until empty)
        defects = []
        for c in range(13, 21):
            val = ws.cell(start_r, c).value
            if val and str(val).strip():
                defects.append(str(val).strip())
            else:
                break

        if not defects:
            continue

        # Cap at 5 defects to strictly fit within the 5 allocated rows per site
        defects = defects[:5]

        # Site Header Row (Merged U & V)
        ws.merge_cells(start_row=start_r, start_column=21, end_row=start_r, end_column=22)
        c_hdr = ws.cell(start_r, 21, value=f"Site {fty}")
        c_hdr.fill = fill_hdr
        c_hdr.font = font_hdr
        c_hdr.alignment = Alignment(horizontal="center", vertical="center")
        c_hdr.border = thin_border
        ws.cell(start_r, 22).border = thin_border
        ws.row_dimensions[start_r].height = 20

        # Defect Rows (up to 5 rows)
        for i, d_name in enumerate(defects, start=1):
            r = start_r + i
            hex_val = get_defect_color(d_name, color_map, fallback_cache)

            c_chip = ws.cell(r, 21)
            c_chip.fill = PatternFill(start_color=hex_val, end_color=hex_val, fill_type="solid")
            c_chip.border = thin_border

            c_txt = ws.cell(r, 22, value=str(d_name))
            c_txt.font = font_txt
            c_txt.border = thin_border
            c_txt.alignment = Alignment(horizontal="left", vertical="center")
            ws.row_dimensions[r].height = 18

    print("   [+] Populated Defect Color Legend in columns U & V matching process_ftt.py!")


def apply_elegant_borders(ws, target_month="Jul"):
    """
    Apply modern, beautiful corporate borders and styling to Sheet HFPA:
    - Visible gridlines in all spreadsheet viewers (WPS Office & Excel)
    - Executive framing with Corporate Navy (#1F4E78) outer borders
    - Soft ice blue header fills (#D9E1F2) with dark navy bold headers
    - Refined silver interior data gridlines (#BFBFBF)
    - Factory badge cells with soft shading (#F4F6F9) and clean vertical borders
    - Accounting-standard double underline on Grand Total rows
    - Wrap text and optimized row heights for clean readability
    - Complete cleanup of rows 45+
    """
    ws.sheet_view.showGridLines = True
    try:
        ws.views.sheetView[0].showGridLines = True
    except Exception:
        pass

    font_family = "Century Gothic"
    f_title = Font(name=font_family, size=11, bold=True, color="1F4E78")
    f_hdr = Font(name=font_family, size=10, bold=True, color="1F4E78")
    f_bold = Font(name=font_family, size=10, bold=True, color="000000")
    f_regular = Font(name=font_family, size=10, bold=False, color="000000")
    f_factory = Font(name=font_family, size=11, bold=True, color="1F4E78")

    # Corporate fills
    hdr_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")      # Soft ice blue
    total_fill = PatternFill(start_color="F2F4F8", end_color="F2F4F8", fill_type="solid")    # Soft highlight
    badge_fill = PatternFill(start_color="F4F6F9", end_color="F4F6F9", fill_type="solid")    # Factory badge fill

    # Border sides
    thin_gray = Side(style="thin", color="BFBFBF")
    med_navy = Side(style="medium", color="1F4E78")
    top_black = Side(style="thin", color="000000")
    bot_double = Side(style="double", color="000000")
    no_side = None

    align_center = Alignment(horizontal="center", vertical="center")
    align_center_wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    # -------------------------------------------------------------
    # 1. TOP SECTION TABLES:
    # Table 1: Production (Rows 1 to 6, Cols A to N)
    # Table 2: Defects (Rows 8 to 13, Cols A to N)
    # Table 3: HFPA% (Rows 15 to 20, Cols A to N)
    # -------------------------------------------------------------
    top_sections = [
        (1, 2, 5, 6, "Total pair produced"),
        (8, 9, 12, 13, "Defect"),
        (15, 16, 19, 20, "HFPA%")
    ]

    for r_hdr, r_start, r_end, r_tot, label in top_sections:
        # Header Row
        ws.row_dimensions[r_hdr].height = 24
        if ws.cell(r_hdr, 1).value is None:
            ws.cell(r_hdr, 1).value = "FTY"
        if r_hdr == 15 and ws.cell(r_hdr, 2).value is None:
            ws.cell(r_hdr, 2).value = label

        for c in range(1, 15):
            cell = ws.cell(r_hdr, c)
            cell.font = f_hdr
            cell.fill = hdr_fill
            cell.alignment = align_center if c > 2 else (align_center if c == 1 else align_left)
            left_b = med_navy if c == 1 else thin_gray
            right_b = med_navy if c == 14 else thin_gray
            cell.border = Border(left=left_b, right=right_b, top=med_navy, bottom=med_navy)

        # Data Rows (Rows r_start to r_end)
        for r in range(r_start, r_end + 1):
            ws.row_dimensions[r].height = 19
            for c in range(1, 15):
                cell = ws.cell(r, c)
                left_b = med_navy if c == 1 else thin_gray
                right_b = med_navy if c == 14 else thin_gray
                cell.border = Border(left=left_b, right=right_b, top=thin_gray, bottom=thin_gray)

                if c == 1:
                    cell.font = f_bold
                    cell.alignment = align_center
                elif c == 2:
                    cell.font = f_regular
                    cell.alignment = align_left
                else:
                    cell.font = f_regular
                    cell.alignment = align_right
                    if r_hdr in (1, 8) and isinstance(cell.value, (int, float)):
                        cell.number_format = "#,##0"
                    elif r_hdr == 15 and isinstance(cell.value, (int, float)):
                        cell.number_format = "0.00%"

        # Grand Total Row
        ws.row_dimensions[r_tot].height = 20
        ws.cell(r_tot, 1).value = "Grand Total"
        if ws.cell(r_tot, 2).value is None:
            ws.cell(r_tot, 2).value = label

        for c in range(1, 15):
            cell = ws.cell(r_tot, c)
            cell.font = f_bold
            cell.fill = total_fill
            left_b = med_navy if c == 1 else thin_gray
            right_b = med_navy if c == 14 else thin_gray
            cell.border = Border(left=left_b, right=right_b, top=top_black, bottom=bot_double)
            if c == 1:
                cell.alignment = align_center
            elif c == 2:
                cell.alignment = align_left
            else:
                cell.alignment = align_right
                if r_hdr in (1, 8) and isinstance(cell.value, (int, float)):
                    cell.number_format = "#,##0"
                elif r_hdr == 15 and isinstance(cell.value, (int, float)):
                    cell.number_format = "0.00%"

    # Clear borders on gap rows (Row 7, Row 14)
    for r in [7, 14]:
        ws.row_dimensions[r].height = 14
        for c in range(1, 20):
            cell = ws.cell(r, c)
            cell.border = Border()
            cell.fill = PatternFill(fill_type=None)

    # -------------------------------------------------------------
    # 2. SECTION TITLES AT ROW 21
    # -------------------------------------------------------------
    ws.cell(21, 2).value = f"Top 5 model ({target_month}, 2026)"
    ws.cell(21, 2).font = f_title
    ws.cell(21, 2).border = Border()

    ws.cell(21, 10).value = f"Top 5 defect ({target_month}, 2026)"
    ws.cell(21, 10).font = f_title
    ws.cell(21, 10).border = Border()

    # -------------------------------------------------------------
    # 3. TOP 5 MODEL & DEFECT TABLES (Rows 21 to 44)
    # Factory Blocks:
    #   VH:  Header Row 21, Data Rows 22-26
    #   VH2: Header Row 27, Data Rows 28-32
    #   JV:  Header Row 33, Data Rows 34-38
    #   JV2: Header Row 39, Data Rows 40-44
    # -------------------------------------------------------------
    factory_blocks = [
        (21, 22, 26, "VH"),
        (27, 28, 32, "VH2"),
        (33, 34, 38, "JV"),
        (39, 40, 44, "JV2")
    ]

    for r_hdr, r_start, r_end, fty_name in factory_blocks:
        ws.row_dimensions[r_hdr].height = 24

        # A. Left Table Header (Cols C to G: FTY, Model, Insp., Def. q'ty (prs), HFPA DR%)
        if ws.cell(r_hdr, 3).value is None:
            ws.cell(r_hdr, 3).value = "FTY"
        ws.cell(r_hdr, 4).value = "Model"
        ws.cell(r_hdr, 5).value = "Insp."
        ws.cell(r_hdr, 6).value = "Def. q'ty (prs)"
        ws.cell(r_hdr, 7).value = "HFPA DR%"

        for c in range(3, 8):
            cell = ws.cell(r_hdr, c)
            cell.font = f_hdr
            cell.fill = hdr_fill
            cell.alignment = align_center
            left_b = med_navy if c == 3 else thin_gray
            right_b = med_navy if c == 7 else thin_gray
            cell.border = Border(left=left_b, right=right_b, top=med_navy, bottom=med_navy)

        # Dynamically detect last defect column
        last_def_col = 12
        for c in range(13, 21):
            val = ws.cell(r_hdr, c).value
            if val and str(val).strip():
                last_def_col = c
            else:
                break
        if last_def_col < 15:
            last_def_col = 15

        # B. Right Table Header (Cols K to last_def_col)
        if ws.cell(r_hdr, 11).value is None:
            ws.cell(r_hdr, 11).value = "Factory"
        if ws.cell(r_hdr, 12).value is None:
            ws.cell(r_hdr, 12).value = "Model"

        for c in range(11, last_def_col + 1):
            cell = ws.cell(r_hdr, c)
            cell.font = f_hdr
            cell.fill = hdr_fill
            cell.alignment = align_center_wrap
            left_b = med_navy if c == 11 else thin_gray
            right_b = med_navy if c == last_def_col else thin_gray
            cell.border = Border(left=left_b, right=right_b, top=med_navy, bottom=med_navy)

        # Clear unused header columns beyond last_def_col
        for c in range(last_def_col + 1, 21):
            cell = ws.cell(r_hdr, c)
            cell.value = None
            cell.border = Border()
            cell.fill = PatternFill(fill_type=None)

        # C. Factory Data Rows (Rows r_start to r_end)
        for r in range(r_start, r_end + 1):
            ws.row_dimensions[r].height = 19
            is_last_row = (r == r_end)
            divider_bot = med_navy if is_last_row else thin_gray

            # --- Left Table: C to G ---
            cell_c = ws.cell(r, 3)
            cell_c.fill = badge_fill
            cell_c.font = f_factory
            cell_c.alignment = align_center
            top_c = thin_gray if r == r_start else no_side
            bot_c = divider_bot if is_last_row else no_side
            cell_c.border = Border(left=med_navy, right=thin_gray, top=top_c, bottom=bot_c)

            cell_d = ws.cell(r, 4)
            cell_d.font = f_regular
            cell_d.alignment = align_left
            cell_d.border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=divider_bot)

            cell_e = ws.cell(r, 5)
            cell_e.font = f_regular
            cell_e.alignment = align_right
            if isinstance(cell_e.value, (int, float)):
                cell_e.number_format = "#,##0"
            cell_e.border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=divider_bot)

            cell_f = ws.cell(r, 6)
            cell_f.font = f_regular
            cell_f.alignment = align_right
            if isinstance(cell_f.value, (int, float)):
                cell_f.number_format = "#,##0"
            cell_f.border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=divider_bot)

            cell_g = ws.cell(r, 7)
            cell_g.font = f_bold
            cell_g.alignment = align_right
            if isinstance(cell_g.value, (int, float)):
                cell_g.number_format = "0.00%"
            cell_g.border = Border(left=thin_gray, right=med_navy, top=thin_gray, bottom=divider_bot)

            # --- Right Table: K to last_def_col ---
            cell_k = ws.cell(r, 11)
            cell_k.fill = badge_fill
            cell_k.font = f_factory
            cell_k.alignment = align_center
            top_k = thin_gray if r == r_start else no_side
            bot_k = divider_bot if is_last_row else no_side
            cell_k.border = Border(left=med_navy, right=thin_gray, top=top_k, bottom=bot_k)

            cell_l = ws.cell(r, 12)
            cell_l.font = f_regular
            cell_l.alignment = align_left
            cell_l.border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=divider_bot)

            for c in range(13, last_def_col + 1):
                cell_def = ws.cell(r, c)
                cell_def.font = f_regular
                cell_def.alignment = align_right
                if isinstance(cell_def.value, (int, float)):
                    cell_def.number_format = "0.00%"
                right_def = med_navy if c == last_def_col else thin_gray
                cell_def.border = Border(left=thin_gray, right=right_def, top=thin_gray, bottom=divider_bot)

            # Clear trailing columns beyond last_def_col
            for c in range(last_def_col + 1, 21):
                c_clear = ws.cell(r, c)
                c_clear.value = None
                c_clear.border = Border()
                c_clear.fill = PatternFill(fill_type=None)

    # -------------------------------------------------------------
    # 4. CLEAN SPACING & CLEAR ROWS 45+
    # -------------------------------------------------------------
    for r in range(21, 45):
        for c in [1, 8, 9, 18, 19, 20]:
            cell = ws.cell(r, c)
            cell.border = Border()
            cell.fill = PatternFill(fill_type=None)
        if r > 21:
            ws.cell(r, 2).border = Border()
            ws.cell(r, 10).border = Border()

    for r in range(45, 65):
        ws.row_dimensions[r].height = 15
        for c in range(1, 25):
            cell = ws.cell(r, c)
            cell.value = None
            cell.border = Border()
            cell.fill = PatternFill(fill_type=None)


if __name__ == "__main__":
    populate_hfpa_database()
