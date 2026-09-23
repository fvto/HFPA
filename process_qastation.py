"""
HFPA & FTT Data Processing Tool
Standard Operating Procedure (SOP): Database/2026 FTT&HFPA EN.pdf

Fully automates the 17-page official SOP workflow:
1. Pages 1-5: Clean HFPA (Quality Tracking) data & build 1st Pivot Table with Plant between Station and Line.
2. Pages 6-15: Reconcile HFPA (Quality Tracking) vs HFPA (Mes410) by Factory, Line, Date, Model, and correct mismatched counts.
3. Page 16: Calculate Top 5 Models sorted descending by Fail Q'ty (DEFECT RATE = Fail Q'ty / Sum Insp Q'ty).
4. Page 17: Calculate Top Defect Types per Top Model (Defect Rate% = Issues Qty / Total Issue Q'ty) transposed to summary table.
5. Export separate combined workbooks and auto-populate Database/HFPA_Template.xlsx.
"""

import os
import sys
import glob
import re
import datetime

if sys.stdout.encoding != "utf-8":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from file_utils import ensure_file_writable, auto_archive_previous_months, detect_file_month

# Paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_FTT_DIR = os.path.join(BASE_DIR, "Input", "FTT")
INPUT_HFPA_DIR = os.path.join(BASE_DIR, "Input", "HFPA")
OUTPUT_DIR = os.path.join(BASE_DIR, "Output")
COLOR_TEMPLATE_PATH = os.path.join(BASE_DIR, "Color_template.xlsx")

# Expected Factories
FACTORIES = ["JV", "JV2", "VH", "VH2"]


def extract_fty_from_filename(filename: str) -> str:
    """Extract factory identifier (JV, JV2, VH, VH2) from filename."""
    base = os.path.splitext(os.path.basename(filename))[0]
    for fty in ["JV2", "VH2", "JV", "VH"]:
        if re.search(rf"[-_\s]{fty}$", base, re.IGNORECASE) or re.search(rf"[-_\s]{fty}[-_\s]", base, re.IGNORECASE):
            return fty
    for fty in ["JV2", "VH2", "JV", "VH"]:
        if fty in base.upper():
            return fty
    return "UNKNOWN"


def load_color_template():
    """Load defect color mappings from Color_template.xlsx."""
    color_map = {}
    if not os.path.exists(COLOR_TEMPLATE_PATH):
        print("[-] Color_template.xlsx not found, skipping color mapping.")
        return color_map

    try:
        wb = openpyxl.load_workbook(COLOR_TEMPLATE_PATH, data_only=True)
        if "Sheet1" in wb.sheetnames:
            ws = wb["Sheet1"]
            for row in ws.iter_rows(min_row=5, values_only=True):
                if len(row) >= 8 and row[2] and row[7]:
                    defect_name = str(row[2]).strip()
                    hex_code = str(row[7]).strip().replace("#", "")
                    if len(hex_code) == 6:
                        color_map[defect_name.lower()] = hex_code
        
        if "bc" in wb.sheetnames:
            ws = wb["bc"]
            for row in ws.iter_rows(min_row=2, values_only=True):
                if len(row) >= 5 and row[0] and row[4]:
                    defect_name = str(row[0]).strip()
                    hex_code = str(row[4]).strip().replace("#", "")
                    if len(hex_code) == 6:
                        color_map[defect_name.lower()] = hex_code
        wb.close()
        print(f"[+] Loaded {len(color_map)} color definitions from Color_template.xlsx")
    except Exception as e:
        print(f"[-] Warning: Failed to load color template: {e}")
    return color_map


def purge_colormap_sheets(wb):
    """Purge any redundant ColorMap_ sheets from workbook."""
    purged = []
    for s_name in list(wb.sheetnames):
        if s_name.startswith("ColorMap_"):
            del wb[s_name]
            purged.append(s_name)
    return purged


def embed_color_template_sheets(wb, color_file_path=COLOR_TEMPLATE_PATH):
    """Deprecated: Color_template.xlsx is read directly; do not duplicate sheets."""
    return purge_colormap_sheets(wb)


def parse_plant_and_line(val):
    """
    Parse Plant and Line from strings like 'N1 Line 1', 'Line 1', 'A Line1', 'A10 Line 1'
    Returns: (Plant, Line)
    """
    if pd.isna(val):
        return "", ""
    s = str(val).strip()
    m = re.match(r"^([A-Za-z0-9]+)\s*Line\s*([0-9]+)", s, re.IGNORECASE)
    if m:
        return m.group(1).upper(), m.group(2)
    m2 = re.match(r"^Line\s*([0-9]+)", s, re.IGNORECASE)
    if m2:
        return "", m2.group(1)
    # numeric line
    cleaned = re.sub(r"[^0-9]", "", s)
    return "", cleaned if cleaned else s


def load_and_combine_ftt():
    """Discover, load and combine all QAStation/FTT files into a single DataFrame."""
    ftt_files = glob.glob(os.path.join(INPUT_FTT_DIR, "*.xlsx")) + glob.glob(os.path.join(INPUT_FTT_DIR, "*.xls"))
    if not ftt_files:
        raise FileNotFoundError(f"No QAStation/FTT files found in {INPUT_FTT_DIR}")

    print(f"\n[1/5] Loading & combining {len(ftt_files)} QAStation/FTT files...")
    df_list = []
    for filepath in sorted(ftt_files):
        filename = os.path.basename(filepath)
        fty = extract_fty_from_filename(filename)
        print(f"   -> Reading: {filename} (Detected FTY: {fty})")
        df = pd.read_excel(filepath)
        df["FTY"] = fty
        df["Source_File"] = filename

        # Standardize column names between QAStation format and FTT-Internal format
        # If QAStation format: 'InspDate', 'Line', 'ShoeName', 'InspQty', 'FailQty', 'Station', 'Issues', 'IssueQty'
        if "InspDate" in df.columns:
            df["Date_Raw"] = df["InspDate"]
            df["Model_Raw"] = df["ShoeName"] if "ShoeName" in df.columns else df.get("Shoename", "")
            df["InspQty_Raw"] = pd.to_numeric(df["InspQty"], errors="coerce").fillna(0)
            df["FailQty_Raw"] = pd.to_numeric(df["FailQty"], errors="coerce").fillna(0)
            df["Issues_Raw"] = df["Issues"] if "Issues" in df.columns else ""
            df["IssueQty_Raw"] = pd.to_numeric(df["IssueQty"], errors="coerce").fillna(0)
            df["Station_Raw"] = df["Station"] if "Station" in df.columns else "HFPA"
            # Parse Plant & Line from 'Line' col
            plants, lines = zip(*df["Line"].apply(parse_plant_and_line))
            df["Plant_Parsed"] = plants
            df["Line_Parsed"] = lines
        else:
            # FTT-Internal format
            df["Date_Raw"] = df["Date"]
            df["Model_Raw"] = df["Model"] if "Model" in df.columns else df.get("ShoeName", "")
            df["InspQty_Raw"] = pd.to_numeric(df["Total Pair Produced"], errors="coerce").fillna(0)
            df["FailQty_Raw"] = pd.to_numeric(df["Issues Q'ty"], errors="coerce").fillna(0)
            df["Issues_Raw"] = df["Defect Issues"] if "Defect Issues" in df.columns else ""
            df["IssueQty_Raw"] = pd.to_numeric(df["Issues Q'ty"], errors="coerce").fillna(0)
            df["Station_Raw"] = df["Station"] if "Station" in df.columns else "Assembly"
            df["Plant_Parsed"] = df["Plant"].astype(str).str.strip().str.upper() if "Plant" in df.columns else ""
            _, lines = zip(*df["Line"].apply(parse_plant_and_line))
            df["Line_Parsed"] = lines

        df_list.append(df)

    combined_ftt = pd.concat(df_list, ignore_index=True)
    print(f"   [+] QAStation/FTT Combined rows: {len(combined_ftt):,}")
    return combined_ftt


def load_and_combine_hfpa():
    """Discover, load and combine all HFPA files into a single DataFrame."""
    hfpa_files = glob.glob(os.path.join(INPUT_HFPA_DIR, "*.xlsx")) + glob.glob(os.path.join(INPUT_HFPA_DIR, "*.xls"))
    if not hfpa_files:
        raise FileNotFoundError(f"No HFPA files found in {INPUT_HFPA_DIR}")

    print(f"\n[2/5] Loading & combining {len(hfpa_files)} HFPA files...")
    df_list = []
    for filepath in sorted(hfpa_files):
        filename = os.path.basename(filepath)
        fty = extract_fty_from_filename(filename)
        print(f"   -> Reading: {filename} (Detected FTY: {fty})")
        df = pd.read_excel(filepath)
        df["FTY"] = fty
        df["Source_File"] = filename
        df_list.append(df)

    combined_hfpa = pd.concat(df_list, ignore_index=True)
    print(f"   [+] HFPA Combined rows: {len(combined_hfpa):,}")
    return combined_hfpa


def validate_and_correct_data(qa_df: pd.DataFrame, hfpa_df: pd.DataFrame):
    """
    Validate QAStation data against HFPA:
    - Standardize Date, Line, Plant
    - Match lots between QAStation and HFPA by Date, FTY, Plant, Line
    - Compare defect quantities: QAStation FailQty / IssueQty vs HFPA Defect Q'ty
    - Update/correct QAStation values referencing HFPA data
    """
    print("\n[3/5] Validating QAStation data against HFPA and adjusting defect counts...")

    # Standardize QAStation
    qa_df["Date_Std"] = pd.to_datetime(qa_df["Date_Raw"], errors="coerce").dt.strftime("%Y-%m-%d")
    qa_df["Plant_Std"] = qa_df["Plant_Parsed"].astype(str).str.strip().str.upper()
    qa_df["Line_Std"] = qa_df["Line_Parsed"].astype(str).str.strip()
    qa_df["ShoeName_Std"] = qa_df["Model_Raw"].astype(str).str.strip()

    # Standardize HFPA
    hfpa_date_col = "Audit Date (dd-mmm-yyyy)" if "Audit Date (dd-mmm-yyyy)" in hfpa_df.columns else "Audit Date"
    hfpa_df["Date_Std"] = pd.to_datetime(hfpa_df[hfpa_date_col], errors="coerce").dt.strftime("%Y-%m-%d")
    hfpa_df["Plant_Std"] = hfpa_df["Plant"].astype(str).str.strip().str.upper()
    _, hfpa_lines = zip(*hfpa_df["Line"].apply(parse_plant_and_line))
    hfpa_df["Line_Std"] = [str(l).strip() for l in hfpa_lines]

    defect_tot_col = "Defect Q'ty (Total Defect Q'ty)" if "Defect Q'ty (Total Defect Q'ty)" in hfpa_df.columns else "Defective Q'ty(pair)"
    sample_col = "Audit Sample Size" if "Audit Sample Size" in hfpa_df.columns else "Audit Sample"

    # Aggregate HFPA at lot level (Date, FTY, Plant, Line)
    hfpa_lot = hfpa_df.groupby(["Date_Std", "FTY", "Plant_Std", "Line_Std"]).agg(
        HFPA_Audit_Sample=(sample_col, "sum"),
        HFPA_Defect_Qty=(defect_tot_col, "sum")
    ).reset_index()

    # Aggregate QAStation at lot level (Date, FTY, Plant, Line)
    # In QAStation, InspQty and FailQty are line/shoe totals. Total defect issue count is sum of IssueQty_Raw.
    qa_lot = qa_df.groupby(["Date_Std", "FTY", "Plant_Std", "Line_Std"]).agg(
        QA_InspQty=("InspQty_Raw", "max"),
        QA_FailQty=("FailQty_Raw", "max"),
        QA_IssueQty_Sum=("IssueQty_Raw", "sum")
    ).reset_index()

    # Merge QA and HFPA at lot level
    lot_comparison = pd.merge(
        qa_lot,
        hfpa_lot,
        on=["Date_Std", "FTY", "Plant_Std", "Line_Std"],
        how="outer"
    )

    lot_comparison["Defect_Variance"] = lot_comparison["QA_IssueQty_Sum"] - lot_comparison["HFPA_Defect_Qty"]

    def get_status(row):
        if pd.isna(row["QA_IssueQty_Sum"]):
            return "HFPA_Only"
        elif pd.isna(row["HFPA_Defect_Qty"]):
            return "QAStation_Only"
        elif row["Defect_Variance"] == 0:
            return "Matched_Exact"
        else:
            return "Different_Count"

    lot_comparison["Validation_Status"] = lot_comparison.apply(get_status, axis=1)

    matched_count = len(lot_comparison[lot_comparison["Validation_Status"] == "Matched_Exact"])
    diff_count = len(lot_comparison[lot_comparison["Validation_Status"] == "Different_Count"])
    qa_only = len(lot_comparison[lot_comparison["Validation_Status"] == "QAStation_Only"])
    hfpa_only = len(lot_comparison[lot_comparison["Validation_Status"] == "HFPA_Only"])

    print(f"   [+] Lot Validation Summary (Date, FTY, Plant, Line):")
    print(f"       - Total unique lots: {len(lot_comparison):,}")
    print(f"       - Matched exact defect count: {matched_count:,}")
    print(f"       - Different defect count (adjusted to HFPA): {diff_count:,}")
    print(f"       - In QAStation only: {qa_only:,}")
    print(f"       - In HFPA only: {hfpa_only:,}")

    # Map validation fields back into QAStation rows
    qa_validated = pd.merge(
        qa_df,
        lot_comparison[["Date_Std", "FTY", "Plant_Std", "Line_Std", "HFPA_Audit_Sample", "HFPA_Defect_Qty", "Defect_Variance", "Validation_Status"]],
        on=["Date_Std", "FTY", "Plant_Std", "Line_Std"],
        how="left"
    )

    # Correct QAStation values so they match HFPA data:
    # 1. Validated FailQty for the lot matches HFPA_Defect_Qty
    # 2. Individual defect issues are scaled proportionally to match HFPA_Defect_Qty exactly
    def calc_corrected_issue_qty(r):
        if r["Validation_Status"] in ["Matched_Exact", "Different_Count"]:
            lot_qa_sum = r["QA_IssueQty_Sum"] if "QA_IssueQty_Sum" in r else None
            # fallback if not merged directly
            hfpa_tot = r["HFPA_Defect_Qty"]
            orig_qty = r["IssueQty_Raw"]
            if pd.notna(hfpa_tot):
                # If exact match or zero, return orig_qty
                if r["Validation_Status"] == "Matched_Exact":
                    return orig_qty
                # If different count, scale to match HFPA total
                lot_comp_row = lot_comparison[
                    (lot_comparison["Date_Std"] == r["Date_Std"]) &
                    (lot_comparison["FTY"] == r["FTY"]) &
                    (lot_comparison["Plant_Std"] == r["Plant_Std"]) &
                    (lot_comparison["Line_Std"] == r["Line_Std"])
                ]
                if not lot_comp_row.empty:
                    qa_sum = lot_comp_row["QA_IssueQty_Sum"].values[0]
                    if qa_sum > 0:
                        return round(orig_qty * (hfpa_tot / qa_sum), 2)
            return orig_qty
        return r["IssueQty_Raw"]

    qa_validated["IssueQty_Validated"] = qa_validated.apply(calc_corrected_issue_qty, axis=1)

    # Validated InspQty (matches HFPA Sample size if available)
    qa_validated["InspQty_Validated"] = qa_validated.apply(
        lambda r: r["HFPA_Audit_Sample"] if pd.notna(r["HFPA_Audit_Sample"]) and r["HFPA_Audit_Sample"] > 0 else r["InspQty_Raw"],
        axis=1
    )

    return qa_validated, hfpa_df, lot_comparison


def generate_pivot1_and_top5(qa_validated: pd.DataFrame):
    """
    Pivot Table #1 — production / failure validation
    Fields:
    - InspDate, FTY, Station, Line, Shoename, InspQty, FailQty
    - Top 5 shoes with highest failures
    - DR% = Sum of FailQty / Grand Total (InspQty)
    """
    print("\n[4/5] Generating Pivot #1 and identifying Top 5 Shoes...")

    # Pivot 1 base: Date, FTY, Station, Plant, Line, ShoeName
    p1 = qa_validated.groupby(["Date_Std", "FTY", "Station_Raw", "Plant_Std", "Line_Std", "ShoeName_Std"]).agg(
        InspQty_Original=("InspQty_Raw", "max"),
        InspQty_Validated=("InspQty_Validated", "max"),
        FailQty_Original=("IssueQty_Raw", "sum"),
        FailQty_Validated=("IssueQty_Validated", "sum")
    ).reset_index()

    p1.rename(columns={
        "Date_Std": "InspDate",
        "Station_Raw": "Station",
        "Plant_Std": "Plant",
        "Line_Std": "Line",
        "ShoeName_Std": "Shoename",
        "InspQty_Validated": "InspQty",
        "FailQty_Validated": "FailQty"
    }, inplace=True)

    p1["DR%_Original"] = np.where(p1["InspQty_Original"] > 0, p1["FailQty_Original"] / p1["InspQty_Original"], 0.0)
    p1["DR%_Validated"] = np.where(p1["InspQty"] > 0, p1["FailQty"] / p1["InspQty"], 0.0)

    # Reorder columns so Plant is strictly between Station and Line
    cols_order = [
        "InspDate", "FTY", "Station", "Plant", "Line", "Shoename",
        "InspQty_Original", "InspQty", "FailQty_Original", "FailQty",
        "DR%_Original", "DR%_Validated"
    ]
    p1 = p1[[c for c in cols_order if c in p1.columns]]

    # Top 5 Shoes across All Factories & per Factory
    top5_dict = {}
    for fty in FACTORIES + ["Grand Total"]:
        if fty == "Grand Total":
            sub_df = qa_validated
        else:
            sub_df = qa_validated[qa_validated["FTY"] == fty]

        # Group by Shoename
        lot_prod = sub_df.groupby(["Date_Std", "Plant_Std", "Line_Std", "ShoeName_Std"])["InspQty_Validated"].max().reset_index()
        model_prod = lot_prod.groupby("ShoeName_Std")["InspQty_Validated"].sum().reset_index()
        model_prod.rename(columns={"InspQty_Validated": "Sum of InspQty"}, inplace=True)

        model_fail = sub_df.groupby("ShoeName_Std").agg(
            FailQty_Original=("IssueQty_Raw", "sum"),
            FailQty_Validated=("IssueQty_Validated", "sum")
        ).reset_index()
        model_fail.rename(columns={
            "FailQty_Original": "Sum of FailQty (Original)",
            "FailQty_Validated": "Sum of FailQty (Validated)"
        }, inplace=True)

        model_summary = pd.merge(model_prod, model_fail, on="ShoeName_Std", how="outer").fillna(0)
        model_summary["DR%"] = np.where(model_summary["Sum of InspQty"] > 0, model_summary["Sum of FailQty (Validated)"] / model_summary["Sum of InspQty"], 0.0)
        
        # Sort descending by FailQty
        model_summary = model_summary.sort_values(by="Sum of FailQty (Validated)", ascending=False).reset_index(drop=True)
        model_summary.rename(columns={"ShoeName_Std": "Shoename"}, inplace=True)

        top5 = model_summary.head(5).copy()
        top5["Rank"] = range(1, len(top5) + 1)
        top5["FTY"] = fty
        top5_dict[fty] = top5

    all_top5 = pd.concat([top5_dict[f] for f in FACTORIES + ["Grand Total"]], ignore_index=True)
    return p1, all_top5, top5_dict


def generate_pivot2_top_defects(qa_validated: pd.DataFrame, top5_dict: dict):
    """
    Pivot Table #2 — defect types
    Fields:
    - Defect type, Defect quantity
    - Top 3 defect types per selected Top 5 shoe
    - Defect % calculations
    """
    print("\n[5/5] Generating Pivot #2 (Defect Types & Top 3 per shoe)...")

    pivot2_records = []
    color_map = load_color_template()

    for fty in FACTORIES + ["Grand Total"]:
        top5_models = top5_dict[fty]["Shoename"].tolist()
        if fty == "Grand Total":
            sub_df = qa_validated[qa_validated["ShoeName_Std"].isin(top5_models)]
        else:
            sub_df = qa_validated[(qa_validated["FTY"] == fty) & (qa_validated["ShoeName_Std"].isin(top5_models))]

        defect_df = sub_df[(sub_df["IssueQty_Validated"] > 0) & (sub_df["Issues_Raw"].notna())]

        for model in top5_models:
            m_df = defect_df[defect_df["ShoeName_Std"] == model]
            total_model_defects = m_df["IssueQty_Validated"].sum()

            defect_grp = m_df.groupby("Issues_Raw")["IssueQty_Validated"].sum().reset_index()
            defect_grp.rename(columns={"Issues_Raw": "Defect Type", "IssueQty_Validated": "Defect Quantity"}, inplace=True)
            defect_grp = defect_grp.sort_values(by="Defect Quantity", ascending=False).reset_index(drop=True)

            top3 = defect_grp.head(3).copy()
            for rank, row in enumerate(top3.itertuples(), start=1):
                def_name = row._1
                def_qty = row._2

                share_pct = (def_qty / total_model_defects) if total_model_defects > 0 else 0.0
                spec_ratio = (total_model_defects / def_qty) if def_qty > 0 else 0.0
                hex_color = color_map.get(str(def_name).lower(), "")

                pivot2_records.append({
                    "FTY": fty,
                    "Shoename": model,
                    "Defect Rank": f"Top {rank}",
                    "Defect Type": def_name,
                    "Defect Quantity": def_qty,
                    "Total Shoe Defects": total_model_defects,
                    "Defect Share %": share_pct,
                    "Spec Ratio (Total/Defects)": spec_ratio,
                    "Color HEX": hex_color
                })

    pivot2_df = pd.DataFrame(pivot2_records)
    return pivot2_df


def style_excel_sheet(ws, title: str):
    """Apply modern corporate styling to openpyxl worksheet."""
    ws.views.sheetView[0].showGridLines = True

    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    border_thin = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )
    align_center = Alignment(horizontal="center", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    for col_idx in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center
        cell.border = border_thin

    max_format_row = min(ws.max_row, 500) if ws.max_row > 1000 else ws.max_row
    for row_idx in range(2, max_format_row + 1):
        fill_color = "F2F5F9" if row_idx % 2 == 0 else "FFFFFF"
        row_fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")

        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.border = border_thin
            if not cell.fill.start_color.rgb or cell.fill.start_color.rgb == "00000000":
                cell.fill = row_fill

            val = cell.value
            if isinstance(val, (int, float)):
                header_name = str(ws.cell(row=1, column=col_idx).value or "").lower()
                if "%" in header_name or "rate" in header_name or "share" in header_name:
                    cell.number_format = "0.00%"
                    cell.alignment = align_right
                elif isinstance(val, float):
                    cell.number_format = "#,##0.0"
                    cell.alignment = align_right
                else:
                    cell.number_format = "#,##0"
                    cell.alignment = align_right
            elif isinstance(val, (datetime.date, datetime.datetime)):
                cell.number_format = "YYYY-MM-DD"
                cell.alignment = align_center

    for col_idx in range(1, ws.max_column + 1):
        sample_vals = [str(ws.cell(r, col_idx).value or "") for r in range(1, min(ws.max_row, 50) + 1)]
        max_len = max(len(v) for v in sample_vals) if sample_vals else 10
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 45)


def export_deliverables(qa_validated: pd.DataFrame, hfpa_combined: pd.DataFrame,
                        lot_comparison: pd.DataFrame, p1_df: pd.DataFrame,
                        all_top5: pd.DataFrame, p2_df: pd.DataFrame):
    """
    Save 1 single unified master combined workbook (Output/HFPA_FTT_Combined_Master.xlsx)
    and the executive analysis report.
    Adheres to SOP and QAStation_Tool_Agent_Spec.md:
    - Merges Quality Tracking and Mes410 into 1 combine file (no split files).
    - Writes audit trail log to logs/reconciliation_audit_trail.csv.
    """
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    logs_dir = os.path.join(BASE_DIR, "logs")
    os.makedirs(logs_dir, exist_ok=True)

    master_combined_path = os.path.join(OUTPUT_DIR, "HFPA_FTT_Combined_Master.xlsx")
    report_out_path = os.path.join(OUTPUT_DIR, "QAStation_HFPA_Analysis_Report.xlsx")

    ensure_file_writable(master_combined_path, "Master Combined (Output/HFPA_FTT_Combined_Master.xlsx)")
    ensure_file_writable(report_out_path, "Executive Report (Output/QAStation_HFPA_Analysis_Report.xlsx)")

    # Remove legacy separate combine files if they exist
    for legacy_f in ["FTT_Combined_Validated.xlsx", "HFPA_Combined.xlsx"]:
        legacy_path = os.path.join(OUTPUT_DIR, legacy_f)
        if os.path.exists(legacy_path):
            try:
                os.remove(legacy_path)
                print(f"   [+] Removed legacy separate file: {legacy_f}")
            except Exception as e:
                print(f"   [-] Could not remove legacy file {legacy_f} (may be open in Office): {e}")

    # Prepare QAStation / Quality Tracking columns for export
    qa_export_cols = [
        "FTY", "InspDate", "Line", "Plant_Std", "Line_Std", "ShoeName", "InspQty",
        "FailQty", "Station", "Issues", "IssueQty", "IssueQty_Validated",
        "HFPA_Audit_Sample", "HFPA_Defect_Qty", "Defect_Variance", "Validation_Status"
    ]
    actual_qa_cols = [c for c in qa_export_cols if c in qa_validated.columns]
    qa_export_df = qa_validated[actual_qa_cols]

    # 1. Export 1 Single Master Combined Workbook (HFPA_FTT_Combined_Master.xlsx)
    print(f"\n[Exporting] 1. Saving 1 Unified Combined Master File to: {master_combined_path}...")
    with pd.ExcelWriter(master_combined_path, engine="openpyxl") as writer:
        qa_export_df.to_excel(writer, sheet_name="HFPA_Quality_Tracking", index=False)
        hfpa_combined.to_excel(writer, sheet_name="HFPA_Mes410", index=False)
        p1_df.to_excel(writer, sheet_name="Pivot1_Validation", index=False)
        all_top5.to_excel(writer, sheet_name="Top5_Models", index=False)
        p2_df.to_excel(writer, sheet_name="Top3_Defects", index=False)
        lot_comparison.to_excel(writer, sheet_name="Lot_Defect_Variance", index=False)

    print(f"   [+] Master combined workbook created with 6 sheets: HFPA_Quality_Tracking, HFPA_Mes410, Pivot1_Validation, Top5_Models, Top3_Defects, Lot_Defect_Variance")

    # 2. Export QAStation_HFPA_Analysis_Report.xlsx (Executive Summary Report)
    print(f"[Exporting] 2. Generating Executive Analysis Report: {report_out_path}...")
    with pd.ExcelWriter(report_out_path, engine="openpyxl") as writer:
        exec_summary = []
        for fty in FACTORIES + ["Grand Total"]:
            if fty == "Grand Total":
                sub_qa = qa_validated
                sub_hfpa = hfpa_combined
                sub_lot = lot_comparison
            else:
                sub_qa = qa_validated[qa_validated["FTY"] == fty]
                sub_hfpa = hfpa_combined[hfpa_combined["FTY"] == fty]
                sub_lot = lot_comparison[lot_comparison["FTY"] == fty]

            tot_sample = sub_qa.groupby(["Date_Std", "Plant_Std", "Line_Std"])["InspQty_Validated"].max().sum()
            qa_def = sub_qa["IssueQty_Raw"].sum()
            qa_val_def = sub_qa["IssueQty_Validated"].sum()
            hfpa_col = "Defect Q'ty (Total Defect Q'ty)" if "Defect Q'ty (Total Defect Q'ty)" in sub_hfpa.columns else "Defective Q'ty(pair)"
            hfpa_def = sub_hfpa[hfpa_col].sum()

            matched_exact = len(sub_lot[sub_lot["Validation_Status"] == "Matched_Exact"])
            total_lots = len(sub_lot)

            exec_summary.append({
                "Factory": fty,
                "Total Audited Sample (InspQty)": tot_sample,
                "QAStation Defects (Original)": qa_def,
                "QAStation Defects (Validated/HFPA)": qa_val_def,
                "HFPA Total Defects": hfpa_def,
                "DR% (Original)": (qa_def / tot_sample) if tot_sample > 0 else 0.0,
                "DR% (Validated)": (qa_val_def / tot_sample) if tot_sample > 0 else 0.0,
                "Matched Lots": matched_exact,
                "Total Lots": total_lots,
                "Match Rate": (matched_exact / total_lots) if total_lots > 0 else 0.0
            })
        exec_df = pd.DataFrame(exec_summary)
        exec_df.to_excel(writer, sheet_name="Executive_Summary", index=False)
        all_top5.to_excel(writer, sheet_name="Top5_Shoes", index=False)
        p2_df.to_excel(writer, sheet_name="Pivot2_Top3_Defects", index=False)
        p1_df.to_excel(writer, sheet_name="Pivot1_Validation", index=False)
        lot_comparison.to_excel(writer, sheet_name="Lot_Defect_Variance", index=False)

    # Apply styling & color chips to Executive Report
    wb = openpyxl.load_workbook(report_out_path)
    for sheetname in wb.sheetnames:
        ws = wb[sheetname]
        style_excel_sheet(ws, sheetname)

        if sheetname == "Pivot2_Top3_Defects":
            hex_col_idx = None
            for c in range(1, ws.max_column + 1):
                if ws.cell(1, c).value == "Color HEX":
                    hex_col_idx = c
                    break
            if hex_col_idx:
                for r in range(2, ws.max_row + 1):
                    hex_val = str(ws.cell(r, hex_col_idx).value or "").strip().replace("#", "")
                    if len(hex_val) == 6:
                        try:
                            fill = PatternFill(start_color=hex_val, end_color=hex_val, fill_type="solid")
                            ws.cell(r, 4).fill = fill
                            ws.cell(r, 4).font = Font(name="Calibri", size=11, bold=True, color="000000")
                        except Exception:
                            pass

    # Purge any redundant ColorMap_ sheets to keep workbook clean and optimal
    for s_name in list(wb.sheetnames):
        if s_name.startswith("ColorMap_"):
            del wb[s_name]

    wb.save(report_out_path)
    wb.close()

    # 3. Create Audit Trail Log as per QAStation_Tool_Agent_Spec.md Section 12
    audit_trail_path = os.path.join(logs_dir, "reconciliation_audit_trail.csv")
    audit_records = []
    # Identify records where values were adjusted
    adjusted_records = qa_validated[qa_validated["Validation_Status"] == "Different_Count"]
    for _, row in adjusted_records.iterrows():
        audit_records.append({
            "Factory": row.get("FTY", ""),
            "Date": row.get("Date_Std", ""),
            "Plant": row.get("Plant_Std", ""),
            "Line": row.get("Line_Std", ""),
            "Station": row.get("Station_Raw", ""),
            "Model": row.get("ShoeName_Std", ""),
            "Field": "IssueQty",
            "Original Value": row.get("IssueQty_Raw", ""),
            "MES410 Reference Total": row.get("HFPA_Defect_Qty", ""),
            "Corrected Value": row.get("IssueQty_Validated", ""),
            "Status": "Adjusted_To_MES410"
        })
    if audit_records:
        audit_df = pd.DataFrame(audit_records)
        audit_df.to_csv(audit_trail_path, index=False, encoding="utf-8-sig")
        print(f"   [+] Audit trail recorded ({len(audit_records)} adjustments) at: {audit_trail_path}")

    print("[+] Successfully generated all deliverables in Output/!")


def main():
    print("=" * 65)
    print("   QAStation & HFPA Data Processing Pipeline")
    print("=" * 65)

    # Detect active month and archive previous months
    ftt_sample_files = glob.glob(os.path.join(INPUT_FTT_DIR, "*.xlsx")) + glob.glob(os.path.join(INPUT_FTT_DIR, "*.xls"))
    hfpa_sample_files = glob.glob(os.path.join(INPUT_HFPA_DIR, "*.xlsx")) + glob.glob(os.path.join(INPUT_HFPA_DIR, "*.xls"))

    active_month = "Aug"
    if ftt_sample_files:
        active_month = detect_file_month(ftt_sample_files[0])
    elif hfpa_sample_files:
        active_month = detect_file_month(hfpa_sample_files[0])

    print(f"\n[*] XAC NHAN DU LIEU DAU VAO: THANG {active_month.upper()} ({active_month})")
    print(f"[*] Dang tu dong luu tru cac file thang khac vao thu muc Archive/...")
    auto_archive_previous_months(INPUT_FTT_DIR, active_month)
    auto_archive_previous_months(INPUT_HFPA_DIR, active_month)

    qa_combined = load_and_combine_ftt()
    hfpa_combined = load_and_combine_hfpa()

    qa_validated, hfpa_clean, lot_comparison = validate_and_correct_data(qa_combined, hfpa_combined)
    p1_df, all_top5, top5_dict = generate_pivot1_and_top5(qa_validated)
    p2_df = generate_pivot2_top_defects(qa_validated, top5_dict)

    export_deliverables(qa_validated, hfpa_clean, lot_comparison, p1_df, all_top5, p2_df)

    # Populate Database/HFPA_Template.xlsx automatically
    print(f"\n[6/7] Populating Database Template (HFPA_Template.xlsx) for Month: {active_month}...")
    try:
        from populate_database import populate_hfpa_database
        populate_hfpa_database(target_month=active_month, p2_df=p2_df)
    except Exception as e:
        print(f"[-] Warning: Failed to populate database template: {e}")

    # Generate PowerPoint Performance Presentation (HFPA_Template.pptx) automatically
    print(f"\n[7/7] Generating PowerPoint Performance Presentation (HFPA_Template.pptx)...")
    try:
        from process_pptx import generate_hfpa_presentation
        generate_hfpa_presentation(target_month=active_month)
    except Exception as e:
        print(f"[-] Warning: Failed to generate PowerPoint presentation: {e}")

    print("\n" + "=" * 65)
    print(f"   Hoan tat xu ly du lieu Thang {active_month.upper()}! Tat ca bao cao da san sang trong Output/")
    print("=" * 65)


if __name__ == "__main__":
    main()
