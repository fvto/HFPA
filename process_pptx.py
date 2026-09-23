"""
Module: process_pptx.py
Automates mapping of HFPA & QAStation analysis data from Output/ into Database/HFPA_Template.pptx.
Implements:
1. Header & metric cards updates with pixel-perfect typography (Calibri, Calibri Light, Arial).
2. Dynamic native PowerPoint Defect Color Legend Table replacing static Picture 2 (matches Image 2 & process_ftt.py).
3. Interactive hyperlinks on navigation buttons, site icons, and metric cards.
4. Updating all 12 charts in the 4x3 layout:
   - Column 1: HFPA Analysis (Combination Stacked Column + Line chart across months).
   - Column 2: Top 5 models (Clustered Column chart with DR%, 7.5pt bold labels).
   - Column 3: Top defect of top models (Stacked Column chart with exact defect color chips,
               following process_ftt.py ranking and Color_template.xlsx styling, 7pt contrasting labels).
5. Dual saving to Database/HFPA_Template.pptx and Output/HFPA_Performance_Report.pptx.
"""

import os
import re
import sys
import shutil
import openpyxl
import pandas as pd
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml
from openpyxl.utils import get_column_letter
from PIL import Image, ImageDraw, ImageFont

from file_utils import ensure_file_writable, detect_file_month

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATABASE_DIR = os.path.join(BASE_DIR, "Database")
OUTPUT_DIR = os.path.join(BASE_DIR, "Output")
TEMPLATE_PPTX = os.path.join(DATABASE_DIR, "HFPA_Template.pptx")
BACKUP_PPTX = os.path.join(DATABASE_DIR, "HFPA_Template_Backup.pptx")
OUTPUT_PPTX = os.path.join(OUTPUT_DIR, "HFPA_Performance_Report.pptx")
UPDATED_EXCEL = os.path.join(OUTPUT_DIR, "HFPA_Template_Updated.xlsx")
ANALYSIS_EXCEL = os.path.join(OUTPUT_DIR, "QAStation_HFPA_Analysis_Report.xlsx")
MASTER_EXCEL = os.path.join(OUTPUT_DIR, "HFPA_FTT_Combined_Master.xlsx")
COLOR_TEMPLATE = os.path.join(BASE_DIR, "Color_template.xlsx")

FACTORIES = ["VH", "VH2", "JV", "JV2"]
MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

# Fallback palette matching process_ftt.py
FALLBACK_PALETTE = [
    "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5",
    "70AD47", "264478", "9E480E", "636363", "997300",
    "255E91", "43682B", "698ED0", "F1975A", "B7B7B7",
]

# Standard legend defects ordering (matches template layout / Image 2)
STANDARD_DEFECT_ORDER = [
    "Rocking",
    "Cleanness",
    "Thread end",
    "Over buffing",
    "Over cement",
    "Collar shape",
    "Bottom Cosmetic",
    "Components inconsistent",
    "Midsole/Outsole to upper bond gap",
    "Sole attachment",
    "Toe off center",
]

# Chart mapping per factory
SITE_CHART_MAP = {
    "VH": {
        "analysis": "Chart 51",
        "top5": "Chart 194",
        "top_defect": "Chart 63",
        "pentagon": "Pentagon 39",
        "dr_rect": "Rectangle 176",
        "excel_rows": {"prod": 2, "defect": 9, "hfpa": 16, "top5_header": 21, "top5_start": 22, "top5_end": 26},
    },
    "VH2": {
        "analysis": "Chart 52",
        "top5": "Chart 195",
        "top_defect": "Chart 64",
        "pentagon": "Pentagon 151",
        "dr_rect": "Rectangle 178",
        "excel_rows": {"prod": 3, "defect": 10, "hfpa": 17, "top5_header": 27, "top5_start": 28, "top5_end": 32},
    },
    "JV": {
        "analysis": "Chart 53",
        "top5": "Chart 196",
        "top_defect": "Chart 65",
        "pentagon": "Pentagon 152",
        "dr_rect": "Rectangle 180",
        "excel_rows": {"prod": 4, "defect": 11, "hfpa": 18, "top5_header": 33, "top5_start": 34, "top5_end": 38},
    },
    "JV2": {
        "analysis": "Chart 54",
        "top5": "Chart 197",
        "top_defect": "Chart 66",
        "pentagon": "Pentagon 156",
        "dr_rect": "Rectangle 181",
        "excel_rows": {"prod": 5, "defect": 12, "hfpa": 19, "top5_header": 39, "top5_start": 40, "top5_end": 44},
    },
}


# ======================================================================
# 1. COLOR MAPPING HELPERS
# ======================================================================
def normalize_defect_name(name):
    """Normalize a defect name for matching: lowercase, alphanumeric only."""
    return re.sub(r"[\s_\-]+", "", str(name).strip().lower())


def get_label_font_color(hex_color):
    """Calculate contrasting font color (white or black) based on background luminance."""
    hex_color = hex_color.lstrip("#")
    try:
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        brightness = (r * 299 + g * 587 + b * 114) / 1000
        return "000000" if brightness >= 135 else "FFFFFF"
    except Exception:
        return "FFFFFF"


def load_defect_color_map(color_file_path=COLOR_TEMPLATE):
    """Load defect -> HEX color mapping from Color_template.xlsx."""
    color_map = {}
    if not os.path.exists(color_file_path):
        print(f"[!] Warning: Color template file not found: {color_file_path}")
        return color_map

    try:
        wb = openpyxl.load_workbook(color_file_path, data_only=True)
        # Scan sheets
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            for r in range(1, ws.max_row + 1):
                for c in range(1, ws.max_column - 1):
                    val = ws.cell(r, c).value
                    if val and "defect" in str(val).lower() and "type" in str(val).lower():
                        header_r = r
                        defect_col = c
                        hex_col = None
                        r_col = g_col = b_col = None
                        for sc in range(c + 1, ws.max_column + 1):
                            sval = str(ws.cell(r, sc).value or "").strip().lower()
                            if sval == "hex":
                                hex_col = sc
                            elif sval in ("red", "r"):
                                r_col = sc
                            elif sval in ("green", "g"):
                                g_col = sc
                            elif sval in ("blue", "b"):
                                b_col = sc

                        for dr in range(header_r + 1, ws.max_row + 1):
                            dname = ws.cell(dr, defect_col).value
                            if not dname or not str(dname).strip():
                                continue
                            norm_name = normalize_defect_name(dname)
                            hex_val = None
                            if hex_col:
                                raw_hex = str(ws.cell(dr, hex_col).value or "").strip().lstrip("#")
                                if re.fullmatch(r"[0-9A-Fa-f]{6}", raw_hex):
                                    hex_val = raw_hex.upper()
                            if not hex_val and r_col and g_col and b_col:
                                try:
                                    rv = int(float(ws.cell(dr, r_col).value or 0))
                                    gv = int(float(ws.cell(dr, g_col).value or 0))
                                    bv = int(float(ws.cell(dr, b_col).value or 0))
                                    hex_val = f"{rv:02X}{gv:02X}{bv:02X}"
                                except Exception:
                                    pass
                            if hex_val:
                                color_map[norm_name] = hex_val
                        break
        wb.close()
    except Exception as e:
        print(f"[!] Error reading color mapping: {e}")

    print(f"   [+] Loaded {len(color_map)} defect color mappings from {os.path.basename(color_file_path)}")
    return color_map


def get_defect_color(defect_name, color_map, fallback_cache):
    """Retrieve defect HEX color with normalized lookup and dynamic fallback."""
    norm = normalize_defect_name(defect_name)
    if norm in color_map:
        return color_map[norm]
    # Partial matching
    for k, v in color_map.items():
        if k in norm or norm in k:
            return v
    if norm in fallback_cache:
        return fallback_cache[norm]
    color = FALLBACK_PALETTE[len(fallback_cache) % len(FALLBACK_PALETTE)]
    fallback_cache[norm] = color
    return color


# ======================================================================
# 2. DATA EXTRACTION FROM OUTPUT/
# ======================================================================
def extract_pipeline_data(target_month=None):
    """
    Extract all required data from Output/ workbooks.
    """
    fallback_excel = os.path.join(OUTPUT_DIR, "HFPA_Template_Updated_New.xlsx")
    if os.path.exists(fallback_excel):
        excel_path = fallback_excel
    elif os.path.exists(UPDATED_EXCEL):
        excel_path = UPDATED_EXCEL
    else:
        excel_path = os.path.join(DATABASE_DIR, "HFPA_Template.xlsx")
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Source Excel not found: {excel_path}")

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    if "HFPA" not in wb.sheetnames:
        raise ValueError(f"Sheet 'HFPA' not found in {excel_path}")
    ws = wb["HFPA"]

    # Detect active months in row 1
    month_cols = {}
    for c in range(3, ws.max_column + 1):
        val = str(ws.cell(1, c).value or "").strip()
        for m in MONTH_NAMES:
            if m.lower() == val.lower():
                month_cols[m] = c
                break

    active_months = []
    for m in MONTH_NAMES:
        if m in month_cols:
            col_idx = month_cols[m]
            has_data = any((ws.cell(r, col_idx).value or 0) > 0 for r in [2, 3, 4, 5])
            if has_data:
                active_months.append(m)

    if not active_months:
        active_months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug"]

    if not target_month:
        target_month = active_months[-1]
    elif target_month not in active_months and target_month in month_cols:
        active_months.append(target_month)

    # Filter active_months to only months up to target_month
    if target_month in MONTH_NAMES:
        target_idx = MONTH_NAMES.index(target_month)
        active_months = [m for m in active_months if MONTH_NAMES.index(m) <= target_idx]

    target_year = "2026"
    print(f"   [+] Presentation Target Month: {target_month} {target_year} (Active range: {active_months[0]}..{active_months[-1]})")

    # Monthly production & defect metrics for Chart 1
    monthly_metrics = {}
    for fty, cfg in SITE_CHART_MAP.items():
        prod_row = cfg["excel_rows"]["prod"]
        def_row = cfg["excel_rows"]["defect"]
        prod_vals = []
        def_vals = []
        hfpa_vals = []

        for m in active_months:
            c = month_cols.get(m, 10)
            pv = float(ws.cell(prod_row, c).value or 0)
            dv = float(ws.cell(def_row, c).value or 0)
            prod_vals.append(pv)
            def_vals.append(dv)
            hfpa_vals.append(1.0 - (dv / pv) if pv > 0 else 0.0)

        monthly_metrics[fty] = {
            "prod": prod_vals,
            "def": def_vals,
            "hfpa": hfpa_vals,
        }

    # Target month summary numbers
    target_col = month_cols.get(target_month, month_cols.get(active_months[-1], 10))
    site_summaries = {}
    total_prod = 0.0
    total_def = 0.0

    for fty, cfg in SITE_CHART_MAP.items():
        p_row = cfg["excel_rows"]["prod"]
        d_row = cfg["excel_rows"]["defect"]
        pv = float(ws.cell(p_row, target_col).value or 0)
        dv = float(ws.cell(d_row, target_col).value or 0)
        dr = dv / pv if pv > 0 else 0.0
        site_summaries[fty] = {"prod": int(pv), "def": int(dv), "dr": dr}
        total_prod += pv
        total_def += dv

    grand_total = {
        "prod": int(total_prod),
        "def": int(total_def),
        "dr": (total_def / total_prod) if total_prod > 0 else 0.0,
    }

    # Top 5 Models & Top Defects Matrix for each factory
    top5_data = {}
    top_defects_matrix = {}

    for fty, cfg in SITE_CHART_MAP.items():
        hdr_row = cfg["excel_rows"]["top5_header"]
        start_row = cfg["excel_rows"]["top5_start"]
        end_row = cfg["excel_rows"]["top5_end"]

        models = []
        top5_list = []
        for r in range(start_row, end_row + 1):
            m_name = ws.cell(r, 4).value
            if not m_name:
                continue
            m_name = str(m_name).strip()
            models.append(m_name)
            insp = float(ws.cell(r, 5).value or 0)
            fail = float(ws.cell(r, 6).value or 0)
            dr_val = ws.cell(r, 7).value
            if isinstance(dr_val, (int, float)):
                dr = float(dr_val)
            else:
                dr = (fail / insp) if insp > 0 else 0.0
            top5_list.append({
                "model": m_name,
                "insp": int(insp),
                "def": int(fail),
                "dr": dr,
            })
        top5_data[fty] = top5_list

        defect_cols = {}
        for c in range(13, ws.max_column + 1):
            val = ws.cell(hdr_row, c).value
            if not val or not str(val).strip():
                break  # Stop at first empty column (boundary between defect cols and site label cols)
            col_name = str(val).strip()
            # Skip non-defect headers that leak from adjacent sections
            if col_name.startswith("Site") or col_name == "Model":
                break  # Reached the next section boundary
            defect_cols[col_name] = c

        defect_values = {}
        defects_in_order = list(defect_cols.keys())

        for d_name in defects_in_order:
            c = defect_cols[d_name]
            vals = []
            for r in range(start_row, start_row + len(models)):
                v = ws.cell(r, c).value
                if isinstance(v, (int, float)) and v > 0:
                    vals.append(float(v))
                else:
                    vals.append(None)
            defect_values[d_name] = vals

        top_defects_matrix[fty] = {
            "models": models,
            "defects": defects_in_order,
            "values": defect_values,
            "start_row": start_row,
            "end_row": end_row,
            "defect_cols": defect_cols,
        }

    wb.close()

    return {
        "target_month": target_month,
        "target_year": target_year,
        "active_months": active_months,
        "monthly_metrics": monthly_metrics,
        "site_summaries": site_summaries,
        "grand_total": grand_total,
        "top5_data": top5_data,
        "top_defects_matrix": top_defects_matrix,
    }


# ======================================================================
# 3. TYPOGRAPHY & SHAPE TEXT UPDATERS
# ======================================================================
def set_run_font(run, name="Calibri", size_pt=10, bold=False, color_rgb=(0, 0, 0)):
    """Apply font family, size in points, bold style, and RGB color to a text run."""
    run.font.name = name
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    if color_rgb:
        run.font.color.rgb = RGBColor(*color_rgb)


def update_slide_metrics_and_links(slide, data):
    """Update text frames, summary badges, site cards, and hyperlinks in Slide 1 with crisp typography."""
    target_month = data["target_month"]
    target_year = data["target_year"]
    site_sums = data["site_summaries"]
    g_total = data["grand_total"]

    for s in slide.shapes:
        # 1. Title: Rectangle 100
        if s.name == "Rectangle 100":
            tf = s.text_frame
            tf.word_wrap = True
            tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
            p = tf.paragraphs[0]
            p.text = ""
            r0 = p.add_run()
            r0.text = "HFPA Performance "
            set_run_font(r0, name="Calibri Light", size_pt=26, bold=True, color_rgb=(0, 32, 96))
            r1 = p.add_run()
            r1.text = f"{target_month}, {target_year}"
            set_run_font(r1, name="Arial", size_pt=14, bold=True, color_rgb=(237, 125, 49))
            print(f"   [+] Updated Title: HFPA Performance {target_month}, {target_year}")

        # 2. Hyperlinks on navigation buttons
        if s.name == "Rectangle 104":  # 'Analysis by Site'
            s.click_action.hyperlink.address = r"../Output/QAStation_HFPA_Analysis_Report.xlsx"
            print("   [+] Set hyperlink: Rectangle 104 -> Output/QAStation_HFPA_Analysis_Report.xlsx")

        if s.name in ("Rectangle 183", "Rectangle 184"):  # 'HFPA' / 'High Frequency Product Audit'
            s.click_action.hyperlink.address = r"../Output/HFPA_FTT_Combined_Master.xlsx"
            print(f"   [+] Set hyperlink: {s.name} -> Output/HFPA_FTT_Combined_Master.xlsx")

        # 3. Update Group 14 (Total pair produced)
        if s.name == "组合 14":
            for sub in s.shapes:
                if sub.name == "Oval 109":
                    tf = sub.text_frame
                    tf.word_wrap = False
                    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
                    tf.clear()
                    p0 = tf.paragraphs[0]
                    p0.alignment = PP_ALIGN.CENTER
                    r0 = p0.add_run()
                    r0.text = "Total pair produced"
                    set_run_font(r0, name="Calibri", size_pt=8.5, bold=True, color_rgb=(255, 255, 255))

                    p1 = tf.add_paragraph()
                    p1.alignment = PP_ALIGN.CENTER
                    r1 = p1.add_run()
                    r1.text = f"~{round(g_total['prod'] / 1000):,}K"
                    set_run_font(r1, name="Calibri", size_pt=16, bold=True, color_rgb=(255, 255, 255))
                    sub.click_action.hyperlink.address = r"../Output/HFPA_Template_Updated.xlsx"

                if sub.shape_type == pptx.enum.shapes.MSO_SHAPE_TYPE.GROUP:
                    factories_prod = {
                        "Rectangle 116": f"VH: {site_sums['VH']['prod']:,} prs",
                        "Rectangle 159": f"VH2: {site_sums['VH2']['prod']:,} prs",
                        "Rectangle 160": f"JV: {site_sums['JV']['prod']:,} prs",
                        "Rectangle 161": f"JV2: {site_sums['JV2']['prod']:,} prs",
                    }
                    for sub2 in sub.shapes:
                        if sub2.name in factories_prod:
                            tf = sub2.text_frame
                            tf.word_wrap = False
                            tf.margin_left = Pt(10)
                            tf.margin_right = Pt(2)
                            tf.margin_top = tf.margin_bottom = 0
                            p = tf.paragraphs[0]
                            p.text = factories_prod[sub2.name]
                            if p.runs:
                                set_run_font(p.runs[0], name="Calibri Light", size_pt=11.5, bold=False, color_rgb=(0, 0, 0))
            print(f"   [+] Updated Group 14: Total pair produced (~{round(g_total['prod']/1000):,}K prs)")

        # 4. Update Group 162 (Defective)
        if s.name == "Group 162":
            for sub in s.shapes:
                if sub.name == "Oval 165":
                    tf = sub.text_frame
                    tf.word_wrap = False
                    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
                    tf.clear()
                    p0 = tf.paragraphs[0]
                    p0.alignment = PP_ALIGN.CENTER
                    r0 = p0.add_run()
                    r0.text = "Defective"
                    set_run_font(r0, name="Calibri", size_pt=8.5, bold=True, color_rgb=(255, 255, 255))

                    p1 = tf.add_paragraph()
                    p1.alignment = PP_ALIGN.CENTER
                    r1 = p1.add_run()
                    r1.text = f"~{round(g_total['def'] / 1000):,}K"
                    set_run_font(r1, name="Calibri", size_pt=16, bold=True, color_rgb=(255, 255, 255))
                    sub.click_action.hyperlink.address = r"../Output/HFPA_Template_Updated.xlsx"

                if sub.shape_type == pptx.enum.shapes.MSO_SHAPE_TYPE.GROUP:
                    factories_def = {
                        "Rectangle 167": f"VH: {site_sums['VH']['def']:,} prs",
                        "Rectangle 168": f"VH2: {site_sums['VH2']['def']:,} prs",
                        "Rectangle 169": f"JV: {site_sums['JV']['def']:,} prs",
                        "Rectangle 170": f"JV2: {site_sums['JV2']['def']:,} prs",
                    }
                    for sub2 in sub.shapes:
                        if sub2.name in factories_def:
                            tf = sub2.text_frame
                            tf.word_wrap = False
                            tf.margin_left = Pt(10)
                            tf.margin_right = Pt(2)
                            tf.margin_top = tf.margin_bottom = 0
                            p = tf.paragraphs[0]
                            p.text = factories_def[sub2.name]
                            if p.runs:
                                set_run_font(p.runs[0], name="Calibri Light", size_pt=11.5, bold=False, color_rgb=(0, 0, 0))
            print(f"   [+] Updated Group 162: Defective (~{round(g_total['def']/1000):,}K prs)")

        # 5. Site Pentagons & DR% Rectangles
        for fty, cfg in SITE_CHART_MAP.items():
            if s.name == cfg["pentagon"]:
                tf = s.text_frame
                tf.word_wrap = False
                tf.clear()
                p0 = tf.paragraphs[0]
                p0.alignment = PP_ALIGN.CENTER
                r0 = p0.add_run()
                r0.text = fty
                set_run_font(r0, name="Calibri", size_pt=12, bold=True, color_rgb=(0, 0, 0))
                s.click_action.hyperlink.address = r"../Output/HFPA_Template_Updated.xlsx"

            elif s.name == cfg["dr_rect"]:
                dr_pct = site_sums[fty]["dr"]
                tf = s.text_frame
                tf.word_wrap = False
                tf.margin_left = tf.margin_right = Pt(1)
                tf.margin_top = tf.margin_bottom = Pt(1)
                tf.clear()

                p0 = tf.paragraphs[0]
                p0.alignment = PP_ALIGN.CENTER
                p0.space_after = Pt(1)
                r0 = p0.add_run()
                r0.text = "HFPA DR:"
                set_run_font(r0, name="Calibri", size_pt=8.5, bold=True, color_rgb=(0, 0, 0))

                p1 = tf.add_paragraph()
                p1.alignment = PP_ALIGN.CENTER
                p1.space_before = Pt(0)
                r1 = p1.add_run()
                r1.text = f"{dr_pct:.2%}"
                set_run_font(r1, name="Calibri", size_pt=9.5, bold=True, color_rgb=(0, 0, 0))
                s.click_action.hyperlink.address = r"../Output/QAStation_HFPA_Analysis_Report.xlsx"

        # 6. Side note: Rectangle 55
        if s.name == "Rectangle 55":
            tf = s.text_frame
            tf.word_wrap = True
            for p in tf.paragraphs:
                for r in p.runs:
                    r.font.name = "Calibri"
                    r.font.size = Pt(9.5)


# ======================================================================
# 4. DYNAMIC DEFECT COLOR LEGEND TABLE
# ======================================================================
def create_defect_color_legend_table(slide, data, color_map, fallback_cache):
    """
    Render Defect Color Legend as a high-resolution, fixed-height Picture shape.
    Dynamically collects ALL unique defect types from chart data so no defect is missing.
    Using a Picture shape ensures the legend NEVER expands downward in WPS Office or PowerPoint,
    guaranteeing a clean 190,000 EMU safety gap above Chart 63.
    """
    # 1. Remove old Picture 2 or Table Defect Legend
    for s in list(slide.shapes):
        if s.name in ("Picture 2", "Table Defect Legend"):
            sp = s._element
            sp.getparent().remove(sp)

    # 2. Dynamically collect ALL unique defects from the actual chart data
    #    Validate each defect against color_map to ensure only real defects are included.
    #    This is the safety net — even if data extraction lets something through,
    #    the legend will only show defects that have a color in Color_template.xlsx.
    all_defects_ordered = []
    seen = set()
    if "top_defects_matrix" in data:
        for fty in data["top_defects_matrix"]:
            matrix = data["top_defects_matrix"][fty]
            for d_name in matrix.get("defects", []):
                if d_name not in seen:
                    # Validate: only include if it has a color mapping (is a real defect)
                    norm = normalize_defect_name(d_name)
                    is_real_defect = norm in color_map or any(k in norm or norm in k for k in color_map)
                    if is_real_defect:
                        all_defects_ordered.append(d_name)
                    else:
                        print(f"   [!] Legend: skipping non-defect entry '{d_name}' (no color mapping)")
                    seen.add(d_name)

    # Fallback if no data available
    if not all_defects_ordered:
        all_defects_ordered = [
            "Rocking", "Cleanness", "Thread end", "Over buffing", "Over cement",
            "Collar shape", "Bottom Cosmetic", "Components inconsistent",
            "Midsole/Outsole to upper bond gap",
        ]

    # Arrange into 2-column grid (left column first, then right column)
    total = len(all_defects_ordered)
    rows = (total + 1) // 2  # ceil division
    if rows < 5:
        rows = 5  # minimum 5 rows for consistent layout

    left_defects = all_defects_ordered[:rows]
    right_defects = all_defects_ordered[rows:]

    print(f"   [i] Legend defects ({total} total): Left={left_defects}, Right={right_defects}")

    left = 8588371
    top = 580000
    width = 2855916
    height = 760000  # Ends at 1,340,000 (leaves 190,000 EMU gap before Chart 63 at 1,529,828)

    # 3. Render high-resolution 4x image
    row_h_px = 68  # fixed row height in pixels
    w_px = 1250
    h_px = rows * row_h_px
    img = Image.new("RGB", (w_px, h_px), (255, 255, 255))
    draw = ImageDraw.Draw(img)

    font = None
    for fp in ["C:/Windows/Fonts/calibri.ttf", "C:/Windows/Fonts/arial.ttf", "calibri.ttf", "arial.ttf"]:
        if os.path.exists(fp):
            try:
                font = ImageFont.truetype(fp, 28)
                break
            except Exception:
                pass
    if font is None:
        font = ImageFont.load_default()

    col_xs = [0, 100, 490, 590, 1250]

    for r in range(rows):
        y0 = r * row_h_px
        y1 = (r + 1) * row_h_px

        # Left chip
        if r < len(left_defects):
            dname1 = left_defects[r]
            hex1 = get_defect_color(dname1, color_map, fallback_cache)
            c1 = (int(hex1[0:2], 16), int(hex1[2:4], 16), int(hex1[4:6], 16))
            draw.rectangle([col_xs[0], y0, col_xs[1], y1], fill=c1, outline=(0, 0, 0), width=1)

            # Left text
            draw.rectangle([col_xs[1], y0, col_xs[2], y1], fill=(255, 255, 255), outline=(0, 0, 0), width=1)
            bbox1 = font.getbbox(dname1)
            th1 = bbox1[3] - bbox1[1]
            ty1 = y0 + (y1 - y0 - th1) // 2 - bbox1[1]
            draw.text((col_xs[1] + 12, ty1), dname1, fill=(0, 0, 0), font=font)
        else:
            draw.rectangle([col_xs[0], y0, col_xs[1], y1], fill=(255, 255, 255), outline=(0, 0, 0), width=1)
            draw.rectangle([col_xs[1], y0, col_xs[2], y1], fill=(255, 255, 255), outline=(0, 0, 0), width=1)

        # Right chip
        if r < len(right_defects):
            dname2 = right_defects[r]
            hex2 = get_defect_color(dname2, color_map, fallback_cache)
            c2 = (int(hex2[0:2], 16), int(hex2[2:4], 16), int(hex2[4:6], 16))
            draw.rectangle([col_xs[2], y0, col_xs[3], y1], fill=c2, outline=(0, 0, 0), width=1)

            # Right text
            draw.rectangle([col_xs[3], y0, col_xs[4] - 1, y1], fill=(255, 255, 255), outline=(0, 0, 0), width=1)
            bbox2 = font.getbbox(dname2)
            th2 = bbox2[3] - bbox2[1]
            ty2 = y0 + (y1 - y0 - th2) // 2 - bbox2[1]
            draw.text((col_xs[3] + 12, ty2), dname2, fill=(0, 0, 0), font=font)
        else:
            # Empty white cells with border
            draw.rectangle([col_xs[2], y0, col_xs[3], y1], fill=(255, 255, 255), outline=(0, 0, 0), width=1)
            draw.rectangle([col_xs[3], y0, col_xs[4] - 1, y1], fill=(255, 255, 255), outline=(0, 0, 0), width=1)

    # Outer border
    draw.rectangle([0, 0, w_px - 1, h_px - 1], outline=(0, 0, 0), width=1)

    rendered_img_path = os.path.join(OUTPUT_DIR, "defect_legend_rendered.png")
    img.save(rendered_img_path, "PNG")

    # 4. Insert as Picture shape
    pic = slide.shapes.add_picture(rendered_img_path, left, top, width, height)
    pic.name = "Picture 2"

    print(f"   [+] Generated pixel-perfect Defect Color Legend Picture ({total} defects, {rows} rows) top={top}, height={height}, bottom={top+height}")


# ======================================================================
# 5. XML CHART BUILDERS & UPDATERS (WITH ACCURATE FONT CONTROLS)
# ======================================================================
def update_chart_relationships(chart_shape):
    """Update oleObject relationship rId1 to point to Output/HFPA_Template_Updated.xlsx."""
    part = chart_shape.chart.part
    if "rId1" in part.rels:
        rel = part.rels["rId1"]
        if rel.is_external:
            rel._target = r"../Output/HFPA_Template_Updated.xlsx"


def update_hfpa_analysis_chart(chart_shape, active_months, metrics_dict, prod_row, def_row, hfpa_row):
    """
    Update Column 1: 'HFPA Analysis' Combination Chart.
    Plot 0: BarChart (Total pair produced & Defect)
    Plot 1: LineChart (HFPA%)
    """
    chart_space = chart_shape.chart._chartSpace
    plot_area = chart_space.find(qn('c:chart')).find(qn('c:plotArea'))

    bar_chart = plot_area.find(qn('c:barChart'))
    line_chart = plot_area.find(qn('c:lineChart'))

    start_col_letter = "C"
    end_col_letter = get_column_letter(2 + len(active_months))
    cat_formula = f"'Current month'!${start_col_letter}$1:${end_col_letter}$1"

    cat_pts = "".join([f'<c:pt idx="{i}"><c:v>{m}</c:v></c:pt>' for i, m in enumerate(active_months)])
    cat_xml = f"""
    <c:cat xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
      <c:strRef>
        <c:f>{cat_formula}</c:f>
        <c:strCache>
          <c:ptCount val="{len(active_months)}"/>
          {cat_pts}
        </c:strCache>
      </c:strRef>
    </c:cat>
    """

    def update_series_vals(ser_elem, title, formula_cell, formula_range, values, num_fmt, dlbl_pos="ctr", dlbl_sz="700"):
        tx = ser_elem.find(qn('c:tx'))
        if tx is not None:
            f = tx.find(qn('c:strRef')).find(qn('c:f'))
            if f is not None:
                f.text = formula_cell
            v = tx.find(qn('c:strRef')).find(qn('c:strCache')).find(qn('c:pt')).find(qn('c:v'))
            if v is not None:
                v.text = title

        old_cat = ser_elem.find(qn('c:cat'))
        if old_cat is not None:
            ser_elem.remove(old_cat)
        ser_elem.append(parse_xml(cat_xml))

        val_pts = "".join([f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i, v in enumerate(values)])
        val_xml = f"""
        <c:val xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:numRef>
            <c:f>{formula_range}</c:f>
            <c:numCache>
              <c:formatCode>{num_fmt}</c:formatCode>
              <c:ptCount val="{len(values)}"/>
              {val_pts}
            </c:numCache>
          </c:numRef>
        </c:val>
        """
        old_val = ser_elem.find(qn('c:val'))
        if old_val is not None:
            ser_elem.remove(old_val)
        ser_elem.append(parse_xml(val_xml))

        # Ensure dLbls typography
        dLbls = ser_elem.find(qn('c:dLbls'))
        if dLbls is not None:
            pos = dLbls.find(qn('c:dLblPos'))
            if pos is not None:
                pos.set('val', dlbl_pos)
            txPr = dLbls.find(qn('c:txPr'))
            if txPr is not None:
                p = txPr.find(qn('a:p'))
                if p is not None and p.find(qn('a:pPr')) is not None and p.find(qn('a:pPr')).find(qn('a:defRPr')) is not None:
                    p.find(qn('a:pPr')).find(qn('a:defRPr')).set('sz', dlbl_sz)

    bar_series = bar_chart.findall(qn('c:ser'))
    if len(bar_series) >= 2:
        update_series_vals(
            bar_series[0],
            "Total pair produced",
            f"'Current month'!$B${prod_row}",
            f"'Current month'!${start_col_letter}${prod_row}:${end_col_letter}${prod_row}",
            metrics_dict["prod"],
            "#,##0",
            dlbl_pos="ctr",
            dlbl_sz="700",
        )
        update_series_vals(
            bar_series[1],
            "Defect",
            f"'Current month'!$B${def_row}",
            f"'Current month'!${start_col_letter}${def_row}:${end_col_letter}${def_row}",
            metrics_dict["def"],
            "#,##0",
            dlbl_pos="inBase",
            dlbl_sz="700",
        )

    if line_chart is not None:
        line_series = line_chart.findall(qn('c:ser'))
        if line_series:
            update_series_vals(
                line_series[0],
                "HFPA%",
                f"'Current month'!$B${hfpa_row}",
                f"'Current month'!${start_col_letter}${hfpa_row}:${end_col_letter}${hfpa_row}",
                metrics_dict["hfpa"],
                "0.00%",
                dlbl_pos="t",
                dlbl_sz="750",
            )

    update_chart_relationships(chart_shape)


def update_top5_models_chart(chart_shape, top5_list, start_row, end_row):
    """
    Update Column 2: 'Top 5 models' Clustered Column Chart.
    Categories: Top 5 models (6pt font)
    Series 0: HFPA DR% with 7.5pt bold labels outside end.
    """
    chart_space = chart_shape.chart._chartSpace
    plot_area = chart_space.find(qn('c:chart')).find(qn('c:plotArea'))
    bar_chart = plot_area.find(qn('c:barChart'))

    models = [item["model"] for item in top5_list]
    dr_vals = [item["dr"] for item in top5_list]

    cat_pts = "".join([f'<c:pt idx="{i}"><c:v>{m}</c:v></c:pt>' for i, m in enumerate(models)])
    val_pts = "".join([f'<c:pt idx="{i}"><c:v>{v:.6f}</c:v></c:pt>' for i, v in enumerate(dr_vals)])

    cat_f = f"'Current month'!$D${start_row}:$D${end_row}"
    val_f = f"'Current month'!$G${start_row}:$G${end_row}"

    cat_xml = f"""
    <c:cat xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
      <c:strRef>
        <c:f>{cat_f}</c:f>
        <c:strCache>
          <c:ptCount val="{len(models)}"/>
          {cat_pts}
        </c:strCache>
      </c:strRef>
    </c:cat>
    """

    val_xml = f"""
    <c:val xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
      <c:numRef>
        <c:f>{val_f}</c:f>
        <c:numCache>
          <c:formatCode>0.00%</c:formatCode>
          <c:ptCount val="{len(dr_vals)}"/>
          {val_pts}
        </c:numCache>
      </c:numRef>
    </c:val>
    """

    series_list = bar_chart.findall(qn('c:ser'))
    if series_list:
        ser0 = series_list[0]
        old_cat = ser0.find(qn('c:cat'))
        if old_cat is not None:
            ser0.remove(old_cat)
        ser0.append(parse_xml(cat_xml))

        old_val = ser0.find(qn('c:val'))
        if old_val is not None:
            ser0.remove(old_val)
        ser0.append(parse_xml(val_xml))

        # Typography on dLbls (7.5pt bold black)
        old_dlbls = ser0.find(qn('c:dLbls'))
        if old_dlbls is not None:
            ser0.remove(old_dlbls)
        dlbls_xml = """
        <c:dLbls xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>
          <c:txPr>
            <a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" lIns="19050" tIns="9525" rIns="19050" bIns="9525" anchor="ctr" anchorCtr="1"/>
            <a:lstStyle/>
            <a:p>
              <a:pPr>
                <a:defRPr lang="en-US" sz="750" b="1" i="0" baseline="0">
                  <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
                  <a:latin typeface="Calibri"/>
                </a:defRPr>
              </a:pPr>
            </a:p>
          </c:txPr>
          <c:numFmt formatCode="0.00%" sourceLinked="0"/>
          <c:dLblPos val="outEnd"/>
          <c:showVal val="1"/>
          <c:showLegendKey val="0"/>
          <c:showCatName val="0"/>
          <c:showSerName val="0"/>
          <c:showPercent val="0"/>
          <c:showBubbleSize val="0"/>
          <c:showLeaderLines val="0"/>
        </c:dLbls>
        """
        ser0.append(parse_xml(dlbls_xml))

    # Category axis font size = 6pt
    catAx = plot_area.find(qn('c:catAx'))
    if catAx is not None:
        txPr = catAx.find(qn('c:txPr'))
        if txPr is not None:
            p = txPr.find(qn('a:p'))
            if p is not None and p.find(qn('a:pPr')) is not None and p.find(qn('a:pPr')).find(qn('a:defRPr')) is not None:
                p.find(qn('a:pPr')).find(qn('a:defRPr')).set('sz', '600')

    update_chart_relationships(chart_shape)


def update_top_defects_chart(chart_shape, defect_matrix, color_map, fallback_cache):
    """
    Update Column 3: 'Top defect of top models' Stacked Column Chart.
    Strictly adheres to process_ftt.py ranking and Color_template.xlsx HEX styling.
    Applies clean 7pt contrasting typography on stacked labels.
    """
    chart_space = chart_shape.chart._chartSpace
    plot_area = chart_space.find(qn('c:chart')).find(qn('c:plotArea'))
    bar_chart = plot_area.find(qn('c:barChart'))

    models = defect_matrix["models"]
    defects = defect_matrix["defects"]
    values = defect_matrix["values"]
    start_row = defect_matrix["start_row"]
    end_row = defect_matrix["end_row"]
    defect_cols = defect_matrix["defect_cols"]

    # Remove all old series
    for old_ser in bar_chart.findall(qn('c:ser')):
        bar_chart.remove(old_ser)

    insert_pos = 0
    for idx, child in enumerate(bar_chart):
        if child.tag.endswith('varyColors') or child.tag.endswith('grouping'):
            insert_pos = idx + 1

    cat_pts = "".join([f'<c:pt idx="{i}"><c:v>{m}</c:v></c:pt>' for i, m in enumerate(models)])
    cat_f = f"'Current month'!$L${start_row}:$L${end_row}"

    for s_idx, d_name in enumerate(defects):
        hex_color = get_defect_color(d_name, color_map, fallback_cache)
        font_color = get_label_font_color(hex_color)
        col_letter = get_column_letter(defect_cols[d_name])

        vals = values.get(d_name, [None] * len(models))
        val_pts = []
        for i, v in enumerate(vals):
            if v is not None and v > 0:
                val_pts.append(f'<c:pt idx="{i}"><c:v>{v:.6f}</c:v></c:pt>')
        val_pts_str = "".join(val_pts)

        ser_xml = f"""
        <c:ser xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <c:idx val="{s_idx}"/>
          <c:order val="{s_idx}"/>
          <c:tx>
            <c:strRef>
              <c:f>'Current month'!${col_letter}$21</c:f>
              <c:strCache>
                <c:ptCount val="1"/>
                <c:pt idx="0"><c:v>{d_name}</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:tx>
          <c:spPr>
            <a:solidFill><a:srgbClr val="{hex_color}"/></a:solidFill>
            <a:ln><a:noFill/></a:ln>
          </c:spPr>
          <c:invertIfNegative val="0"/>
          <c:dLbls>
            <c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>
            <c:txPr>
              <a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" lIns="19050" tIns="9525" rIns="19050" bIns="9525" anchor="ctr" anchorCtr="1"><a:spAutoFit/></a:bodyPr>
              <a:lstStyle/>
              <a:p>
                <a:pPr>
                  <a:defRPr lang="en-US" sz="700" b="1" i="0" baseline="0">
                    <a:solidFill><a:srgbClr val="{font_color}"/></a:solidFill>
                    <a:latin typeface="Calibri"/>
                  </a:defRPr>
                </a:pPr>
              </a:p>
            </c:txPr>
            <c:numFmt formatCode="0.00%" sourceLinked="0"/>
            <c:dLblPos val="ctr"/>
            <c:showLegendKey val="0"/>
            <c:showVal val="1"/>
            <c:showCatName val="0"/>
            <c:showSerName val="0"/>
            <c:showPercent val="0"/>
            <c:showBubbleSize val="0"/>
            <c:showLeaderLines val="0"/>
          </c:dLbls>
          <c:cat>
            <c:strRef>
              <c:f>{cat_f}</c:f>
              <c:strCache>
                <c:ptCount val="{len(models)}"/>
                {cat_pts}
              </c:strCache>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:f>'Current month'!${col_letter}${start_row}:${col_letter}${end_row}</c:f>
              <c:numCache>
                <c:formatCode>0.00%</c:formatCode>
                <c:ptCount val="{len(models)}"/>
                {val_pts_str}
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
        """
        bar_chart.insert(insert_pos + s_idx, parse_xml(ser_xml))

    # Category axis font size = 6pt
    catAx = plot_area.find(qn('c:catAx'))
    if catAx is not None:
        txPr = catAx.find(qn('c:txPr'))
        if txPr is not None:
            p = txPr.find(qn('a:p'))
            if p is not None and p.find(qn('a:pPr')) is not None and p.find(qn('a:pPr')).find(qn('a:defRPr')) is not None:
                p.find(qn('a:pPr')).find(qn('a:defRPr')).set('sz', '600')

    update_chart_relationships(chart_shape)


# ======================================================================
# 6. MAIN PRESENTATION GENERATION WORKFLOW
# ======================================================================
def generate_hfpa_presentation(target_month=None, template_path=TEMPLATE_PPTX, output_path=OUTPUT_PPTX):
    """
    Main entry point: Read data from Output/, map into PPTX presentation,
    re-render all 12 charts, generate native Defect Color Legend table,
    set hyperlinks, and save both in Database/ and Output/.
    """
    print("\n" + "=" * 65)
    print("   GENERATING HFPA POWERPOINT PRESENTATION (HFPA_Template.pptx)")
    print("   [Typography & Color Palette Enhanced]")
    print("=" * 65)

    # Use pristine backup if available, otherwise template
    source_pptx = BACKUP_PPTX if os.path.exists(BACKUP_PPTX) else template_path
    if not os.path.exists(source_pptx):
        print(f"[-] Error: Master template not found at {source_pptx}")
        return False

    # 1. Extract data from Output/
    print("\n[1/6] Extracting quality & defect metrics from Output/...")
    data = extract_pipeline_data(target_month=target_month)
    month = data["target_month"]

    # 2. Load defect colors
    print("\n[2/6] Loading defect color palette from Color_template.xlsx...")
    color_map = load_defect_color_map(COLOR_TEMPLATE)
    fallback_cache = {}

    # 3. Load presentation & backup
    print("\n[3/6] Loading PPTX presentation and creating safe backup...")
    if not os.path.exists(BACKUP_PPTX):
        try:
            shutil.copy2(template_path, BACKUP_PPTX)
            print(f"   [+] Created permanent pristine backup: {os.path.basename(BACKUP_PPTX)}")
        except Exception as e:
            print(f"   [-] Backup notice: {e}")

    prs = pptx.Presentation(source_pptx)
    slide = prs.slides[0]

    # 4. Update texts, metric summaries, and hyperlinks
    print("\n[4/6] Updating Slide text summaries, KPI cards, and interactive hyperlinks...")
    update_slide_metrics_and_links(slide, data)

    # 5. Create Dynamic Defect Color Legend Table (Replacing Picture 2)
    print("\n[5/6] Generating native Defect Color Legend Table (Color_template mapping)...")
    create_defect_color_legend_table(slide, data, color_map, fallback_cache)

    # 6. Update 12 Charts
    print("\n[6/6] Rebuilding and restyling 12 charts with enhanced typography...")
    shape_dict = {s.name: s for s in slide.shapes if s.has_chart}

    for fty, cfg in SITE_CHART_MAP.items():
        print(f"   ├─ Processing Site [{fty}]:")

        # Column 1: HFPA Analysis
        c_analysis_name = cfg["analysis"]
        if c_analysis_name in shape_dict:
            c_shape = shape_dict[c_analysis_name]
            update_hfpa_analysis_chart(
                c_shape,
                data["active_months"],
                data["monthly_metrics"][fty],
                cfg["excel_rows"]["prod"],
                cfg["excel_rows"]["defect"],
                cfg["excel_rows"]["hfpa"],
            )
            print(f"   │  ├─ Updated {c_analysis_name} ('HFPA Analysis' Combo Chart)")

        # Column 2: Top 5 models
        c_top5_name = cfg["top5"]
        if c_top5_name in shape_dict:
            c_shape = shape_dict[c_top5_name]
            update_top5_models_chart(
                c_shape,
                data["top5_data"][fty],
                cfg["excel_rows"]["top5_start"],
                cfg["excel_rows"]["top5_end"],
            )
            print(f"   │  ├─ Updated {c_top5_name} ('Top 5 models' Clustered Column)")

        # Column 3: Top defect of top models
        c_top_def_name = cfg["top_defect"]
        if c_top_def_name in shape_dict:
            c_shape = shape_dict[c_top_def_name]
            update_top_defects_chart(
                c_shape,
                data["top_defects_matrix"][fty],
                color_map,
                fallback_cache,
            )
            def_count = len(data["top_defects_matrix"][fty]["defects"])
            print(f"   │  └─ Updated {c_top_def_name} ('Top defect of top models' {def_count} defect series)")

    # 7. Save presentation to Database/ and Output/
    target_output = ensure_file_writable(output_path, "Báo cáo PowerPoint đầu ra (Output/HFPA_Performance_Report.pptx)")
    prs.save(target_output)
    print(f"\n[+] Successfully saved presentation to Output: {target_output}")

    target_database = ensure_file_writable(template_path, "Template PowerPoint gốc (Database/HFPA_Template.pptx)", max_retries=3)
    prs.save(target_database)
    print(f"[+] Master presentation template synchronized: {target_database}")

    print("\n" + "=" * 65)
    print(f"   COMPLETED! HFPA PowerPoint Report generated for Month: {month.upper()}")
    print("=" * 65 + "\n")
    return True


def main():
    target_month = sys.argv[1] if len(sys.argv) > 1 else None
    generate_hfpa_presentation(target_month=target_month)


if __name__ == "__main__":
    main()
