"""
Module: agents.color_manager
Automates dynamic color provisioning for newly detected quality defects in Color_template.xlsx.

Design Criteria:
1. STRICTLY EXCLUDES RED (critical defect / fail indicator) and GREEN (pass / ok indicator).
2. Human-Perceptible Distinction:
   - Uses curated, professional, non-neon color families (Oranges, Blues, Purples,
     Blue-Grays, Mauves, Teals/Petroleum, Warm Chocolates, Bronzes, Cashmere).
   - Successive defects cycle across distinct color families to ensure high visual contrast.
3. Normalized Name Matching:
   - Handles legacy non-breaking spaces (\xa0), punctuation, and casing so defects like
     'Hairy edge' or 'Stitching margin/SPI' map to their official template colors.
4. Preserves 100% of Excel schema, borders, fonts (Calibri 14pt/11pt), and swatch fills.
"""

from __future__ import annotations

import colorsys
import math
import os
import re
import shutil
from typing import Any, Dict, List, Optional, Set, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


# Curated palette of visually distinct corporate colors:
# Strictly non-red (no red/crimson) and strictly non-green (no green/lime/emerald).
# 11 distinct color families: Golds, Navy/Royals, Berries/Magentas, Sands/Khakis,
# Violets/Purples, Sky/Steel Blues, Warm Oranges, Slates/Pewters, Lavenders/Lilacs,
# Earthy Browns/Caramels, and Dark Teals/Petroleums.
COLOR_FAMILIES = {
    "Amber/Gold": ["D4AC0D", "DAA520", "E1AD01", "B8860B", "EBA83A", "EAA221", "E3963E", "F4B41A", "C49102", "D89E00"],
    "Navy/Royal": ["1F77B4", "0047AB", "2980B9", "0F52BA", "2B547E", "1034A6", "2E4053", "26619C", "154360", "003366"],
    "Berry/Magenta": ["C51162", "871F78", "8B008B", "B53471", "A22A60", "C27BA0", "6B2D5C", "880E4F", "AD1457", "9C27B0"],
    "Sand/Khaki": ["C19A6B", "D5C4A1", "B38B6D", "C68A4C", "D2B48C", "EDC9AF", "E5AA70", "BCAAA4", "D7CCC8", "CFB997"],
    "Purple/Violet": ["5E35B1", "8E44AD", "6A1B9A", "5B2C6F", "4A148C", "7852FF", "3F51B5", "7D3C98", "512DA8", "4527A0"],
    "Sky/Steel Blue": ["3498DB", "4682B4", "5D8AA8", "6495ED", "6082B6", "729FCF", "4A6572", "537895", "5DADE2", "6BA4B8"],
    "Warm Orange": ["ED7D31", "F39C12", "E67E22", "B9770E", "F57F17", "E68A00", "EF6C00", "D87A00", "FF8F00", "FB8C00"],
    "Slate/Pewter": ["37474F", "2F4F4F", "696969", "546E7A", "52595D", "78909C", "91A3B0", "455A64", "607D8B", "5C6B73"],
    "Lavender/Lilac": ["BA55D3", "9370DB", "B39DDB", "9575CD", "857094", "A569BD", "9E7B9B", "C39BD3", "D2B4DE", "CE93D8"],
    "Brown/Caramel": ["8B4513", "7B3F00", "CD853F", "6F4E37", "D2691E", "8C6239", "A1887F", "935116", "7E5109", "6D4C41"],
    "Dark Teal/Petrol": ["006064", "00838F", "0097A7", "007A87", "1F4E5B", "3B6978", "0E6655", "004D40", "008B8B", "005B60"],
}


class ColorManager:
    """Manages color generation and synchronization for Color_template.xlsx."""

    def __init__(self, color_template_path: Optional[str] = None):
        if color_template_path is None:
            base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            self.color_template_path = os.path.join(base_dir, "Color_template.xlsx")
        else:
            self.color_template_path = color_template_path

    @staticmethod
    def normalize_name(name: str) -> str:
        """Strip punctuation, non-breaking spaces, and whitespace for fuzzy key lookup."""
        return re.sub(r"[^a-z0-9]", "", str(name).lower())

    @staticmethod
    def hex_to_rgb(hex_str: str) -> Tuple[int, int, int]:
        clean = hex_str.strip().replace("#", "").upper()
        return int(clean[0:2], 16), int(clean[2:4], 16), int(clean[4:6], 16)

    @staticmethod
    def rgb_to_hex(r: int, g: int, b: int) -> str:
        return f"{r:02X}{g:02X}{b:02X}"

    @staticmethod
    def is_red_or_green(r: int, g: int, b: int) -> bool:
        """
        Check if an RGB color falls into Red or Green spectrum.
        Red: dominant R with low G/B or hue in red arc.
        Green: dominant G with lower R/B or hue in green arc.
        """
        h, s, v = colorsys.rgb_to_hsv(r / 255.0, g / 255.0, b / 255.0)
        h_deg = h * 360.0

        # Exclude Red hue [335, 360] and [0, 24] with noticeable saturation
        if (h_deg >= 335.0 or h_deg <= 24.0) and s > 0.25:
            return True
        # Exclude Green hue [62, 168] with noticeable saturation
        if (62.0 <= h_deg <= 168.0) and s > 0.20:
            return True

        if r > 185 and g < 95 and b < 95:
            return True
        if g > 155 and r < 120 and b < 120:
            return True

        return False

    def load_existing_template(self) -> Tuple[Dict[str, str], Dict[str, str], Set[str], int]:
        """
        Load existing defect mapping (both exact and normalized), existing hexes,
        and last populated row.
        Returns: (exact_map, norm_map, existing_hexes, last_row)
        """
        exact_map: Dict[str, str] = {}
        norm_map: Dict[str, str] = {}
        existing_hexes: Set[str] = set()
        last_row = 4

        if not os.path.exists(self.color_template_path):
            return exact_map, norm_map, existing_hexes, last_row

        wb = openpyxl.load_workbook(self.color_template_path, data_only=True)
        if "Sheet1" in wb.sheetnames:
            ws = wb["Sheet1"]
            for r in range(5, ws.max_row + 1):
                dname = ws.cell(r, 3).value
                hex_val = ws.cell(r, 8).value
                if dname and str(dname).strip():
                    last_row = r
                    d_raw = str(dname).strip()
                    d_norm = self.normalize_name(d_raw)
                    if hex_val:
                        h_clean = str(hex_val).strip().replace("#", "").upper()
                        if len(h_clean) == 6:
                            exact_map[d_raw.lower()] = h_clean
                            norm_map[d_norm] = h_clean
                            existing_hexes.add(h_clean)

        if "bc" in wb.sheetnames:
            ws_bc = wb["bc"]
            for r in range(2, ws_bc.max_row + 1):
                dname = ws_bc.cell(r, 1).value
                hex_val = ws_bc.cell(r, 5).value
                if dname and str(dname).strip() and hex_val:
                    h_clean = str(hex_val).strip().replace("#", "").upper()
                    if len(h_clean) == 6:
                        d_raw = str(dname).strip()
                        exact_map[d_raw.lower()] = h_clean
                        norm_map[self.normalize_name(d_raw)] = h_clean
                        existing_hexes.add(h_clean)

        wb.close()
        return exact_map, norm_map, existing_hexes, last_row

    def get_curated_palette(self, count: int, existing_hexes: Set[str]) -> List[str]:
        """
        Return `count` distinct colors from the curated palette, interleaved across
        11 different color families so adjacent colors are strongly contrasting to human eyes.
        """
        order_keys = [
            "Amber/Gold", "Navy/Royal", "Berry/Magenta", "Sand/Khaki",
            "Purple/Violet", "Sky/Steel Blue", "Warm Orange", "Slate/Pewter",
            "Lavender/Lilac", "Brown/Caramel", "Dark Teal/Petrol"
        ]

        interleaved: List[str] = []
        max_len = max(len(COLOR_FAMILIES[k]) for k in order_keys)
        for i in range(max_len):
            for k in order_keys:
                fam = COLOR_FAMILIES[k]
                if i < len(fam):
                    c = fam[i].upper()
                    if c not in interleaved and c not in existing_hexes:
                        interleaved.append(c)

        # Dynamic fallback if count exceeds predefined 11-family pool
        extra_step = 0
        while len(interleaved) < count:
            extra_hues = [45.0, 215.0, 315.0, 35.0, 275.0, 200.0, 25.0, 250.0, 285.0, 38.0, 185.0]
            hue = extra_hues[extra_step % len(extra_hues)] / 360.0
            light = 0.35 + ((extra_step // len(extra_hues)) * 0.07) % 0.40
            sat = 0.60
            r_f, g_f, b_f = colorsys.hls_to_rgb(hue, min(light, 0.75), sat)
            r = int(round(r_f * 255))
            g = int(round(g_f * 255))
            b = int(round(b_f * 255))
            if not self.is_red_or_green(r, g, b):
                h_str = self.rgb_to_hex(r, g, b)
                if h_str not in interleaved and h_str not in existing_hexes:
                    interleaved.append(h_str)
            extra_step += 1
            if extra_step > 1000:
                break

        return interleaved[:count]

    def sync_unmapped_defects(self, defects: List[str]) -> Tuple[int, List[Dict[str, Any]]]:
        """
        Identify unmapped defects, generate distinct non-red/non-green colors,
        and append them with exact formatting into Color_template.xlsx.
        """
        exact_map, norm_map, existing_hexes, last_row = self.load_existing_template()

        unmapped = []
        seen = set()
        for d in defects:
            clean = str(d).strip()
            if not clean or clean.lower() == "nan":
                continue
            lower_d = clean.lower()
            norm_d = self.normalize_name(clean)

            # Check exact and normalized match
            if lower_d not in exact_map and norm_d not in norm_map and norm_d not in seen:
                seen.add(norm_d)
                unmapped.append(clean)

        if not unmapped:
            return 0, []

        print(f"[*] ColorManager: Detected {len(unmapped)} truly new defect types requiring distinct colors...")

        # Generate colors using human-friendly interleaved palette
        new_hexes = self.get_curated_palette(len(unmapped), existing_hexes)
        if len(new_hexes) < len(unmapped):
            raise RuntimeError(f"Could not provision enough distinct colors: {len(new_hexes)}/{len(unmapped)}")

        wb = openpyxl.load_workbook(self.color_template_path)
        ws = wb["Sheet1"]

        thin_side = Side(style="thin", color="000000")
        border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
        align_center = Alignment(horizontal="center", vertical="center")
        align_left = Alignment(horizontal="left", vertical="center")
        font_c = Font(name="Calibri", size=14, bold=False)
        font_data = Font(name="Calibri", size=11, bold=False)

        current_row = last_row + 1
        provisions = []

        for defect_name, hex_code in zip(unmapped, new_hexes):
            r, g, b = self.hex_to_rgb(hex_code)

            ws.row_dimensions[current_row].height = 18.0

            # Col C: Defect type
            cell_c = ws.cell(row=current_row, column=3, value=defect_name)
            cell_c.font = font_c
            cell_c.alignment = align_left
            cell_c.border = border_all

            # Col D: Reference (Color Swatch Fill)
            cell_d = ws.cell(row=current_row, column=4, value=None)
            cell_d.fill = PatternFill(start_color="FF" + hex_code, end_color="FF" + hex_code, fill_type="solid")
            cell_d.border = border_all

            # Col E: Red
            cell_e = ws.cell(row=current_row, column=5, value=r)
            cell_e.font = font_data
            cell_e.alignment = align_center
            cell_e.border = border_all

            # Col F: Green
            cell_f = ws.cell(row=current_row, column=6, value=g)
            cell_f.font = font_data
            cell_f.alignment = align_center
            cell_f.border = border_all

            # Col G: Blue
            cell_g = ws.cell(row=current_row, column=7, value=b)
            cell_g.font = font_data
            cell_g.alignment = align_center
            cell_g.border = border_all

            # Col H: HEX
            cell_h = ws.cell(row=current_row, column=8, value=hex_code)
            cell_h.font = font_data
            cell_h.alignment = align_center
            cell_h.border = border_all

            provisions.append({
                "defect": defect_name,
                "hex": hex_code,
                "rgb": (r, g, b),
                "row": current_row
            })

            current_row += 1

        wb.save(self.color_template_path)
        wb.close()

        print(f"[+] ColorManager: Successfully added {len(provisions)} distinct non-red, non-green colors to Color_template.xlsx (Rows {last_row + 1} to {current_row - 1})!")
        return len(provisions), provisions
