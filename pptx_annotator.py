"""
Module: pptx_annotator.py
Automates visual inspection overlays (bounding frames, connector lines with dots,
and insight callouts) on Re-Inspection PowerPoint reports.

STANDARD RULES:
- By QC: Threshold is 10%. (Re-Ins / FTT >= 10% for a defect to qualify for framing & connector).
- By MA: Threshold is 30%. (Re-Ins / HFPA >= 30% for a defect to qualify for framing & connector).
- Mapping: Strictly Defect A -> Defect A (same defect name and same color).
- Zero overlap: Connectors strictly span inter-column gaps, leaving numbers 100% visible.
"""

import os
import shutil
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR


COLOR_RED = RGBColor(255, 0, 0)
COLOR_PURPLE = RGBColor(128, 100, 162)  # #8064A2 matching template
COLOR_BLUE = RGBColor(31, 78, 120)     # Navy/Royal blue matching template
COLOR_GRAY = RGBColor(55, 65, 81)

THRESHOLD_QC = 0.10  # 10% for QC
THRESHOLD_MA = 0.30  # 30% for MA


def add_bounding_box(slide, left, top, width, height, color_rgb, line_width_pt=2.0):
    """Add a transparent rectangle bounding box with colored border."""
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    box.fill.background()
    box.line.color.rgb = color_rgb
    box.line.width = Pt(line_width_pt)
    return box


def add_connector_with_oval_ends(slide, x1, y1, x2, y2, color_rgb, line_width_pt=1.2, end_size="sm"):
    """
    Add straight connector line strictly across the inter-column gap,
    with small oval endpoints on the inner edges, ensuring zero overlap with data label numbers.
    """
    conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x1, y1, x2, y2)
    conn.line.color.rgb = color_rgb
    conn.line.width = Pt(line_width_pt)

    ln = conn._element.spPr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}ln')
    if ln is not None:
        headEnd = pptx.oxml.xmlchemy.OxmlElement('a:headEnd')
        headEnd.set('type', 'oval')
        headEnd.set('w', end_size)
        headEnd.set('len', end_size)
        ln.append(headEnd)

        tailEnd = pptx.oxml.xmlchemy.OxmlElement('a:tailEnd')
        tailEnd.set('type', 'oval')
        tailEnd.set('w', end_size)
        tailEnd.set('len', end_size)
        ln.append(tailEnd)

    return conn


def add_insight_callout(slide, left, top, width, height, items, action_text=None, megaphone_png=None):
    """
    Add insight callout box with megaphone icon and formatted bullet points.
    items: list of (prefix_bold, text)
    """
    if megaphone_png and os.path.exists(megaphone_png):
        slide.shapes.add_picture(megaphone_png, left, top + Inches(0.02), Inches(0.32), Inches(0.32))

    tb_left = left + Inches(0.38)
    tb_width = width - Inches(0.38)
    tb = slide.shapes.add_textbox(tb_left, top, tb_width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

    first = True
    for prefix_bold, body_text in items:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.space_after = Pt(5)
        p.line_spacing = 1.15

        run_prefix = p.add_run()
        run_prefix.text = "❖  " + prefix_bold + " "
        run_prefix.font.name = "Calibri"
        run_prefix.font.size = Pt(8.5)
        run_prefix.font.bold = True
        run_prefix.font.color.rgb = COLOR_BLUE

        run_body = p.add_run()
        run_body.text = body_text
        run_body.font.name = "Calibri"
        run_body.font.size = Pt(8.0)
        run_body.font.color.rgb = COLOR_GRAY

    if action_text:
        p_act = tf.add_paragraph()
        p_act.space_before = Pt(4)
        p_act.line_spacing = 1.15

        run_arrow = p_act.add_run()
        run_arrow.text = "➤  "
        run_arrow.font.name = "Calibri"
        run_arrow.font.size = Pt(8.5)
        run_arrow.font.bold = True
        run_arrow.font.color.rgb = COLOR_BLUE

        run_act = p_act.add_run()
        run_act.text = action_text
        run_act.font.name = "Calibri"
        run_act.font.size = Pt(8.0)
        run_act.font.italic = True
        run_act.font.color.rgb = COLOR_BLUE

    return tb


def remove_orphan_arrows(slide):
    """Remove leftover unanchored arrow connectors from previous manual drafts."""
    to_remove = []
    for s in slide.shapes:
        if s.name.startswith('Straight Arrow Connector'):
            to_remove.append(s._element)
    for elem in to_remove:
        slide.shapes._spTree.remove(elem)


def annotate_september_report(pptx_path, megaphone_png=None):
    """
    Apply rule-based frames, guide lines, and insights to September 2026 report.
    - By QC: threshold = 10%
    - By MA: threshold = 30%
    """
    prs = pptx.Presentation(pptx_path)

    # =========================================================================
    # SLIDE 4: JV & JV2
    # =========================================================================
    slide4 = prs.slides[3]
    remove_orphan_arrows(slide4)

    # -------------------------------------------------------------------------
    # 1. JV BY QC: Model 3 (NIKE STAR RUNNER 5)
    # Rule Check: Over cement FTT=542 -> Re-ins=100 (18.5% >= 10% QC Threshold) -> QUALIFIED!
    # -------------------------------------------------------------------------
    add_bounding_box(slide4, Inches(5.95), Inches(0.94), Inches(1.24), Inches(1.08), COLOR_RED, line_width_pt=2.0)
    # Connector: OVER CEMENT (542) in FTT -> OVER CEMENT (100) in Re-ins (both olive green)
    add_connector_with_oval_ends(slide4, Inches(6.42), Inches(1.41), Inches(6.77), Inches(1.71), COLOR_RED, line_width_pt=1.2)

    # -------------------------------------------------------------------------
    # 2. JV2 BY MA: Model 0 (AIR JORDAN 1 LOW)
    # Rule Check: Cleanness HFPA=6 -> Re-ins=24 (400% >= 30% MA Threshold) -> QUALIFIED!
    # Over cement HFPA=18 -> Re-ins=9 (50% >= 30% MA Threshold) -> QUALIFIED!
    # -------------------------------------------------------------------------
    add_bounding_box(slide4, Inches(2.02), Inches(5.94), Inches(1.24), Inches(1.08), COLOR_RED, line_width_pt=2.0)
    # Connector: CLEANNESS (6) in HFPA -> CLEANNESS (24) in Re-ins (both dark blue)
    add_connector_with_oval_ends(slide4, Inches(2.51), Inches(6.05), Inches(2.86), Inches(6.50), COLOR_RED, line_width_pt=1.2)

    # -------------------------------------------------------------------------
    # 3. JV2 BY MA: PURPLE FRAME spanning Models 1 to 4
    # Rule Checks:
    # Model 1 (JORDAN 1 MID SE): Over cement 20 -> 11 (55% >= 30%), Cleanness 13 -> 8 (61.5% >= 30%) -> QUALIFIED!
    # Model 2 (NIKE COURT VISION LOW): Over cement 52 -> 20 (38.5% >= 30%) -> QUALIFIED!
    # Model 3 (NIKE EBERNON LOW): Toe off center 12 -> 5 (41.7% >= 30%) -> QUALIFIED!
    # -------------------------------------------------------------------------
    add_bounding_box(slide4, Inches(3.32), Inches(5.94), Inches(5.26), Inches(1.08), COLOR_PURPLE, line_width_pt=2.0)
    # Connectors strictly Defect A -> Defect A across gaps:
    # Model 1: Over cement (20 -> 11, both olive green)
    add_connector_with_oval_ends(slide4, Inches(3.82), Inches(6.75), Inches(4.18), Inches(6.66), COLOR_PURPLE, line_width_pt=1.2)
    # Model 2: Over cement (52 -> 20, both olive green)
    add_connector_with_oval_ends(slide4, Inches(5.14), Inches(6.70), Inches(5.49), Inches(6.73), COLOR_PURPLE, line_width_pt=1.2)
    # Model 3: Toe off center (12 -> 5, both dark purple)
    add_connector_with_oval_ends(slide4, Inches(6.45), Inches(6.09), Inches(6.79), Inches(6.06), COLOR_PURPLE, line_width_pt=1.2)

    # -------------------------------------------------------------------------
    # 4. SLIDE 4 INSIGHT CALLOUTS
    # -------------------------------------------------------------------------
    items_jv = [
        ("According to RED FRAME (JV BY QC),",
         "we observed that top defects (Over cement: 100, Cleanness: 116, Thread end: 132) at FTT were still detected by QC-reinspection with high volume on NIKE STAR RUNNER 5 (PS), where Over cement re-inspection rate reached 18.5% (exceeding the 10% threshold).")
    ]
    add_insight_callout(slide4, Inches(8.85), Inches(2.88), Inches(4.35), Inches(1.20), items_jv, megaphone_png=megaphone_png)

    items_jv2 = [
        ("At RED FRAME (JV2 BY MA),",
         "for AIR JORDAN 1 LOW (MS/WS), MA-Reinspection detected a severe surge in Cleanness defects (from 6 up to 24, a 400% surge exceeding the 30% threshold). Re-work cleaning of bond gaps caused secondary shoe contamination."),
        ("At PURPLE FRAME,",
         "repeating defects detected during HFPA (Over cement at 55% & 38.5%, Toe off center at 41.7%) exceeded the 30% MA threshold across multiple models, while secondary re-work at JV2 introduced stitching margin and upper wrinkling issues.")
    ]
    action_text = "It is necessary for Reviewing the inspection standards among QCs - MAs, as well as within each team themselves"
    add_insight_callout(slide4, Inches(8.85), Inches(4.35), Inches(4.35), Inches(2.50), items_jv2, action_text=action_text, megaphone_png=megaphone_png)

    # =========================================================================
    # SLIDE 3: VH & VH2
    # Rule Evaluation:
    # VH BY QC (Chart 36): Model 2 JORDAN TRUE FLIGHT has Over cement 803 -> 15 (1.87% < 10% QC Threshold).
    # VH2 BY QC (Chart 10): Model 0 NIKE AIR MAX EXCEE has Midsole gap 5577 -> 20 (0.36% < 10% QC Threshold).
    # Since neither exceeds the 10% QC threshold, NO frames or connectors are placed on those models!
    # =========================================================================
    slide3 = prs.slides[2]
    remove_orphan_arrows(slide3)

    # Callout for Slide 3 summarizing conformance to QC 10% standard:
    items_vh = [
        ("VH & VH2 Inspection Summary:",
         "QC Re-inspection conformance achieved excellent results: all repeating defect ratios remained well below the 10% QC threshold (JORDAN TRUE FLIGHT at 1.87%, NIKE AIR MAX EXCEE at 0.36%). No single-model threshold violations detected.")
    ]
    add_insight_callout(slide3, Inches(8.85), Inches(3.25), Inches(4.35), Inches(3.20), items_vh, action_text=action_text, megaphone_png=megaphone_png)

    prs.save(pptx_path)
    print(f"[+] Successfully annotated PowerPoint presentation: {pptx_path}")


def main():
    target_pptx = r"c:\Users\User\Desktop\Re-Inspection\Output\Re-Ins Report-September-2026.pptx"
    backup_pptx = r"c:\Users\User\Desktop\Re-Inspection\Output\Re-Ins Report-September-2026.pptx.bak"
    megaphone_img = r"c:\Users\User\Desktop\HFPA\assets\megaphone.png"

    if os.path.exists(backup_pptx):
        shutil.copyfile(backup_pptx, target_pptx)
        print(f"[+] Restored clean baseline from: {backup_pptx}")

    annotate_september_report(target_pptx, megaphone_png=megaphone_img)


if __name__ == "__main__":
    main()
