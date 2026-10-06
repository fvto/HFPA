# Project Guidelines: HFPA & QAStation Analytics Pipeline

## 1. Defect Color Palette Standard (Color_template.xlsx)

Whenever generating reports, PowerPoint charts, Excel summaries, or updating [`Color_template.xlsx`](file:///c:/Users/User/Desktop/HFPA/Color_template.xlsx), all agents MUST strictly comply with the following rules:

### 1.1. Absolute Prohibitions (Critical Safety Rule)
1. **NEVER USE RED (Tuyệt đối không dùng màu Đỏ)**:
   - Reserved strictly and exclusively for critical defect indicators and scrapped shoes (giày lỗi nghiêm trọng / phế phẩm).
   - Exclude any hue in the red spectrum ($335^\circ \le H \le 24^\circ$) or dominant red RGB ($R > 185, G < 95, B < 95$).
2. **NEVER USE GREEN (Tuyệt đối không dùng màu Xanh lá)**:
   - Reserved strictly and exclusively for passing inspection and quality conformance (giày đạt chuẩn / Pass).
   - Exclude any hue in the green spectrum ($62^\circ \le H \le 168^\circ$) or dominant green RGB ($G > 155, R < 120, B < 120$).

### 1.2. Human-Centric Visual Contrast (Không gây rối mắt, phân biệt rõ ràng)
- **NO NEON COLORS**: Do not use glaring neon hues.
- **NO MONOTONOUS BANDING**: NEVER generate adjacent defects using the same or similar colors. Under NO circumstances should consecutive rows share purplish, bluish, or dark muddy tones.
- **11-Family Round-Robin Interleaving**: When provisioning colors for new defects, strictly rotate across 11 distinct families so adjacent rows strongly contrast:
  1. **Amber / Gold** (`#D4AC0D`, `#DAA520`, `#EAA221`...)
  2. **Navy / Royal Blue** (`#1F77B4`, `#1034A6`, `#2980B9`...)
  3. **Berry / Magenta** (`#C51162`, `#C27BA0`, `#6B2D5C`...)
  4. **Sand / Desert Khaki** (`#C19A6B`, `#EDC9AF`, `#E5AA70`...)
  5. **Royal Purple / Violet** (`#5E35B1`, `#7852FF`, `#3F51B5`...)
  6. **Sky / Steel Blue** (`#3498DB`, `#729FCF`, `#4A6572`...)
  7. **Warm Orange / Tangerine** (`#ED7D31`, `#E68A00`, `#EF6C00`...)
  8. **Slate / Charcoal Pewter** (`#37474F`, `#78909C`, `#455A64`...)
  9. **Lavender / Lilac Pastel** (`#BA55D3`, `#A569BD`, `#9E7B9B`...)
  10. **Brown / Caramel Mocha** (`#8B4513`, `#8C6239`, `#935116`...)
  11. **Dark Petrol Teal** (`#006064`, `#3B6978`, `#0E6655`...)
- Ensure minimum Delta-E optical contrast between consecutive rows is at least $> 25$ (average $\approx 85$).

### 1.3. Excel Formatting & Synchronization Specifications
- **Sheet1**:
  - Column C: Defect type (Calibri 14pt, left-aligned, thin border).
  - Column D: Reference swatch (Solid PatternFill `FF` + HEX, thin border).
  - Column E, F, G: Red, Green, Blue integers (Calibri 11pt, center-aligned, thin border).
  - Column H: 6-character uppercase HEX without `#` (Calibri 11pt, center-aligned, thin border).
  - Row height: exactly 18.0 pt.
- **Sheet bc**:
  - Must keep columns B (Red), C (Green), D (Blue), E (HEX), and F (Reference Swatch) fully populated and synchronized.
- **Defect Name Normalization**:
  - Always strip legacy non-breaking spaces (`\xa0`) and redundant whitespace so names like `Hairy edge`, `Stitching margin/SPI`, `Color migration` correctly match SOP definitions.
- **Code Execution**:
  - Always utilize [`agents/color_manager.py`](file:///c:/Users/User/Desktop/HFPA/agents/color_manager.py) to manage and provision defect colors.

## 2. Automated PowerPoint Annotation & Framing Standard (pptx_annotator.py)

Whenever adding visual overlays (bounding frames, connector lines, insight callouts) on Re-Inspection PowerPoint reports, all agents MUST strictly follow:

### 2.1. Threshold Rules for Bounding Box Framing (Quy tắc khoanh vùng)
1. **By QC Threshold (10%)**:
   - Only frame a shoe model if there is at least one defect where:
     $$\frac{\text{Defect Count in Re-ins}}{\text{Defect Count in FTT}} \ge 10\%$$
   - If no defect reaches $\ge 10\%$, do NOT frame the model.
2. **By MA Threshold (30%)**:
   - Only frame a shoe model if there is at least one defect where:
     $$\frac{\text{Defect Count in Re-ins}}{\text{Defect Count in HFPA}} \ge 30\%$$
   - If no defect reaches $\ge 30\%$, do NOT frame the model.

### 2.2. Defect Mapping Standard (Quy tắc nối đường chỉ dẫn)
1. **Strict Same-Defect Mapping (Lỗi A chỉ nối với Lỗi A)**:
   - A connector line MUST connect the exact same defect name and color palette across the FTT/HFPA column and Re-ins column (e.g. Over cement $\rightarrow$ Over cement, Cleanness $\rightarrow$ Cleanness).
   - NEVER connect different defects or different color segments together.
2. **Inter-Column Gap Placement (Không đè lên số liệu)**:
   - Connector lines MUST strictly span the inter-column gap between the two bars (width $\approx 0.34 - 0.36\text{ in}$).
   - Endpoints must sit on the inner boundary edges of the colored segments with compact dots (`headEnd` & `tailEnd` set to `oval`, size `sm`).
   - NEVER cross through the center of the columns where data label numbers are displayed. All numbers must remain 100% visible and unblocked.

