# QAStation / HFPA / FTT Data Processing Tool — AI Agent Build Specification

## 1. Objective

Create a Python tool that automates the QAStation, HFPA and FTT data-processing workflow, plus a Windows `.bat` launcher.

The tool must reproduce the business logic described in the QAStation workflow and the HFPA SOP. It should replace manual Excel PivotTable/filter/formula operations with reproducible Python processing.

Do not invent business rules that are not supported by the supplied requirements or SOP.

---

## 2. Input folder structure

The user will place source files into the tool's input directory.

Recommended structure:

```text
QAStation_Tool/
├── input/
│   ├── HFPA/
│   ├── FTT/
│   └── QAStation/
├── output/
├── logs/
└── ...
```

### HFPA input

The HFPA folder may contain:

- HFPA Quality Tracking files
- HFPA MES410 files

The tool must distinguish the two datasets from filename and/or detected columns where possible.

**HFPA MES410 is the reference source for reconciliation against HFPA Quality Tracking.**

### FTT input

The FTT folder contains FTT source files supplied by the user.

The exact FTT business logic is not defined in the current specification/SOP. Therefore:

- Load and validate FTT files.
- Make FTT available to the processing workflow.
- Keep FTT processing modular.
- Do **not** invent additional FTT calculation or mapping rules.

### QAStation input

The QAStation folder contains QAStation source files.

The tool should automatically discover files instead of requiring individual filenames to be hard-coded.

If a required input category is missing, report a clear error identifying the missing category.

---

# 3. QAStation workflow

## 3.1 FTY mapping

For each QAStation file:

1. Add a column named `FTY`.
2. Determine/map the factory from the ending of the file name.
3. Keep the FTY mapping configurable rather than hard-coding it inside processing logic.

## 3.2 QAStation Pivot #1

Create the first aggregation using:

- `InspDate`
- `FTY`
- `Station`
- `Line`
- `Shoename`
- `InspQty`
- `FailQty`

The purpose is to compare QAStation data with the corresponding HFPA reference data.

## 3.3 QAStation validation against HFPA

Validate QAStation against HFPA.

The validation must identify incorrect QAStation data and determine the corresponding correct values from the reference data.

The comparison should support:

- FTY
- Date
- Station
- Line
- Shoe/model

When a mismatch is identified:

1. Identify the exact mismatch.
2. Determine the corresponding reference value.
3. Correct the QAStation data.
4. Recalculate downstream aggregations.

## 3.4 Top 5 models/pairs

After validation/correction:

1. Aggregate the relevant production and failure values.
2. Sort descending.
3. Select the Top 5 models/pairs with the highest failure quantity.

The resulting values must provide accurate:

- `Sum of InspQty`
- `Sum of FailQty`

## 3.5 Database mapping

For the current month:

- `Sum of InspQty` → `Total pair produced`
- `Sum of FailQty` → `Defect`

Calculate:

```text
DR% = Sum of FailQty / Grand Total
```

Map the resulting percentage into the database/output.

Repeat for all 4 factories.

---

# 4. HFPA SOP workflow

The HFPA SOP requires preparation and reconciliation of:

- HFPA Quality Tracking
- HFPA MES410

## 4.1 HFPA Quality Tracking preparation

The SOP specifies adding a factory column after `Insp Date`.

Create the first aggregation with:

- `Insp Date`
- `Line`
- `Shoe Name`
- `Insp Q’ty`
- `Fail Q’ty`
- `Station`
- `FTY`

Duplicate values/records represented in the source must not be silently removed.

The Excel SOP uses:

- Repeat All Items Labels
- Show in Tabular Form
- Grand Totals off for rows and columns
- Subtotals not shown

The Python implementation should reproduce the resulting data aggregation rather than automate Excel mouse/keyboard actions.

## 4.2 HFPA Quality Tracking vs MES410

The SOP explicitly requires comparison between:

**HFPA Quality Tracking**

and

**HFPA MES410**

MES410 is the reference when Quality Tracking does not align.

Required semantic mapping:

| HFPA Quality Tracking | HFPA MES410 |
|---|---|
| `Insp Q’ty` | `Audit Sample Size` |
| `Fail Q’ty` | `Defective Q’ty` |
| `Insp Date` | `Audit Date` |
| `Shoe Name` / `Model` | `Model2` / model field |
| `FTY` | `Plant` |
| `Line` | `Line` |

The SOP defines:

- Inspection Qty as Audit Sample Size.
- FailQty as Defective Quantity in pairs.
- InspDate as Audit Date.

If Quality Tracking and MES410 are not aligned, the MES410 data must be used to correct the corresponding Quality Tracking data.

## 4.3 Reconciliation granularity

The comparison must not only compare overall totals.

Reconcile at the most specific available level:

1. Factory / Plant
2. Date / Audit Date
3. Line
4. Model / Shoe Name
5. Station where applicable

The reconciliation result should identify:

- comparison key
- match/mismatch status
- Quality Tracking value
- MES410 reference value
- corrected value
- relevant source record where available

## 4.4 Required HFPA comparison aggregations

### HFPA Quality Tracking — factory level

Rows:

- `FTY`

Values:

- `Insp Q’ty`
- `Fail Q’ty`

Filters:

- `Insp Date`
- `Station`

### HFPA Quality Tracking — line level

Rows:

- `Line`

Values:

- `Insp Q’ty`
- `Fail Q’ty`

Filters:

- `Insp Date`
- `Station`
- `FTY`

### HFPA MES410

Rows:

- `Audit Date`
- `Model2`

Values:

- `Audit Sample Size`
- `Defective Q’ty`

Filters:

- `Plant`
- `Line`

These are the business aggregations represented by the SOP.

## 4.5 HFPA data correction

When a mismatch is found:

1. Locate the corresponding Quality Tracking records using date, line, model, factory/plant and station where applicable.
2. Use MES410 as the reference.
3. Correct:

```text
Insp Q’ty = Audit Sample Size
Fail Q’ty = Defective Q’ty
```

4. Recalculate all affected aggregations.
5. Use corrected values for every downstream calculation.

The SOP explicitly requires affected pivot calculations to be refreshed/recalculated after correction.

---

# 5. Top 5 model calculation

After reconciliation and correction:

1. Aggregate `Insp Q’ty` and `Fail Q’ty` as required.
2. Sort descending.
3. Select Top 5 models.
4. Calculate:

```text
Defect Rate = Fail Q’ty / Sum Insp Q’ty
```

5. Export/map the Top 5 results into the summary report/database output.

The output must preserve descending order.

---

# 6. Pivot / analysis #2 — defects

Create a second analysis focused on defect/issue type and quantity.

Use:

- Defect/Issue type
- `Issues Q’ty`

Sort defect quantities descending.

## Top 3 defects of each Top 5 model

For each selected Top 5 model:

1. Filter by:
   - Shoe Name
   - Station
   - FTY
2. Aggregate `Issues Q’ty`.
3. Sort descending.
4. Select Top 3 defects.
5. Calculate:

```text
Defect Rate % = Issues Qty / Total Issue Q’ty
```

6. Produce the equivalent of the SOP's copy-and-transpose summary output.

The final defect list must remain sorted from highest to lowest for visual clarity.

---

# 7. Four-factory processing

The complete workflow must be executed for all 4 factories.

Each factory should be processed independently for:

1. Input loading
2. FTY/Plant mapping
3. QAStation/HFPA reconciliation
4. Data correction
5. Production/failure aggregation
6. Top 5 model selection
7. Defect-rate calculation
8. Top 3 defect selection
9. Defect-rate calculation
10. Database/summary mapping

Do not mix factory-specific reconciliation data unless an explicit consolidated calculation is required.

---

# 8. Database/output mapping

The database layer must be configurable because the exact database schema has not been supplied.

At minimum, prepare an output/mapping layer for:

- Factory
- Current month
- Total pair produced
- Defect
- DR%
- Top 5 models
- Top 5 model production/failure quantities
- Top 3 defects for each Top 5 model
- Defect percentage

If the actual database schema is unavailable, do not invent one.

Instead:

- Export a clearly structured Excel/CSV output.
- Keep database integration isolated in a separate module.
- Document exactly which fields require mapping once the database template is provided.

---

# 9. FTT handling

FTT is an expected input source.

The tool must:

1. Discover FTT files in `input/FTT/`.
2. Validate that files can be read.
3. Record detected columns and file metadata.
4. Keep FTT processing in a separate modular layer.

Do not infer additional FTT calculations from the HFPA SOP.

When the exact FTT business rules/schema are supplied, the FTT module can be connected to the appropriate calculations.

---

# 10. Recommended Python architecture

Use Python, preferably:

- `pandas` for tabular processing
- `openpyxl` for Excel input/output where required
- `pathlib` for filesystem handling
- `logging` for processing logs

Suggested structure:

```text
QAStation_Tool/
├── run_tool.bat
├── main.py
├── config.yaml
├── requirements.txt
├── input/
│   ├── HFPA/
│   ├── FTT/
│   └── QAStation/
├── output/
├── logs/
└── src/
    ├── __init__.py
    ├── file_discovery.py
    ├── loaders.py
    ├── column_mapping.py
    ├── fty_mapping.py
    ├── hfpa_mes410.py
    ├── hfpa_quality_tracking.py
    ├── reconciliation.py
    ├── qa_station.py
    ├── ftt.py
    ├── top_models.py
    ├── defect_analysis.py
    ├── database_mapping.py
    └── exporters.py
```

Keep functions modular and testable.

---

# 11. Error handling

The tool must clearly report:

- Missing input folder
- Missing HFPA files
- Missing MES410 data
- Missing QAStation files
- Missing FTT files
- Unsupported file format
- Missing required columns
- Invalid dates
- Invalid numeric quantities
- Duplicate/unexpected records
- Factory mapping failure
- Unable to match Quality Tracking to MES410
- Database/output mapping failure

The tool must not silently discard records.

---

# 12. Logging and auditability

Every run should produce a readable log containing:

- Run timestamp
- Files discovered
- Files processed
- Factory detected
- Number of source records
- Number of matched records
- Number of mismatched records
- Number of corrected records
- Top 5 models
- Top 3 defects per Top 5 model
- Output files generated
- Errors/warnings

For corrections, retain a reconciliation/audit table showing:

```text
Factory
Date
Line
Station
Model
Field
Original Value
MES410 Reference Value
Corrected Value
Status
```

This is important because the workflow changes Quality Tracking data based on MES410.

---

# 13. Excel-to-Python implementation principle

Do not literally automate Excel mouse/keyboard operations.

Translate the SOP operations into deterministic Python logic:

| Excel SOP action | Python implementation |
|---|---|
| Insert PivotTable | pandas groupby/pivot |
| Filter | dataframe filtering |
| Sort Descending | dataframe sort_values |
| Excel `=` comparison | programmatic comparison |
| Refresh PivotTable | recompute aggregations |
| Copy to new table | create output dataframe |
| Copy and Transpose | reshape/pivot output |
| Grand Totals off | control aggregation output |
| Repeat All Items Labels | explicit values in output columns |

The final numerical/business results should match the SOP while eliminating repetitive manual Excel work.

---

# 14. Expected end-to-end flow

```text
INPUT
│
├── HFPA
│   ├── Quality Tracking
│   └── MES410
│
├── FTT
│
└── QAStation
        │
        ▼
Discover and validate files
        │
        ▼
Map FTY / Plant
        │
        ▼
Load QAStation + HFPA
        │
        ▼
HFPA Quality Tracking ↔ HFPA MES410 reconciliation
        │
        ├── Match
        │
        └── Mismatch
              │
              ▼
        Correct Quality Tracking from MES410
              │
              ▼
        Recalculate aggregations
        │
        ▼
QAStation validation/correction
        │
        ▼
Pivot / aggregation #1
        │
        ▼
Sum Insp Q’ty + Sum Fail Q’ty
        │
        ▼
Sort descending
        │
        ▼
Top 5 Models
        │
        ├── Defect Rate
        │
        └── Database/Summary mapping
        │
        ▼
Defect analysis
        │
        ▼
Top 3 defects per Top 5 model
        │
        ├── Issues Qty
        ├── Defect Rate %
        │
        └── Summary mapping
        │
        ▼
Repeat for 4 factories
        │
        ▼
FINAL OUTPUT
```

---

# 15. Required deliverables

The AI coding agent must create:

1. A working Python tool.
2. `run_tool.bat` for Windows.
3. `requirements.txt`.
4. Configuration file if needed.
5. Structured output files.
6. Processing/reconciliation log.
7. README explaining:
   - input folder structure
   - how to run the tool
   - expected file formats
   - required columns
   - output structure
   - assumptions
   - unresolved FTT/database requirements

The `.bat` file should allow the user to run the tool without manually entering Python commands.

Example:

```bat
@echo off
cd /d "%~dp0"
python main.py
pause
```

If a virtual environment is used, the `.bat` should activate/use it appropriately.

---

# 16. Important business-rule constraints

1. **MES410 is the reference for HFPA Quality Tracking reconciliation.**
2. Do not replace MES410 reference values with Quality Tracking values when they disagree.
3. Corrections must be traceable in an audit/reconciliation output.
4. Recalculate downstream results after corrections.
5. Top 5 models are selected after reconciliation/correction.
6. Top 3 defects are selected within the Top 5 models.
7. Results are sorted descending where specified by the SOP.
8. The workflow is repeated for 4 factories.
9. FTT is an input source, but its exact calculation/mapping rules are not yet defined.
10. Database integration must remain configurable until the actual database template/schema is supplied.
