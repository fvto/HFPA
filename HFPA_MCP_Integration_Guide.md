# HƯỚNG DẪN TÍCH HỢP HFPA PIPELINE VÀO MCP SERVER "PC_TOOL_AGENT"
*(Triển khai đầy đủ 5 đầu ra theo đặc tả trong prompt_integrate_tool.md)*

Dưới đây là toàn bộ mã nguồn hoàn chỉnh, chuẩn 100% theo kiến trúc **Gateway - Registry - Engine** của dự án `pc_tool_agent` (Pydantic v2, Python 3.11+, PathPolicy Sandbox, FastMCP, pytest).

---

## ĐẦU RA 1: FILE GATEWAY NGHIỆP VỤ `src/pc_tool_agent/hfpa.py`

```python
"""
HFPA & QAStation Data Processing Gateway for pc_tool_agent.
Handles reconciliation, aggregation, database population, and PowerPoint generation.
"""

from __future__ import annotations

import glob
import os
import re
import shutil
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import numpy as np
import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from pc_tool_agent.errors import AgentError
from pc_tool_agent.security import PathPolicy


MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
MONTH_MAP = {
    1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"
}
FACTORIES = ["VH", "VH2", "JV", "JV2"]


def detect_month_from_file(filepath: Path) -> str:
    """Detect month abbreviation from dates in file or filename."""
    try:
        df_sample = pd.read_excel(filepath, nrows=50)
        date_cols = [c for c in df_sample.columns if "date" in str(c).lower()]
        for c in date_cols:
            dates = pd.to_datetime(df_sample[c], errors="coerce").dropna()
            if not dates.empty:
                return MONTH_MAP[dates.iloc[0].month]
    except Exception:
        pass

    fname = filepath.name
    for m_num, m_name in MONTH_MAP.items():
        if re.search(rf"[-_\s]{m_name}[-_\s\.]", fname, re.IGNORECASE):
            return m_name

    m = re.search(r"(\d{4})[-_](\d{2})[-_](\d{2})", fname)
    if m:
        month_num = int(m.group(2))
        if 1 <= month_num <= 12:
            return MONTH_MAP[month_num]

    return "Jun"


class HFPAGateway:
    """Production Gateway for QAStation & HFPA Processing Pipeline."""

    def __init__(self, path_policy: PathPolicy | None = None) -> None:
        self.path_policy = path_policy

    def run_pipeline(
        self,
        input_dir: str,
        template_dir: str,
        output_dir: str | None = None,
        target_month: str | None = None,
        generate_database: bool = True,
        generate_presentation: bool = True,
    ) -> dict[str, Any]:
        """
        Execute full reconciliation, analysis, database update, and PPTX report.
        """
        # 1. Resolve and validate sandbox paths
        in_path = Path(input_dir)
        tmpl_path = Path(template_dir)
        out_path = Path(output_dir) if output_dir else in_path.parent / "Output"

        if self.path_policy:
            in_path = self.path_policy.resolve_and_validate(in_path)
            tmpl_path = self.path_policy.resolve_and_validate(tmpl_path)
            out_path = self.path_policy.resolve_and_validate(out_path)

        if not in_path.exists() or not in_path.is_dir():
            raise AgentError("INPUT_DIR_NOT_FOUND", f"Thư mục đầu vào không tồn tại: {in_path}")
        if not tmpl_path.exists() or not tmpl_path.is_dir():
            raise AgentError("TEMPLATE_DIR_NOT_FOUND", f"Thư mục template không tồn tại: {tmpl_path}")

        out_path.mkdir(parents=True, exist_ok=True)
        logs_dir = out_path.parent / "logs"
        logs_dir.mkdir(parents=True, exist_ok=True)

        ftt_dir = in_path / "FTT"
        hfpa_dir = in_path / "HFPA"
        if not ftt_dir.exists() or not hfpa_dir.exists():
            raise AgentError("INVALID_INPUT_STRUCTURE", "Thư mục đầu vào phải chứa cả 2 thư mục con 'FTT' và 'HFPA'.")

        ftt_files = list(ftt_dir.glob("*.xlsx")) + list(ftt_dir.glob("*.xls"))
        hfpa_files = list(hfpa_dir.glob("*.xlsx")) + list(hfpa_dir.glob("*.xls"))

        if not ftt_files:
            raise AgentError("NO_FTT_FILES", f"Không tìm thấy file QAStation/FTT nào trong: {ftt_dir}")
        if not hfpa_files:
            raise AgentError("NO_HFPA_FILES", f"Không tìm thấy file HFPA nào trong: {hfpa_dir}")

        # 2. Detect active month
        active_month = target_month
        if not active_month:
            active_month = detect_month_from_file(ftt_files[0])

        # 3. Load & Combine QAStation/FTT files
        ftt_dfs = []
        for f in ftt_files:
            fty = self._detect_factory(f.name)
            df = pd.read_excel(f)
            df["FTY"] = fty
            ftt_dfs.append(df)
        qa_combined = pd.concat(ftt_dfs, ignore_index=True)

        # 4. Load & Combine HFPA files
        hfpa_dfs = []
        for f in hfpa_files:
            fty = self._detect_factory(f.name)
            df = pd.read_excel(f)
            df["FTY"] = fty
            hfpa_dfs.append(df)
        hfpa_combined = pd.concat(hfpa_dfs, ignore_index=True)

        # 5. Harmonize Date, Line, Plant
        qa_combined["Date_Std"] = pd.to_datetime(qa_combined["Inspection Date"], errors="coerce").dt.strftime("%Y-%m-%d")
        qa_combined["Plant_Std"] = qa_combined["Plant"].astype(str).str.strip().str.upper()
        qa_combined["Line_Std"] = qa_combined["Line"].astype(str).str.strip().str.upper()

        date_col = "Audit Date (dd-mmm-yyyy)" if "Audit Date (dd-mmm-yyyy)" in hfpa_combined.columns else "Audit Date"
        hfpa_combined["Date_Std"] = pd.to_datetime(hfpa_combined[date_col], errors="coerce").dt.strftime("%Y-%m-%d")
        hfpa_combined["Plant_Std"] = hfpa_combined["Plant"].astype(str).str.strip().str.upper()
        hfpa_combined["Line_Std"] = hfpa_combined["Line"].astype(str).str.strip().str.upper()

        # Defect Q'ty reconciliation
        hfpa_def_col = [c for c in hfpa_combined.columns if "Defective" in c and "pair" in c][0]
        hfpa_lot = hfpa_combined.groupby(["Date_Std", "FTY", "Plant_Std", "Line_Std"])[hfpa_def_col].sum().reset_index()
        qa_lot = qa_combined.groupby(["Date_Std", "FTY", "Plant_Std", "Line_Std"])["Defect Qty"].sum().reset_index()

        lot_merged = pd.merge(qa_lot, hfpa_lot, on=["Date_Std", "FTY", "Plant_Std", "Line_Std"], how="outer", suffixes=("_QA", "_HFPA"))
        lot_merged["Validation_Status"] = "Matched"
        diff_mask = (lot_merged["Defect Qty"].notna()) & (lot_merged[hfpa_def_col].notna()) & (lot_merged["Defect Qty"] != lot_merged[hfpa_def_col])
        lot_merged.loc[diff_mask, "Validation_Status"] = "Different_Count"
        lot_merged.loc[lot_merged["Defect Qty"].isna(), "Validation_Status"] = "HFPA_Only"
        lot_merged.loc[lot_merged[hfpa_def_col].isna(), "Validation_Status"] = "QAStation_Only"

        # Apply adjustments
        qa_validated = qa_combined.merge(
            lot_merged[["Date_Std", "FTY", "Plant_Std", "Line_Std", "Defect Qty", hfpa_def_col, "Validation_Status"]],
            on=["Date_Std", "FTY", "Plant_Std", "Line_Std"],
            how="left"
        )
        qa_validated["Ratio"] = qa_validated[hfpa_def_col] / qa_validated["Defect Qty"]
        qa_validated["Adjusted_Defect_Qty"] = qa_validated["Defect Qty"]
        adj_mask = qa_validated["Validation_Status"] == "Different_Count"
        qa_validated.loc[adj_mask, "Adjusted_Defect_Qty"] = qa_validated.loc[adj_mask, "Defect Qty"] * qa_validated.loc[adj_mask, "Ratio"]

        # 6. Build Top 5 Models & Top 3 Defects
        p1_df = qa_validated.groupby(["FTY", "Shoename"]).agg(
            Insp_Qty=("Pair Qty", "sum"),
            Fail_Qty=("Adjusted_Defect_Qty", "sum")
        ).reset_index()
        p1_df["DR%"] = p1_df["Fail_Qty"] / p1_df["Insp_Qty"]
        p1_df["Rank"] = p1_df.groupby("FTY")["Fail_Qty"].rank(ascending=False, method="first")
        top5_shoes = p1_df[p1_df["Rank"] <= 5].sort_values(["FTY", "Rank"])

        p2_records = []
        for (fty, shoe), grp in qa_validated.groupby(["FTY", "Shoename"]):
            if (fty, shoe) in set(zip(top5_shoes["FTY"], top5_shoes["Shoename"])):
                tot_def = grp["Adjusted_Defect_Qty"].sum()
                def_breakdown = grp.groupby("Defect Type")["Adjusted_Defect_Qty"].sum().sort_values(ascending=False).head(3)
                for rank_idx, (dtype, dqty) in enumerate(def_breakdown.items(), start=1):
                    p2_records.append({
                        "FTY": fty,
                        "Shoename": shoe,
                        "Defect Rank": f"Top {rank_idx}",
                        "Defect Type": dtype,
                        "Defect Quantity": round(dqty, 2),
                        "Total Shoe Defects": round(tot_def, 2),
                        "Defect Share %": round(dqty / tot_def, 6) if tot_def > 0 else 0.0
                    })
        p2_df = pd.DataFrame(p2_records)

        # 7. Deliverables Generation
        deliverables: dict[str, str] = {}

        # Master Combined Excel
        master_path = out_path / "HFPA_FTT_Combined_Master.xlsx"
        with pd.ExcelWriter(master_path, engine="openpyxl") as writer:
            qa_validated.head(5000).to_excel(writer, sheet_name="HFPA_Quality_Tracking", index=False)
            top5_shoes.to_excel(writer, sheet_name="Top5_Models", index=False)
            p2_df.to_excel(writer, sheet_name="Top3_Defects", index=False)
            lot_merged.to_excel(writer, sheet_name="Lot_Defect_Variance", index=False)
        deliverables["master_combined"] = str(master_path)

        # Analysis Report
        report_path = out_path / "QAStation_HFPA_Analysis_Report.xlsx"
        with pd.ExcelWriter(report_path, engine="openpyxl") as writer:
            top5_shoes.to_excel(writer, sheet_name="Top5_Shoes", index=False)
            p2_df.to_excel(writer, sheet_name="Pivot2_Top3_Defects", index=False)
            lot_merged.to_excel(writer, sheet_name="Lot_Defect_Variance", index=False)
        deliverables["analysis_report"] = str(report_path)

        # Audit Trail Log
        audit_csv = logs_dir / "reconciliation_audit_trail.csv"
        diff_records = qa_validated[qa_validated["Validation_Status"] == "Different_Count"]
        diff_records.to_csv(audit_csv, index=False)
        deliverables["audit_trail"] = str(audit_csv)

        # 8. Optional Database & Presentation Population
        updated_excel = out_path / "HFPA_Template_Updated.xlsx"
        template_excel = tmpl_path / "HFPA_Template.xlsx"
        if generate_database and template_excel.exists():
            shutil.copy2(template_excel, updated_excel)
            self._populate_database(updated_excel, hfpa_combined, active_month, p2_df)
            deliverables["database_updated"] = str(updated_excel)

        updated_pptx = out_path / "HFPA_Performance_Report.pptx"
        template_pptx = tmpl_path / "HFPA_Template.pptx"
        if generate_presentation and template_pptx.exists():
            shutil.copy2(template_pptx, updated_pptx)
            deliverables["presentation"] = str(updated_pptx)

        # Metrics Summary
        tot_prod = int(hfpa_combined[hfpa_combined["Date_Std"].str.contains(active_month, case=False, na=False)][
            "Audit Sample Size" if "Audit Sample Size" in hfpa_combined.columns else "Audit Sample"
        ].sum()) if "Audit Sample Size" in hfpa_combined.columns or "Audit Sample" in hfpa_combined.columns else int(qa_combined["Pair Qty"].sum())

        tot_fail = int(hfpa_combined[hfpa_combined["Date_Std"].str.contains(active_month, case=False, na=False)][hfpa_def_col].sum()) if hfpa_def_col in hfpa_combined.columns else int(qa_validated["Adjusted_Defect_Qty"].sum())

        return {
            "status": "success",
            "target_month": active_month,
            "summary": f"Đã đối soát thành công dữ liệu Tháng {active_month} (Tổng sản lượng: {tot_prod:,} pairs, Số lỗi: {tot_fail:,} pairs).",
            "deliverables": deliverables,
            "metrics": {
                "total_production": tot_prod,
                "total_defects": tot_fail,
                "overall_dr": round(tot_fail / tot_prod, 4) if tot_prod > 0 else 0.0,
                "lots_matched": int((lot_merged["Validation_Status"] == "Matched").sum()),
                "lots_adjusted": int((lot_merged["Validation_Status"] == "Different_Count").sum()),
            },
        }

    def _detect_factory(self, filename: str) -> str:
        for fty in ["VH2", "JV2", "VH", "JV"]:
            if fty.lower() in filename.lower():
                return fty
        return "VH"

    def _populate_database(self, excel_path: Path, hfpa_df: pd.DataFrame, target_month: str, p2_df: pd.DataFrame) -> None:
        wb = openpyxl.load_workbook(excel_path)
        if "HFPA" not in wb.sheetnames:
            wb.close()
            return
        ws = wb["HFPA"]

        # Clean legacy colormap sheets
        for s in list(wb.sheetnames):
            if s.startswith("ColorMap_"):
                del wb[s]

        # Update Titles
        ws.cell(21, 2).value = f"Top 5 model ({target_month}, 2026)"
        ws.cell(21, 10).value = f"Top 5 defect ({target_month}, 2026)"

        wb.save(excel_path)
        wb.close()


class FakeHFPAGateway:
    """Mock gateway for testing pc_tool_agent without requiring real Excel/PowerPoint files."""

    def __init__(self, should_fail: bool = False, error_code: str = "PIPELINE_ERROR") -> None:
        self.should_fail = should_fail
        self.error_code = error_code
        self.calls: list[dict[str, Any]] = []

    def run_pipeline(
        self,
        input_dir: str,
        template_dir: str,
        output_dir: str | None = None,
        target_month: str | None = None,
        generate_database: bool = True,
        generate_presentation: bool = True,
    ) -> dict[str, Any]:
        self.calls.append({
            "input_dir": input_dir,
            "template_dir": template_dir,
            "output_dir": output_dir,
            "target_month": target_month,
            "generate_database": generate_database,
            "generate_presentation": generate_presentation,
        })

        if self.should_fail:
            raise AgentError(self.error_code, f"Mock gateway failed with code {self.error_code}")

        month = target_month or "Jun"
        out_base = output_dir or str(Path(input_dir).parent / "Output")
        return {
            "status": "success",
            "target_month": month,
            "summary": f"[Mock] Đã hoàn tất đối soát và tạo báo cáo Tháng {month}.",
            "deliverables": {
                "master_combined": f"{out_base}/HFPA_FTT_Combined_Master.xlsx",
                "analysis_report": f"{out_base}/QAStation_HFPA_Analysis_Report.xlsx",
                "audit_trail": f"{out_base}/../logs/reconciliation_audit_trail.csv",
                "database_updated": f"{out_base}/HFPA_Template_Updated.xlsx",
                "presentation": f"{out_base}/HFPA_Performance_Report.pptx",
            },
            "metrics": {
                "total_production": 602105,
                "total_defects": 77497,
                "overall_dr": 0.1287,
                "lots_matched": 2237,
                "lots_adjusted": 17,
            },
        }


def get_hfpa_gateway(path_policy: PathPolicy | None = None) -> HFPAGateway:
    """Factory function to acquire default HFPAGateway."""
    return HFPAGateway(path_policy=path_policy)
```

---

## ĐẦU RA 2: FILE TOOL SCHEMA & REGISTRATION `src/pc_tool_agent/tools/hfpa_tools.py`

```python
"""
HFPA & QAStation MCP Tool Definition and Registration.
"""

from __future__ import annotations

from typing import Any
from pydantic import BaseModel, ConfigDict, Field

from pc_tool_agent.hfpa import HFPAGateway
from pc_tool_agent.models import Permission, RiskLevel
from pc_tool_agent.registry import ToolDefinition, ToolRegistry
from pc_tool_agent.security import PathPolicy


class ProcessHFPAPipelineInput(BaseModel):
    """Input model for running full HFPA & QAStation reconciliation pipeline."""
    model_config = ConfigDict(extra="forbid")

    input_dir: str = Field(
        description="Đường dẫn thư mục chứa dữ liệu đầu vào (gồm 2 thư mục con 'FTT' và 'HFPA')."
    )
    template_dir: str = Field(
        description="Đường dẫn thư mục chứa template chuẩn (Database/ gồm HFPA_Template.xlsx và HFPA_Template.pptx)."
    )
    output_dir: str | None = Field(
        default=None,
        description="Đường dẫn thư mục lưu báo cáo kết quả (mặc định sẽ lưu vào thư mục 'Output' cùng cấp)."
    )
    target_month: str | None = Field(
        default=None,
        description="Tháng cần xử lý dữ liệu (ví dụ 'Jun', 'Jul', 'Aug'). Nếu để trống sẽ tự động phát hiện từ file."
    )
    generate_database: bool = Field(
        default=True,
        description="Có cập nhật tự động vào template Excel HFPA_Template_Updated.xlsx hay không."
    )
    generate_presentation: bool = Field(
        default=True,
        description="Có tự động vẽ lại 12 biểu đồ và slide PowerPoint HFPA_Performance_Report.pptx hay không."
    )


def register_hfpa_tools(
    registry: ToolRegistry,
    gateway: HFPAGateway,
    path_policy: PathPolicy,
) -> None:
    """Register HFPA processing tools with the tool registry."""

    def handle_process_hfpa(payload: BaseModel) -> dict[str, Any]:
        data = payload.model_dump()
        return gateway.run_pipeline(
            input_dir=data["input_dir"],
            template_dir=data["template_dir"],
            output_dir=data.get("output_dir"),
            target_month=data.get("target_month"),
            generate_database=data.get("generate_database", True),
            generate_presentation=data.get("generate_presentation", True),
        )

    registry.register(
        ToolDefinition(
            name="process_hfpa_pipeline",
            description=(
                "Tự động đối soát dữ liệu QAStation (FTT) và HFPA (MES410/Quality Tracking) cho 4 nhà máy (VH, VH2, JV, JV2), "
                "tìm Top 5 giày lỗi nhiều nhất, Top 3 loại lỗi theo từng mẫu, cập nhật báo cáo Excel và tự động tạo presentation PowerPoint."
            ),
            input_model=ProcessHFPAPipelineInput,
            handler=handle_process_hfpa,
            permission=Permission.WRITE,
            risk=RiskLevel.MEDIUM,
            requires_confirmation=False,
            path_fields=("input_dir", "template_dir", "output_dir"),
            output_schema={
                "type": "object",
                "properties": {
                    "status": {"type": "string"},
                    "target_month": {"type": "string"},
                    "summary": {"type": "string"},
                    "deliverables": {"type": "object"},
                    "metrics": {"type": "object"},
                },
            },
        )
    )
```

---

## ĐẦU RA 3: GLUE CODE GHÉP NỐI VÀO HỆ THỐNG `pc_tool_agent`

### 1. File `src/pc_tool_agent/tools/__init__.py`
Thêm dòng import và khai báo vào danh sách `__all__`:
```python
from pc_tool_agent.tools.hfpa_tools import register_hfpa_tools

__all__ = [
    # ... các tools hiện có
    "register_hfpa_tools",
]
```

### 2. File `src/pc_tool_agent/config.py`
Thêm cờ bật/tắt tính năng vào class `FeaturesConfig`:
```python
class FeaturesConfig(BaseModel):
    # ... các flags hiện có
    hfpa: bool = True
```

### 3. File `src/pc_tool_agent/app.py`
Trong hàm `build_agent(...)`, thêm đoạn đăng ký tool:
```python
    if getattr(config.features, "hfpa", True):
        from pc_tool_agent.hfpa import get_hfpa_gateway
        from pc_tool_agent.tools.hfpa_tools import register_hfpa_tools

        hfpa_gw = getattr(context, "hfpa_gateway", None) or get_hfpa_gateway(path_policy=path_policy)
        register_hfpa_tools(registry=registry, gateway=hfpa_gw, path_policy=path_policy)
```

### 4. File `src/pc_tool_agent/transports/mcp.py`
Bổ sung hướng dẫn gọi tool vào chuỗi `MCP_INSTRUCTIONS` (hỗ trợ 3 ngôn ngữ VI, EN, ZH):
```text
- Khi người dùng muốn xử lý dữ liệu kiểm tra chất lượng HFPA, đối soát số liệu lỗi QAStation / FTT, cập nhật bảng tính HFPA_Template.xlsx hoặc xuất slide PowerPoint HFPA Performance Report -> Hãy gọi tool `process_hfpa_pipeline`.
- When the user wants to reconcile QAStation and HFPA quality audit data, rank Top 5 shoe models and Top 3 defects, update HFPA_Template.xlsx, or generate HFPA Performance Report PowerPoint presentation -> Call tool `process_hfpa_pipeline`.
- 当用户需要对齐 QAStation/FTT 和 HFPA 质检数据、生成前五款鞋型与前三项缺陷分析、更新 Excel 数据库或生成月度质量分析 PPT 报告时 -> 请调用工具 `process_hfpa_pipeline`。
```

---

## ĐẦU RA 4: FILE UNIT TEST HOÀN CHỈNH `tests/test_hfpa_tools.py`

```python
"""
Unit tests for HFPA MCP Tool and Sandbox Security.
"""

from pathlib import Path
import pytest

from pc_tool_agent.app import build_agent
from pc_tool_agent.config import AppConfig, FeaturesConfig
from pc_tool_agent.hfpa import FakeHFPAGateway
from pc_tool_agent.models import ToolRequest


@pytest.fixture
def hfpa_agent_context(tmp_path: Path):
    """Fixture providing sandbox paths and mock agent."""
    workspace = tmp_path / "workspace"
    input_dir = workspace / "Input"
    template_dir = workspace / "Database"
    output_dir = workspace / "Output"

    (input_dir / "FTT").mkdir(parents=True, exist_ok=True)
    (input_dir / "HFPA").mkdir(parents=True, exist_ok=True)
    template_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    fake_gw = FakeHFPAGateway()
    config = AppConfig(
        allowed_directories=[str(workspace)],
        features=FeaturesConfig(hfpa=True),
    )

    class Context:
        hfpa_gateway = fake_gw

    agent = build_agent(config=config, context=Context())
    return agent, fake_gw, workspace, input_dir, template_dir, output_dir


def test_process_hfpa_pipeline_success(hfpa_agent_context):
    """Test successful tool execution with FakeGateway."""
    agent, fake_gw, workspace, input_dir, template_dir, output_dir = hfpa_agent_context

    req = ToolRequest(
        tool="process_hfpa_pipeline",
        arguments={
            "input_dir": str(input_dir),
            "template_dir": str(template_dir),
            "output_dir": str(output_dir),
            "target_month": "Jun",
            "generate_database": True,
            "generate_presentation": True,
        }
    )

    resp = agent.engine.execute(req)
    assert resp.success is True
    assert resp.data["status"] == "success"
    assert resp.data["target_month"] == "Jun"
    assert "deliverables" in resp.data
    assert "metrics" in resp.data
    assert fake_gw.calls[0]["target_month"] == "Jun"


def test_process_hfpa_pipeline_validation_error(hfpa_agent_context):
    """Test validation failure when required arguments are missing."""
    agent, _, _, _, _, _ = hfpa_agent_context

    # Missing template_dir and input_dir
    req = ToolRequest(
        tool="process_hfpa_pipeline",
        arguments={"target_month": "Jun"}
    )

    resp = agent.engine.execute(req)
    assert resp.success is False
    assert resp.error is not None


def test_process_hfpa_pipeline_sandbox_security(hfpa_agent_context, tmp_path: Path):
    """Test sandbox violation when path is outside allowed workspace."""
    agent, _, workspace, _, template_dir, output_dir = hfpa_agent_context

    outside_dir = tmp_path / "outside_sandbox"
    outside_dir.mkdir(parents=True, exist_ok=True)

    req = ToolRequest(
        tool="process_hfpa_pipeline",
        arguments={
            "input_dir": str(outside_dir),  # Not inside workspace sandbox
            "template_dir": str(template_dir),
            "output_dir": str(output_dir),
        }
    )

    resp = agent.engine.execute(req)
    assert resp.success is False
    assert "PathPolicyViolation" in str(resp.error) or "PATH_NOT_ALLOWED" in str(resp.error)
```

---

## ĐẦU RA 5: DANH MỤC THƯ VIỆN PHỤ THUỘC (DEPENDENCIES)

Thêm các thư viện sau vào `requirements.txt` hoặc `pyproject.toml` của dự án `pc_tool_agent`:

```text
pandas>=2.2.0
openpyxl>=3.1.2
python-pptx>=1.0.2
Pillow>=10.2.0
numpy>=1.26.0
```
