"""
Agent 1: IngestionWatcherAgent
Proactive File Ingestion, Non-Blocking Lock Detection, and Completeness Validation.

Addresses Inefficiency 1:
- Eliminates manual batch triggering and human babysitting.
- Prevents premature, destructive file archiving on disk.
- Replaces blocking console input prompts with non-blocking file-lock diagnostics.
- Verifies 4-factory completeness (VH, VH2, JV, JV2) across both FTT and HFPA streams.
"""

from __future__ import annotations

import glob
import os
import re
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

import sys
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

import pandas as pd

MONTH_MAP = {
    1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"
}
MONTH_REVERSE_MAP = {v.lower(): k for k, v in MONTH_MAP.items()}
EXPECTED_FACTORIES = ["VH", "VH2", "JV", "JV2"]


@dataclass
class FileMetadata:
    path: str
    filename: str
    file_type: str  # 'FTT' or 'HFPA'
    factory: str
    detected_month: str
    detected_year: str
    size_bytes: int
    modified_time: float
    is_locked: bool = False
    lock_reason: str = ""


@dataclass
class IngestionStatus:
    is_ready: bool
    active_month: str
    active_year: str
    present_ftt: List[str] = field(default_factory=list)
    missing_ftt: List[str] = field(default_factory=list)
    present_hfpa: List[str] = field(default_factory=list)
    missing_hfpa: List[str] = field(default_factory=list)
    locked_files: List[str] = field(default_factory=list)
    files: List[FileMetadata] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)


class IngestionWatcherAgent:
    """
    Proactive agent monitoring input directories for data readiness,
    validating file integrity, detecting locks non-blockingly, and preparing staged datasets.
    """

    def __init__(self, base_dir: Optional[str] = None):
        self.base_dir = base_dir or os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        self.input_ftt_dir = os.path.join(self.base_dir, "Input", "FTT")
        self.input_hfpa_dir = os.path.join(self.base_dir, "Input", "HFPA")
        self.output_dir = os.path.join(self.base_dir, "Output")
        self.database_dir = os.path.join(self.base_dir, "Database")

    @staticmethod
    def extract_factory(filename: str) -> str:
        """Extract factory identifier (JV, JV2, VH, VH2) from filename."""
        base = os.path.splitext(os.path.basename(filename))[0]
        for fty in ["JV2", "VH2", "JV", "VH"]:
            if re.search(rf"[-_\s]{fty}$", base, re.IGNORECASE) or re.search(rf"[-_\s]{fty}[-_\s]", base, re.IGNORECASE):
                return fty
        for fty in ["JV2", "VH2", "JV", "VH"]:
            if fty in base.upper():
                return fty
        return "UNKNOWN"

    @staticmethod
    def probe_file_lock(filepath: str) -> Tuple[bool, str]:
        """
        Non-blocking lock check. Returns (is_locked, reason).
        Never prompts the user or hangs the process.
        """
        if not os.path.exists(filepath):
            return False, "File does not exist"
        try:
            with open(filepath, "r+b"):
                return False, ""
        except PermissionError:
            return True, "File is locked by another process (e.g. Microsoft Excel / WPS Office)"
        except Exception as e:
            return True, f"File access error: {str(e)}"

    def detect_file_month_and_year(self, filepath: str) -> Tuple[str, str]:
        """Inspect file header dates and filename to determine month and year."""
        detected_month = None
        detected_year = None

        # 1. Inspect first rows of date columns
        try:
            df_sample = pd.read_excel(filepath, nrows=30)
            date_cols = [c for c in df_sample.columns if "date" in str(c).lower()]
            for c in date_cols:
                dates = pd.to_datetime(df_sample[c], errors="coerce").dropna()
                if not dates.empty:
                    first_date = dates.iloc[0]
                    detected_month = MONTH_MAP.get(first_date.month)
                    detected_year = str(first_date.year)
                    if detected_month and detected_year:
                        return detected_month, detected_year
        except Exception:
            pass

        # 2. Filename patterns
        filename = os.path.basename(filepath)
        m = re.search(r"(\d{4})[-_](\d{2})[-_](\d{2})", filename)
        if m:
            detected_year = m.group(1)
            month_num = int(m.group(2))
            if 1 <= month_num <= 12:
                detected_month = MONTH_MAP.get(month_num)

        for m_num, m_name in MONTH_MAP.items():
            if re.search(rf"[-_\s]{m_name}[-_\s\.]", filename, re.IGNORECASE):
                detected_month = m_name
                break

        y_m = re.search(r"(20\d{2})", filename)
        if y_m:
            detected_year = y_m.group(1)

        return detected_month or "Jul", detected_year or "2026"

    def scan_files(self) -> List[FileMetadata]:
        """Scan all FTT and HFPA files and build metadata catalog."""
        metadata_list = []

        # Scan FTT
        ftt_patterns = [
            os.path.join(self.input_ftt_dir, "*.xlsx"),
            os.path.join(self.input_ftt_dir, "*.xls"),
        ]
        for pattern in ftt_patterns:
            for fp in glob.glob(pattern):
                fname = os.path.basename(fp)
                if fname.startswith("~$") or fname.startswith("."):
                    continue
                fty = self.extract_factory(fname)
                m, y = self.detect_file_month_and_year(fp)
                locked, reason = self.probe_file_lock(fp)
                stat = os.stat(fp)
                metadata_list.append(
                    FileMetadata(
                        path=fp,
                        filename=fname,
                        file_type="FTT",
                        factory=fty,
                        detected_month=m,
                        detected_year=y,
                        size_bytes=stat.st_size,
                        modified_time=stat.st_mtime,
                        is_locked=locked,
                        lock_reason=reason,
                    )
                )

        # Scan HFPA
        hfpa_patterns = [
            os.path.join(self.input_hfpa_dir, "*.xlsx"),
            os.path.join(self.input_hfpa_dir, "*.xls"),
        ]
        for pattern in hfpa_patterns:
            for fp in glob.glob(pattern):
                fname = os.path.basename(fp)
                if fname.startswith("~$") or fname.startswith("."):
                    continue
                fty = self.extract_factory(fname)
                m, y = self.detect_file_month_and_year(fp)
                locked, reason = self.probe_file_lock(fp)
                stat = os.stat(fp)
                metadata_list.append(
                    FileMetadata(
                        path=fp,
                        filename=fname,
                        file_type="HFPA",
                        factory=fty,
                        detected_month=m,
                        detected_year=y,
                        size_bytes=stat.st_size,
                        modified_time=stat.st_mtime,
                        is_locked=locked,
                        lock_reason=reason,
                    )
                )

        return metadata_list

    def evaluate_readiness(
        self, target_month: Optional[str] = None, target_year: Optional[str] = None
    ) -> IngestionStatus:
        """
        Evaluate full 4-factory completeness and readiness for pipeline processing.
        """
        all_files = self.scan_files()
        warnings: List[str] = []
        errors: List[str] = []

        if not all_files:
            return IngestionStatus(
                is_ready=False,
                active_month=target_month or "Unknown",
                active_year=target_year or "Unknown",
                missing_ftt=list(EXPECTED_FACTORIES),
                missing_hfpa=list(EXPECTED_FACTORIES),
                errors=["Không tìm thấy file nào trong thư mục Input/FTT hoặc Input/HFPA!"],
            )

        # Detect active month if not supplied
        if not target_month or not target_year:
            # Frequency vote for month and year
            month_counts: Dict[str, int] = {}
            year_counts: Dict[str, int] = {}
            for f in all_files:
                month_counts[f.detected_month] = month_counts.get(f.detected_month, 0) + 1
                year_counts[f.detected_year] = year_counts.get(f.detected_year, 0) + 1

            target_month = max(month_counts.items(), key=lambda x: x[1])[0]
            target_year = max(year_counts.items(), key=lambda x: x[1])[0]

        # Filter files for active month
        active_files = [
            f for f in all_files
            if f.detected_month.lower() == target_month.lower()
            and (not target_year or f.detected_year == target_year)
        ]

        # Check for files belonging to other months (flagged as non-destructive warning)
        other_month_files = [f for f in all_files if f not in active_files]
        if other_month_files:
            other_names = [f"{f.filename} ({f.detected_month}/{f.detected_year})" for f in other_month_files]
            warnings.append(
                f"Phát hiện {len(other_month_files)} file thuộc tháng khác trong thư mục Input: {', '.join(other_names)}"
            )

        # Check locked files
        locked_files = [f.filename for f in active_files if f.is_locked]
        if locked_files:
            errors.append(
                f"Có {len(locked_files)} file đang bị khóa bởi Excel/Office: {', '.join(locked_files)}"
            )

        # Check FTT completeness
        ftt_by_fty = {f.factory: f for f in active_files if f.file_type == "FTT"}
        present_ftt = [fty for fty in EXPECTED_FACTORIES if fty in ftt_by_fty]
        missing_ftt = [fty for fty in EXPECTED_FACTORIES if fty not in ftt_by_fty]

        # Check HFPA completeness
        hfpa_by_fty = {f.factory: f for f in active_files if f.file_type == "HFPA"}
        present_hfpa = [fty for fty in EXPECTED_FACTORIES if fty in hfpa_by_fty]
        missing_hfpa = [fty for fty in EXPECTED_FACTORIES if fty not in hfpa_by_fty]

        if missing_ftt:
            errors.append(f"Thiếu dữ liệu QAStation/FTT cho nhà máy: {', '.join(missing_ftt)}")
        if missing_hfpa:
            errors.append(f"Thiếu dữ liệu MES410/HFPA cho nhà máy: {', '.join(missing_hfpa)}")

        # Check output destinations lock status
        output_master = os.path.join(self.output_dir, f"HFPA_FTT_Combined_Master_{target_month}_{target_year}.xlsx")
        output_report = os.path.join(self.output_dir, f"QAStation_HFPA_Analysis_Report_{target_month}_{target_year}.xlsx")
        db_template = os.path.join(self.database_dir, "HFPA_Template.xlsx")
        pptx_template = os.path.join(self.database_dir, "HFPA_Template.pptx")

        for dest_path, desc in [
            (output_master, "Master Excel"),
            (output_report, "Analysis Report"),
            (db_template, "Database Template"),
            (pptx_template, "PowerPoint Template"),
        ]:
            if os.path.exists(dest_path):
                is_lock, msg = self.probe_file_lock(dest_path)
                if is_lock:
                    errors.append(f"{desc} ('{os.path.basename(dest_path)}') đang mở trong Excel/PowerPoint. Vui lòng đóng trước khi chạy!")
                    locked_files.append(os.path.basename(dest_path))

        is_ready = len(missing_ftt) == 0 and len(missing_hfpa) == 0 and len(locked_files) == 0

        return IngestionStatus(
            is_ready=is_ready,
            active_month=target_month,
            active_year=target_year,
            present_ftt=present_ftt,
            missing_ftt=missing_ftt,
            present_hfpa=present_hfpa,
            missing_hfpa=missing_hfpa,
            locked_files=locked_files,
            files=active_files,
            warnings=warnings,
            errors=errors,
        )

    def print_readiness_report(self, status: IngestionStatus) -> None:
        """Render a formatted terminal readiness report."""
        print("=" * 70)
        print(f" [AGENT 1: INGESTION WATCHER] Data Readiness Report")
        print(f" Target Period: {status.active_month.upper()} {status.active_year}")
        print("=" * 70)

        print("\n [1] Factory Dataset Completeness Matrix (4/4 Target):")
        for fty in EXPECTED_FACTORIES:
            ftt_ok = "[OK]" if fty in status.present_ftt else "[MISSING]"
            hfpa_ok = "[OK]" if fty in status.present_hfpa else "[MISSING]"
            print(f"     * Factory {fty:4s} -> QAStation/FTT: {ftt_ok:10s} | MES410/HFPA: {hfpa_ok}")

        if status.locked_files:
            print("\n [!] File Lock Alerts (Action Required):")
            for f in status.locked_files:
                print(f"     - LOCKED: {f}")

        if status.warnings:
            print("\n [i] Warnings:")
            for w in status.warnings:
                print(f"     - {w}")

        if status.errors:
            print("\n [x] Errors / Incomplete streams:")
            for e in status.errors:
                print(f"     - {e}")

        status_text = "[READY TO PROCESS]" if status.is_ready else "[BLOCKED / INCOMPLETE]"
        print("\n" + "-" * 70)
        print(f" OVERALL INGESTION STATUS: {status_text}")
        print("=" * 70 + "\n")

    def watch_and_trigger(
        self,
        interval_seconds: int = 10,
        trigger_callback: Optional[Callable[[str, str], Any]] = None,
        max_checks: Optional[int] = None,
    ) -> None:
        """
        Proactively watch input directory in daemon mode, triggering pipeline execution
        as soon as all 4 factories have submitted complete data.
        """
        print(f"[*] Starting IngestionWatcherAgent in proactive watch mode (poll every {interval_seconds}s)...")
        checks = 0
        last_ready_month = None

        while True:
            checks += 1
            status = self.evaluate_readiness()

            if status.is_ready:
                combo_key = f"{status.active_month}_{status.active_year}"
                if combo_key != last_ready_month:
                    print(f"\n[!] PROACTIVE TRIGGER: Complete 4-factory dataset detected for {combo_key}!")
                    self.print_readiness_report(status)
                    if trigger_callback:
                        print(f"[*] Dispatching trigger callback to ReconciliationEngineAgent...")
                        trigger_callback(status.active_month, status.active_year)
                    last_ready_month = combo_key
            else:
                timestamp = time.strftime("%H:%M:%S")
                missing_str = f"FTT missing: {status.missing_ftt}, HFPA missing: {status.missing_hfpa}"
                print(f"[{timestamp}] Watching... {missing_str} (Locked: {len(status.locked_files)})", end="\r")

            if max_checks and checks >= max_checks:
                break
            time.sleep(interval_seconds)
