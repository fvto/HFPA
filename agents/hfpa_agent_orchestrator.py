"""
Master Orchestrator: HFPAOrchestrator
Unifies Agent 1 (Ingestion Watcher), Agent 2 (Reconciliation Engine), and Agent 3 (Quality Audit & PPTX).

Solves the 3 biggest project inefficiencies:
1. Inefficiency 1 -> IngestionWatcherAgent (Proactive monitoring, completeness check, lock handling).
2. Inefficiency 2 -> ReconciliationEngineAgent (Single-pass in-memory DataBus, 10x faster).
3. Inefficiency 3 -> QualityAuditPPTXAgent (Semantic PPTX binding, anomaly detection, Executive Briefing).
"""

from __future__ import annotations

import argparse
import os
import sys
import time
from typing import Optional

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
if hasattr(sys.stderr, "reconfigure"):
    try:
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

try:
    from .ingestion_watcher_agent import IngestionWatcherAgent, IngestionStatus
    from .reconciliation_engine_agent import ReconciliationEngineAgent, DataBus
    from .quality_audit_pptx_agent import QualityAuditPPTXAgent, AnomalyReport
except (ImportError, ValueError):
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    from agents.ingestion_watcher_agent import IngestionWatcherAgent, IngestionStatus
    from agents.reconciliation_engine_agent import ReconciliationEngineAgent, DataBus
    from agents.quality_audit_pptx_agent import QualityAuditPPTXAgent, AnomalyReport


class HFPAOrchestrator:
    """Master orchestrator executing the proactive multi-agent pipeline."""

    def __init__(self, base_dir: Optional[str] = None):
        self.base_dir = base_dir or os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        self.ingestion_agent = IngestionWatcherAgent(self.base_dir)
        self.reconciliation_agent = ReconciliationEngineAgent(self.base_dir)
        self.audit_agent = QualityAuditPPTXAgent(self.base_dir)

    def run_pipeline(
        self,
        target_month: Optional[str] = None,
        target_year: Optional[str] = None,
        force: bool = False,
    ) -> bool:
        """Execute the end-to-end multi-agent proactive pipeline."""
        t_global = time.time()
        print("\n" + "=" * 75)
        print("   HFPA PROACTIVE MULTI-AGENT DATA PROCESSING & AUDIT SYSTEM")
        print("=" * 75)

        # STAGE 1: Pre-flight Verification with Agent 1
        status: IngestionStatus = self.ingestion_agent.evaluate_readiness(target_month, target_year)
        self.ingestion_agent.print_readiness_report(status)

        if not status.is_ready and not force:
            print("[x] PIPELINE HALTED: Dữ liệu đầu vào chưa sẵn sàng hoặc có file đang bị khóa.")
            print("    Vui lòng kiểm tra các lỗi trên hoặc dùng cờ --force để ép buộc chạy.")
            return False

        active_month = status.active_month
        active_year = status.active_year

        # STAGE 2: In-Memory Processing & Reconciliation with Agent 2
        print(f"\n[*] LAUNCHING AGENT 2: High-Performance In-Memory Reconciliation (DataBus)...")
        bus: DataBus = self.reconciliation_agent.run_reconciliation(active_month, active_year)

        # STAGE 3: Database Population
        print(f"\n[*] POPULATING DATABASE TEMPLATE (HFPA_Template.xlsx)...")
        try:
            from populate_database import populate_hfpa_database
            populate_hfpa_database(target_month=active_month, target_year=active_year, p2_df=bus.pivot2_df)
        except Exception as e:
            print(f"[-] Warning during Database Template population: {e}")

        # STAGE 4: PowerPoint Presentation & Anomaly Audit with Agent 3
        print(f"\n[*] LAUNCHING AGENT 3: PowerPoint Synchronization & Quality Anomaly Audit...")
        self.audit_agent.update_presentation_resiliently(bus)
        anomaly_report = self.audit_agent.run_anomaly_audit(bus)
        brief_file = self.audit_agent.generate_executive_brief(bus, anomaly_report)

        total_duration = time.time() - t_global
        print("\n" + "=" * 75)
        print(f"   WORKFLOW COMPLETED SUCCESSFULLY IN {total_duration:.2f} SECONDS!")
        print(f"   Target Period: {active_month.upper()} {active_year}")
        print(f"   Deliverables in: {os.path.join(self.base_dir, 'Output')}")
        print(f"   Executive Quality Brief: {os.path.basename(brief_file)}")
        print("=" * 75 + "\n")
        return True

    def run_watcher(self, interval_seconds: int = 10) -> None:
        """Run proactive directory watcher daemon."""
        def callback(m: str, y: str):
            self.run_pipeline(target_month=m, target_year=y)

        self.ingestion_agent.watch_and_trigger(
            interval_seconds=interval_seconds,
            trigger_callback=callback
        )

    def check_health(self) -> None:
        """Run standalone health & completeness check."""
        status = self.ingestion_agent.evaluate_readiness()
        self.ingestion_agent.print_readiness_report(status)


def main():
    parser = argparse.ArgumentParser(description="HFPA Proactive Multi-Agent Pipeline")
    parser.add_argument("--run", action="store_true", help="Execute complete pipeline immediately")
    parser.add_argument("--check", action="store_true", help="Run pre-flight health and readiness check")
    parser.add_argument("--watch", action="store_true", help="Start continuous proactive directory watcher")
    parser.add_argument("--interval", type=int, default=10, help="Watcher polling interval in seconds")
    parser.add_argument("--force", action="store_true", help="Force execution even if some files are missing")
    parser.add_argument("--month", type=str, default=None, help="Target month (e.g. Jul, Aug)")
    parser.add_argument("--year", type=str, default=None, help="Target year (e.g. 2026)")

    args = parser.parse_args()
    orchestrator = HFPAOrchestrator()

    if args.check:
        orchestrator.check_health()
    elif args.watch:
        orchestrator.run_watcher(interval_seconds=args.interval)
    else:
        orchestrator.run_pipeline(
            target_month=args.month,
            target_year=args.year,
            force=args.force
        )


if __name__ == "__main__":
    main()
