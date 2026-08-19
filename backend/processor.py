from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd


WORK_GROUPS = ("inbound", "pick", "outbound")


@dataclass(frozen=True)
class WorkGroupSummary:
    work_group: str
    total: int
    completed: int
    pending: int
    in_progress: int

    @property
    def progress(self) -> float:
        if self.total <= 0:
            return 0
        return round((self.completed / self.total) * 100, 2)


def normalize_status(value: Any) -> str:
    text = str(value or "").strip().lower()
    if text in {"complete", "completed", "done", "finish", "finished", "เสร็จ", "สำเร็จ"}:
        return "completed"
    if text in {"progress", "in progress", "processing", "working", "กำลังทำ"}:
        return "in_progress"
    if text in {"cancel", "cancelled", "canceled", "ยกเลิก"}:
        return "cancelled"
    return "pending"


def read_workbook(path: str | Path) -> dict[str, pd.DataFrame]:
    workbook_path = Path(path)
    if not workbook_path.exists():
        raise FileNotFoundError(f"Excel file not found: {workbook_path}")

    sheets = pd.read_excel(workbook_path, sheet_name=None, engine="openpyxl")
    return {name.strip().lower(): frame.fillna("") for name, frame in sheets.items()}


def detect_work_group(sheet_name: str) -> str | None:
    normalized = sheet_name.lower()
    for group in WORK_GROUPS:
        if group in normalized:
            return group
    return None


def summarize_frame(work_group: str, frame: pd.DataFrame) -> WorkGroupSummary:
    if frame.empty:
        return WorkGroupSummary(work_group, 0, 0, 0, 0)

    status_column = next((col for col in frame.columns if str(col).strip().lower() in {"status", "สถานะ"}), None)
    if status_column is None:
        total = len(frame.index)
        return WorkGroupSummary(work_group, total, 0, total, 0)

    statuses = frame[status_column].map(normalize_status)
    total = len(statuses)
    completed = int((statuses == "completed").sum())
    in_progress = int((statuses == "in_progress").sum())
    pending = int((statuses == "pending").sum())
    return WorkGroupSummary(work_group, total, completed, pending, in_progress)


def process_excel_file(path: str | Path) -> dict[str, Any]:
    sheets = read_workbook(path)
    summaries: dict[str, WorkGroupSummary] = {
        group: WorkGroupSummary(group, 0, 0, 0, 0)
        for group in WORK_GROUPS
    }

    for sheet_name, frame in sheets.items():
        work_group = detect_work_group(sheet_name)
        if work_group is None:
            continue
        summaries[work_group] = summarize_frame(work_group, frame)

    return {
        "work_groups": {
            key: {
                "total": summary.total,
                "completed": summary.completed,
                "pending": summary.pending,
                "in_progress": summary.in_progress,
                "progress": summary.progress,
            }
            for key, summary in summaries.items()
        },
        "sheet_names": list(sheets.keys()),
    }


def generate_mock_dashboard() -> dict[str, Any]:
    return {
        "work_groups": {
            "inbound": {"total": 128, "completed": 94, "pending": 18, "in_progress": 16, "progress": 73.44},
            "pick": {"total": 246, "completed": 171, "pending": 52, "in_progress": 23, "progress": 69.51},
            "outbound": {"total": 86, "completed": 61, "pending": 12, "in_progress": 13, "progress": 70.93},
        },
        "sheet_names": ["Inbound", "Pick", "Outbound"],
    }
