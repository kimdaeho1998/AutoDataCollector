from __future__ import annotations

from dataclasses import dataclass
from datetime import date, timedelta
from enum import Enum
from pathlib import Path
from typing import Protocol

from openpyxl import load_workbook

from .mapping.template_resolver import CellType, SalesAdminTemplateResolver
from .models import SalesAdminDailyRecord, Store, TodayStoreSalesResult


class ProductionSalesClient(Protocol):
    def get_today_store_sales(
        self,
        *,
        business_date: date,
        brand_idx: str,
        store_idx: str,
        store_name: str,
    ) -> TodayStoreSalesResult: ...


def default_target_date(run_date: date | None = None) -> date:
    """Production collection always targets the previous business calendar day."""
    return (run_date or date.today()) - timedelta(days=1)


class SingleDaySalesCollector:
    """Collect only the confirmed production metrics for one store and one date."""

    def __init__(self, client: ProductionSalesClient, *, brand_idx: str, brand_name: str) -> None:
        self.client = client
        self.brand_idx = brand_idx
        self.brand_name = brand_name

    def collect(self, *, store: Store, business_date: date) -> SalesAdminDailyRecord:
        today_sales = self.client.get_today_store_sales(
            business_date=business_date,
            brand_idx=self.brand_idx,
            store_idx=store.magic_store_id,
            store_name=store.store_name,
        )
        if (
            isinstance(today_sales.receipt_count, int)
            and today_sales.receipt_count < 0
        ) or (
            isinstance(today_sales.gross_sales_amount, int)
            and today_sales.gross_sales_amount < 0
        ):
            raise ValueError("SOURCE_VALUE_INVALID")
        return SalesAdminDailyRecord(
            business_date=business_date,
            store_id=store.magic_store_id,
            store_name=store.store_name,
            receipt_count=today_sales.receipt_count,
            gross_sales_amount=today_sales.gross_sales_amount,
        )


class DryRunStatus(str, Enum):
    READY = "READY"
    SAME_VALUE = "SAME_VALUE"
    CONFLICT = "CONFLICT"
    BLOCKED_FORMULA = "BLOCKED_FORMULA"
    RESOLUTION_FAILED = "RESOLUTION_FAILED"


@dataclass(frozen=True)
class PlannedChange:
    sheet: str
    cell: str
    old_value: object
    new_value: int | str
    metric: str


@dataclass(frozen=True)
class SingleDayDryRunResult:
    record: SalesAdminDailyRecord
    status: DryRunStatus
    changes: tuple[PlannedChange, ...]
    reason: str | None = None


class SalesAdminDryRun:
    """Resolve an atomic two-cell write without changing the original workbook."""

    def __init__(self, template_path: str | Path) -> None:
        self.template_path = Path(template_path)

    def preview(self, record: SalesAdminDailyRecord) -> SingleDayDryRunResult:
        workbook = load_workbook(self.template_path, data_only=False)
        try:
            targets = SalesAdminTemplateResolver(workbook).resolve_target_cells(
                record.business_date, record.store_id, record.store_name
            )
        except ValueError as exc:
            return SingleDayDryRunResult(record, DryRunStatus.RESOLUTION_FAILED, (), str(exc))

        sheet = workbook[targets.sheet_name]
        changes = (
            PlannedChange(sheet.title, targets.receipt_cell, sheet[targets.receipt_cell].value, record.receipt_count, "receipt_count"),
            PlannedChange(sheet.title, targets.sales_cell, sheet[targets.sales_cell].value, record.gross_sales_amount, "gross_sales_amount"),
        )
        if any(SalesAdminTemplateResolver.classify_cell(sheet[change.cell]) == CellType.FORMULA for change in changes):
            return SingleDayDryRunResult(record, DryRunStatus.BLOCKED_FORMULA, changes, "TARGET_CELL_IS_FORMULA")
        if all(change.old_value is None for change in changes):
            return SingleDayDryRunResult(record, DryRunStatus.READY, changes)
        if all(change.old_value == change.new_value for change in changes):
            return SingleDayDryRunResult(record, DryRunStatus.SAME_VALUE, changes)
        # ERP is the source of truth for input cells, so a later run replaces
        # prior manually entered or previously collected values.
        return SingleDayDryRunResult(record, DryRunStatus.READY, changes)
