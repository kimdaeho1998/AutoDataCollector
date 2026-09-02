from __future__ import annotations

import re
from dataclasses import dataclass
from enum import Enum

from openpyxl.utils import get_column_letter

from ..mapping.menu_mapping import MenuMappingPreview, MenuMappingStatus
from .menu_template_profile import MenuTemplateProfile
from .menu_template_resolver import (
    ExcelDisposition,
    MenuExcelTargetPreview,
    MenuTargetStatus,
    build_menu_excel_target_preview,
    reconcile_other_residual,
    resolve_menu_target,
    resolve_store_sales_row,
)


class CellPlanStatus(str, Enum):
    READY = "READY"
    SAME_VALUE = "SAME_VALUE"
    CONFLICT = "CONFLICT"
    FORMULA_PROTECTED = "FORMULA_PROTECTED"
    TARGET_HEADER_MISSING = "TARGET_HEADER_MISSING"
    TARGET_HEADER_DUPLICATE = "TARGET_HEADER_DUPLICATE"
    STORE_NOT_FOUND = "STORE_NOT_FOUND"
    STORE_DUPLICATE = "STORE_DUPLICATE"
    NO_SOURCE_VALUE = "NO_SOURCE_VALUE"
    VALIDATE_ONLY = "VALIDATE_ONLY"
    BLOCKED = "BLOCKED"


@dataclass(frozen=True)
class MenuExcelDryRunCell:
    store_name: str
    store_row: int
    canonical_code: str
    excel_group: str | None
    excel_header: str | None
    target_column: str | None
    target_cell: str | None
    current_value: object
    proposed_value: int
    current_formula: str | None
    status: CellPlanStatus
    reason: str


@dataclass(frozen=True)
class MenuExcelAcPlan:
    store_name: str
    store_row: int
    target_column: str
    target_cell: str
    current_value: object
    proposed_value: int | None
    current_formula: str | None
    status: CellPlanStatus
    reason: str


@dataclass(frozen=True)
class MenuExcelAbValidation:
    store_name: str
    store_row: int
    target_column: str
    target_cell: str
    current_formula: str | None
    status: CellPlanStatus
    formula_valid: bool
    reason: str


@dataclass(frozen=True)
class MenuExcelDryRunPlan:
    store_name: str
    store_row: int
    source: MenuMappingPreview
    disposition: MenuExcelTargetPreview
    cells: tuple[MenuExcelDryRunCell, ...]
    ac_plan: MenuExcelAcPlan
    ab_validation: MenuExcelAbValidation
    source_total_sales: int | None
    direct_target_sales: int
    source_other_residual: int
    calculated_other_residual: int | None
    residual_match: bool

    def count_cells(self, status: CellPlanStatus) -> int:
        return sum(1 for cell in self.cells if cell.status == status)

    @property
    def ab_write_plan_count(self) -> int:
        return 0

    @property
    def ad_write_plan_count(self) -> int:
        return 0


def build_menu_excel_dry_run_plan(
    mapping_preview: MenuMappingPreview,
    worksheet,
    profile: MenuTemplateProfile,
    store_name: str,
) -> MenuExcelDryRunPlan:
    store_row = resolve_store_sales_row(worksheet, profile, store_name)
    disposition = build_menu_excel_target_preview(mapping_preview, worksheet, profile)
    if store_row.status != MenuTargetStatus.TARGET_RESOLVED:
        ac_plan = _blocked_ac_plan(store_name, store_row.row_index, store_row.status.value)
        ab_validation = _blocked_ab_validation(store_name, store_row.row_index, store_row.status.value)
        return MenuExcelDryRunPlan(
            store_name=store_name,
            store_row=store_row.row_index,
            source=mapping_preview,
            disposition=disposition,
            cells=(),
            ac_plan=ac_plan,
            ab_validation=ab_validation,
            source_total_sales=mapping_preview.source.source_total_sales,
            direct_target_sales=disposition.direct_target_sales,
            source_other_residual=disposition.other_residual_sales,
            calculated_other_residual=None,
            residual_match=False,
        )

    cells = tuple(_direct_cell_plan(mapping_preview, worksheet, profile, store_name, store_row.row_index))
    ac_plan = _ac_plan(mapping_preview, worksheet, store_name, store_row.row_index)
    ab_validation = _ab_validation(worksheet, store_name, store_row.row_index)
    reconciliation = reconcile_other_residual(disposition)
    calculated_other = reconciliation.expected_other_residual if reconciliation.source_total_sales else None
    return MenuExcelDryRunPlan(
        store_name=store_name,
        store_row=store_row.row_index,
        source=mapping_preview,
        disposition=disposition,
        cells=cells,
        ac_plan=ac_plan,
        ab_validation=ab_validation,
        source_total_sales=mapping_preview.source.source_total_sales,
        direct_target_sales=disposition.direct_target_sales,
        source_other_residual=disposition.other_residual_sales,
        calculated_other_residual=calculated_other,
        residual_match=reconciliation.is_pass,
    )


def _direct_cell_plan(mapping_preview: MenuMappingPreview, worksheet, profile: MenuTemplateProfile, store_name: str, row_index: int):
    for aggregate in mapping_preview.aggregates:
        resolved = resolve_menu_target(worksheet, profile, aggregate.canonical_code)
        if resolved.status != MenuTargetStatus.TARGET_RESOLVED or not resolved.column_letter:
            continue
        current_value = worksheet[f"{resolved.column_letter}{row_index}"].value
        current_formula = current_value if isinstance(current_value, str) and current_value.startswith("=") else None
        status, reason = _cell_status(current_value, aggregate.sales_amount)
        yield MenuExcelDryRunCell(
            store_name=store_name,
            store_row=row_index,
            canonical_code=aggregate.canonical_code,
            excel_group=resolved.group_header,
            excel_header=resolved.menu_header,
            target_column=resolved.column_letter,
            target_cell=f"{resolved.column_letter}{row_index}",
            current_value=current_value,
            proposed_value=aggregate.sales_amount,
            current_formula=current_formula,
            status=status,
            reason=reason,
        )


def _ac_plan(mapping_preview: MenuMappingPreview, worksheet, store_name: str, row_index: int) -> MenuExcelAcPlan:
    current_value = worksheet[f"AC{row_index}"].value
    current_formula = current_value if isinstance(current_value, str) and current_value.startswith("=") else None
    proposed = mapping_preview.source.source_total_sales
    if proposed is None:
        status = CellPlanStatus.BLOCKED
        reason = "SOURCE_TOTAL_MISSING"
    else:
        status, reason = _cell_status(current_value, proposed)
    return MenuExcelAcPlan(
        store_name=store_name,
        store_row=row_index,
        target_column="AC",
        target_cell=f"AC{row_index}",
        current_value=current_value,
        proposed_value=proposed,
        current_formula=current_formula,
        status=status,
        reason=reason,
    )


def _ab_validation(worksheet, store_name: str, row_index: int) -> MenuExcelAbValidation:
    cell = f"AB{row_index}"
    formula = worksheet[cell].value
    formula_valid = _is_ab_residual_formula(formula, row_index)
    return MenuExcelAbValidation(
        store_name=store_name,
        store_row=row_index,
        target_column="AB",
        target_cell=cell,
        current_formula=formula if isinstance(formula, str) and formula.startswith("=") else None,
        status=CellPlanStatus.VALIDATE_ONLY if formula_valid else CellPlanStatus.BLOCKED,
        formula_valid=formula_valid,
        reason="AB_FORMULA_VALID" if formula_valid else "AB_FORMULA_MISSING_OR_UNEXPECTED",
    )


def _cell_status(current_value, proposed_value: int) -> tuple[CellPlanStatus, str]:
    if isinstance(current_value, str) and current_value.startswith("="):
        return CellPlanStatus.FORMULA_PROTECTED, "TARGET_CELL_HAS_FORMULA"
    if current_value in (None, ""):
        return CellPlanStatus.READY, "BLANK_CELL"
    if _numeric_value(current_value) == proposed_value:
        return CellPlanStatus.SAME_VALUE, "CURRENT_EQUALS_PROPOSED"
    if _numeric_value(current_value) is not None:
        return CellPlanStatus.CONFLICT, "CURRENT_STATIC_VALUE_DIFFERS"
    return CellPlanStatus.CONFLICT, "CURRENT_NON_NUMERIC_VALUE_DIFFERS"


def _numeric_value(value) -> int | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return int(value)
    if isinstance(value, str):
        cleaned = value.replace(",", "").strip()
        if re.fullmatch(r"-?\d+", cleaned):
            return int(cleaned)
    return None


def _is_ab_residual_formula(value, row_index: int) -> bool:
    if not isinstance(value, str) or not value.startswith("="):
        return False
    formula = value.replace(" ", "").replace("$", "").upper()
    expected = f"=AC{row_index}-SUM(G{row_index}:AA{row_index})"
    return formula == expected


def _blocked_ac_plan(store_name: str, row_index: int, reason: str) -> MenuExcelAcPlan:
    row = row_index or 0
    return MenuExcelAcPlan(store_name, row, "AC", None if row == 0 else f"AC{row}", None, None, None, CellPlanStatus.BLOCKED, reason)


def _blocked_ab_validation(store_name: str, row_index: int, reason: str) -> MenuExcelAbValidation:
    row = row_index or 0
    return MenuExcelAbValidation(store_name, row, "AB", None if row == 0 else f"AB{row}", None, CellPlanStatus.BLOCKED, False, reason)


def summarize_other_residual_by_reason(plan: MenuExcelDryRunPlan) -> dict[str, int]:
    buckets = {
        "GENERAL_DRINK": 0,
        "UNSPECIFIED_TTEOKBOKKI": 0,
        "MILKIT": 0,
        "DELIVERY_FEE": 0,
        "OTHER_NEW_MENU": 0,
    }
    for item in plan.disposition.items:
        if item.disposition != ExcelDisposition.OTHER_RESIDUAL:
            continue
        name = item.mapping.normalized_name
        amount = item.mapping.record.sales_amount
        if item.mapping.canonical_code == "DRINK":
            buckets["GENERAL_DRINK"] += amount
        elif "떡볶이" in name and item.mapping.status == MenuMappingStatus.AMBIGUOUS and "밀키트" not in name:
            buckets["UNSPECIFIED_TTEOKBOKKI"] += amount
        elif "밀키트" in name:
            buckets["MILKIT"] += amount
        elif name == "배달료":
            buckets["DELIVERY_FEE"] += amount
        else:
            buckets["OTHER_NEW_MENU"] += amount
    return buckets
