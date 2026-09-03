from __future__ import annotations

import argparse
import calendar
import shutil
import sys
from dataclasses import dataclass
from datetime import date
from pathlib import Path


# ----------------------------------------------------------------------------------------------------
# Project src
# ----------------------------------------------------------------------------------------------------

ROOT = Path(__file__).resolve().parent
SRC = ROOT / "src"

if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))


from openpyxl import load_workbook
from openpyxl.styles import Font

from sales_data_collector.client import ServiceClient

from sales_data_collector.cli import (
    infer_menu_template_profile,
    login_and_get_stores,
    resolve_stores,
)
from sales_data_collector.models import Store

from sales_data_collector.excel.menu_excel_dry_run import (
    CellPlanStatus,
    build_menu_excel_dry_run_plan,
)

from sales_data_collector.excel.menu_quantity_row import (
    DIRECT_END_COLUMN,
    DIRECT_START_COLUMN,
    OTHER_COLUMN,
    QUANTITY_LABEL,
    RATIO_LABEL,
    TOTAL_COLUMN,
    UNTOUCHED_COLUMN,
    _capture_affected_merges,
    _capture_moved_formulas,
    _copy_row_style,
    _label,
    _remove_affected_merges,
    _restore_affected_merges,
    _translate_moved_formulas,
    insert_quantity_row,
)

from sales_data_collector.excel.menu_template_resolver import (
    ExcelDisposition,
    MenuTargetStatus,
    resolve_store_sales_row,
)

from sales_data_collector.mapping.menu_mapping import (
    MenuMappingStatus,
    MenuRowType,
    build_menu_mapping_preview,
)

from sales_data_collector.excel.menu_copy_writer import (
    MenuMonthlyCopyWriter,
    _display_name,
    _single_unit_price,
)


# ====================================================================================================
# ARGUMENTS
# ====================================================================================================

def build_parser() -> argparse.ArgumentParser:

    parser = argparse.ArgumentParser(
        prog="Monthly Menu Multi Store Collector"
    )

    parser.add_argument(
        "--template",
        required=True,
    )

    parser.add_argument(
        "--output",
        required=True,
    )

    parser.add_argument(
        "--year",
        type=int,
        required=True,
    )

    parser.add_argument(
        "--month",
        type=int,
        required=True,
    )

    parser.add_argument(
        "--store-name",
        action="append",
        required=True,
    )

    return parser


# ====================================================================================================
# WRITABLE GATE
# ====================================================================================================

def ensure_plan_writable(plan) -> None:

    if not plan.store_row:
        raise ValueError(
            f"STORE_ROW_NOT_RESOLVED:{plan.store_name}"
        )

    if not plan.residual_match:
        raise ValueError(
            f"SALES_RECONCILIATION_FAILED:{plan.store_name}"
        )

    if not plan.ab_validation.formula_valid:
        raise ValueError(
            f"AB_FORMULA_INVALID:{plan.store_name}"
        )

    if not _is_writable_value_cell(
        plan.ac_plan
    ):
        raise ValueError(
            f"AC_NOT_WRITABLE:"
            f"{plan.store_name}:"
            f"{plan.ac_plan.status.value}"
        )

    blocked = [
        cell
        for cell in plan.cells
        if not _is_writable_value_cell(
            cell
        )
    ]

    if blocked:

        detail = ",".join(
            f"{cell.target_cell}:{cell.status.value}"
            for cell in blocked
        )

        raise ValueError(
            f"MENU_CELL_NOT_WRITABLE:"
            f"{plan.store_name}:"
            f"{detail}"
        )

    source_total_quantity = (
        plan.source.source.source_total_quantity
    )

    if source_total_quantity is not None:

        calculated = (
            int(plan.direct_target_quantity)
            + int(plan.source_other_residual_quantity)
            + int(plan.option_quantity)
        )

        if calculated != int(source_total_quantity):

            raise ValueError(
                f"QUANTITY_RECONCILIATION_FAILED:"
                f"{plan.store_name}:"
                f"SOURCE={source_total_quantity}:"
                f"CALCULATED={calculated}"
            )


def repair_ab_residual_formula(
    worksheet,
    plan,
) -> bool:
    if plan.ab_validation.formula_valid:
        return False

    if not plan.store_row:
        return False

    target_cell = f"AB{plan.store_row}"
    expected = (
        f"=AC{plan.store_row}"
        f"-SUM(G{plan.store_row}:AA{plan.store_row})"
    )

    worksheet[
        target_cell
    ].value = expected

    return True


def _is_writable_value_cell(
    cell,
) -> bool:
    """W6-style policy: ERP values may replace existing static values."""

    if cell.status in {
        CellPlanStatus.READY,
        CellPlanStatus.SAME_VALUE,
        CellPlanStatus.ZERO_PLACEHOLDER,
        CellPlanStatus.CONFLICT,
        CellPlanStatus.FORMULA_PROTECTED,
    }:
        return True

    return False


# ====================================================================================================
# ANALYSIS SHEET
# ====================================================================================================

ANALYSIS_HEADERS = [
    "연도",
    "월",
    "지역",
    "가맹점명",
    "Canonical Code",
    "메뉴명",
    "판매건수",
    "판매금액",
    "기준단가",
    "건당 평균매출",
    "매출비율",
    "판매건수비율",
    "Excel 처리구분",
    "Excel 대상컬럼",
    "Source 상태",
]


def create_analysis_sheet(
    workbook,
    month: int,
):

    name = f"{month:02d}월 메뉴분석"

    if name in workbook.sheetnames:
        del workbook[name]

    sheet = workbook.create_sheet(name)

    sheet.append(
        ANALYSIS_HEADERS
    )

    for cell in sheet[1]:
        cell.font = Font(
            bold=True
        )

    sheet.freeze_panes = "A2"

    widths = [
        10,
        8,
        16,
        24,
        24,
        32,
        12,
        14,
        12,
        16,
        12,
        14,
        20,
        16,
        18,
    ]

    for index, width in enumerate(
        widths,
        start=1,
    ):

        letter = sheet.cell(
            row=1,
            column=index,
        ).column_letter

        sheet.column_dimensions[
            letter
        ].width = width

    return sheet


def append_analysis_rows(
    sheet,
    plan,
    profile,
    *,
    year: int,
    month: int,
) -> int:

    count = 0

    direct_by_code = {
        cell.canonical_code: cell
        for cell in plan.cells
    }

    # --------------------------------------------------------------------------------
    # Direct target menus
    # --------------------------------------------------------------------------------

    for aggregate in plan.source.aggregates:

        direct_cell = direct_by_code.get(
            aggregate.canonical_code
        )

        if direct_cell is None:
            continue

        unit_price = _single_unit_price(
            plan,
            aggregate.canonical_code,
        )

        MenuMonthlyCopyWriter._append_analysis_row(
            sheet,
            year,
            month,
            profile.name,
            plan.store_name,
            aggregate.canonical_code,
            _display_name(
                aggregate.canonical_code,
                direct_cell.excel_header,
            ),
            aggregate.quantity,
            aggregate.sales_amount,
            unit_price,
            ExcelDisposition.DIRECT_TARGET.value,
            direct_cell.target_column,
            MenuMappingStatus.MAPPED.value,
            plan.source_total_sales,
            plan.business_menu_count,
        )

        count += 1

    # --------------------------------------------------------------------------------
    # Other / option / residual
    # --------------------------------------------------------------------------------

    for item in plan.disposition.items:

        if (
            item.disposition
            == ExcelDisposition.DIRECT_TARGET
        ):
            continue

        record = item.mapping.record

        MenuMonthlyCopyWriter._append_analysis_row(
            sheet,
            year,
            month,
            profile.name,
            plan.store_name,
            item.mapping.canonical_code,
            record.menu_name,
            record.sales_quantity or 0,
            record.sales_amount,
            record.unit_price,
            item.disposition.value,
            None,
            item.mapping.status.value,
            plan.source_total_sales,
            (
                plan.business_menu_count
                if item.mapping.row_type
                != MenuRowType.OPTION
                else 0
            ),
        )

        count += 1

    return count


# ====================================================================================================
# APPLY ONE STORE
# ====================================================================================================

def apply_store(
    *,
    client,
    workbook,
    worksheet,
    profile,
    store,
    brand_idx: str,
    brand_name: str,
    period_start: date,
    period_end: date,
    year: int,
    month: int,
    analysis_sheet,
):

    # --------------------------------------------------------------------------------
    # MagicERP monthly menu source
    # --------------------------------------------------------------------------------

    source = client.get_menu_monthly_sales(
        start_date=period_start,
        end_date=period_end,
        brand_idx=brand_idx,
        brand_name=brand_name,
        store_idx=store.magic_store_id,
        store_name=store.store_name,
    )

    if is_empty_menu_source(
        source
    ):
        raise ValueError(
            "MENU_SOURCE_NO_DATA"
        )

    # --------------------------------------------------------------------------------
    # Mapping
    # --------------------------------------------------------------------------------

    mapping_preview = (
        build_menu_mapping_preview(
            source
        )
    )

    # --------------------------------------------------------------------------------
    # CRITICAL
    #
    # Build plan against CURRENT worksheet.
    #
    # Any rows inserted by previous stores are already reflected here.
    # We do not reuse old row coordinates.
    # --------------------------------------------------------------------------------

    plan = build_menu_excel_dry_run_plan(
        mapping_preview,
        worksheet,
        profile,
        store.store_name,
    )

    if repair_ab_residual_formula(
        worksheet,
        plan,
    ):
        plan = build_menu_excel_dry_run_plan(
            mapping_preview,
            worksheet,
            profile,
            store.store_name,
        )

    ensure_plan_writable(
        plan
    )

    sales_row = int(
        plan.store_row
    )

    # --------------------------------------------------------------------------------
    # Sales
    # --------------------------------------------------------------------------------

    written = 0
    same = 0

    for column in range(
        DIRECT_START_COLUMN,
        DIRECT_END_COLUMN + 1,
    ):
        worksheet.cell(
            row=sales_row,
            column=column,
        ).value = 0

    # G:AA were intentionally reset to zero above.
    #
    # Therefore every ERP-backed direct menu cell MUST be written again,
    # including SAME_VALUE cells.
    #
    # SAME_VALUE only means:
    #     original workbook value == ERP proposed value
    #
    # It does NOT mean the write can be skipped after the row has already
    # been zero-initialized.
    for cell in plan.cells:

        if _is_writable_value_cell(cell):

            worksheet[
                cell.target_cell
            ].value = cell.proposed_value

            if cell.status == CellPlanStatus.SAME_VALUE:
                same += 1
            else:
                written += 1

    # AC = monthly source total sales
    if (
        _is_writable_value_cell(
            plan.ac_plan
        )
        and plan.ac_plan.status
        != CellPlanStatus.SAME_VALUE
    ):

        worksheet[
            plan.ac_plan.target_cell
        ].value = (
            plan.ac_plan.proposed_value
        )

        written += 1

    elif (
        plan.ac_plan.status
        == CellPlanStatus.SAME_VALUE
    ):

        same += 1

    # --------------------------------------------------------------------------------
    # Analysis rows
    # --------------------------------------------------------------------------------

    analysis_count = (
        append_analysis_rows(
            analysis_sheet,
            plan,
            profile,
            year=year,
            month=month,
        )
    )

    # --------------------------------------------------------------------------------
    # Quantity physical row
    #
    # 매출
    # 건수
    # 비율
    # --------------------------------------------------------------------------------

    quantity_result = insert_quantity_row(
        worksheet,
        plan,
    )

    # --------------------------------------------------------------------------------
    # Summary
    # --------------------------------------------------------------------------------

    print(
        "[PREVIEW] status=READY source=ERP "
        f"period={period_start:%Y-%m} "
        f"store={store.store_name} "
        f"quantity={quantity_result.source_total_quantity} "
        f"gross={plan.source_total_sales} "
        f"cells={quantity_result.sales_row},{quantity_result.quantity_row},{quantity_result.ratio_row}"
    )

    return {
        "store": store.store_name,
        "sales_row": quantity_result.sales_row,
        "quantity_row": quantity_result.quantity_row,
        "ratio_row": quantity_result.ratio_row,
        "sales": plan.source_total_sales,
        "quantity": quantity_result.source_total_quantity,
        "analysis_rows": analysis_count,
    }


def apply_skipped_store(
    *,
    worksheet,
    profile,
    store_name: str,
    reason: str,
):
    """
    W6-compatible skip policy for inactive or no-data stores.

    The source workbook is already copied. For skipped menu stores we still
    make the visible three-row block explicit:

        sales / quantity / ratio

    and write '-' to G:AC so users do not see misleading #DIV/0! formulas.
    """

    store_row = resolve_store_sales_row(
        worksheet,
        profile,
        store_name,
    )

    if store_row.status != MenuTargetStatus.TARGET_RESOLVED:
        raise ValueError(
            f"SKIP_STORE_ROW_NOT_RESOLVED:{store_name}:{store_row.status.value}"
        )

    sales_row = int(store_row.row_index)
    next_row = sales_row + 1
    next_label = _label(
        worksheet.cell(
            row=next_row,
            column=4,
        ).value
    )

    if next_label == QUANTITY_LABEL:
        quantity_row = next_row
        ratio_row = sales_row + 2
    else:
        if next_label != RATIO_LABEL:
            raise ValueError(
                f"SKIP_RATIO_ROW_LABEL_MISMATCH:{store_name}:F{next_row}:{next_label!r}"
            )

        quantity_row = sales_row + 1
        ratio_row = sales_row + 2

        formula_snapshots = _capture_moved_formulas(
            worksheet,
            quantity_row,
        )
        merge_snapshots = _capture_affected_merges(
            worksheet,
            quantity_row,
        )

        _remove_affected_merges(
            worksheet,
            merge_snapshots,
        )

        worksheet.insert_rows(
            quantity_row,
            amount=1,
        )

        _copy_row_style(
            worksheet,
            source_row=sales_row,
            destination_row=quantity_row,
        )

        _translate_moved_formulas(
            worksheet,
            formula_snapshots,
            quantity_row,
        )

        for column in range(
            1,
            UNTOUCHED_COLUMN + 1,
        ):
            worksheet.cell(
                row=quantity_row,
                column=column,
            ).value = None

        worksheet.cell(
            row=quantity_row,
            column=4,
        ).value = QUANTITY_LABEL

        _restore_affected_merges(
            worksheet,
            merge_snapshots,
        )

    for row in (
        sales_row,
        quantity_row,
        ratio_row,
    ):
        for column in range(
            DIRECT_START_COLUMN,
            TOTAL_COLUMN + 1,
        ):
            worksheet.cell(
                row=row,
                column=column,
            ).value = "-"

    worksheet.cell(
        row=quantity_row,
        column=4,
    ).value = QUANTITY_LABEL

    reason_code = _skip_reason_code(reason)

    print(
        f"[SKIP] status=SKIPPED_PLACEHOLDER store={store_name} "
        f"reason={reason_code} sales_row={sales_row} "
        f"quantity_row={quantity_row} ratio_row={ratio_row}"
    )

    return {
        "store": store_name,
        "sales_row": sales_row,
        "quantity_row": quantity_row,
        "ratio_row": ratio_row,
        "sales": "-",
        "quantity": "-",
        "analysis_rows": 0,
        "skipped": True,
        "skip_reason": reason_code,
    }


def menu_store_category(
    worksheet,
    profile,
    store_name: str,
) -> str | None:
    store_row = resolve_store_sales_row(
        worksheet,
        profile,
        store_name,
    )

    if store_row.status != MenuTargetStatus.TARGET_RESOLVED:
        return None

    value = worksheet.cell(
        row=store_row.row_index,
        column=1,
    ).value

    return str(value).strip() if value is not None else None


def is_inactive_menu_category(
    category: str | None,
) -> bool:
    if not category:
        return False

    compact = str(category).replace(" ", "")
    return "\uc911\ub2e8" in compact or "\ud3d0\uc810" in compact


def is_empty_menu_source(
    source,
) -> bool:
    records = getattr(source, "records", None) or []
    total_sales = getattr(source, "source_total_sales", None)
    total_quantity = getattr(source, "source_total_quantity", None)

    return (
        not records
        and total_sales in (None, 0)
        and total_quantity in (None, 0)
    )


def _summary_value(
    value,
) -> str:
    if isinstance(
        value,
        int,
    ):
        return f"{value:,}"

    return str(
        value
    )


def _skip_reason_code(
    reason: str,
) -> str:
    if not reason:
        return "SKIPPED"

    if "INACTIVE_CATEGORY" in reason:
        return "INACTIVE"

    if "MAGICERP_STORE_NOT_FOUND" in reason:
        return "NO_ERP_STORE"

    if (
        "contains no parseable menu rows" in reason
        or "MENU_SOURCE_NO_DATA" in reason
        or "NO_DATA" in reason
    ):
        return "NO_MENU_DATA"

    if "AC_NOT_WRITABLE" in reason:
        return "TOTAL_CELL_NOT_WRITABLE"

    if "MENU_CELL_NOT_WRITABLE" in reason:
        return "MENU_CELL_NOT_WRITABLE"

    if "FORMULA_PROTECTED" in reason:
        return "FORMULA_PROTECTED"

    return reason.split(";", 1)[0].split(":", 1)[0]


def _is_placeholder_skip_reason(
    reason: str,
) -> bool:
    code = _skip_reason_code(
        reason
    )

    return code in {
        "INACTIVE",
        "NO_ERP_STORE",
        "NO_MENU_DATA",
    }


# ====================================================================================================
# STORE RESOLUTION
# ====================================================================================================

@dataclass(frozen=True)
class RequestedStoreTarget:
    requested_name: str
    store: Store | None
    skip_reason: str | None = None


def resolve_requested_stores(
    available_stores,
    requested_names,
):

    resolved = []

    for requested_name in requested_names:

        matches = resolve_stores(
            available_stores,
            all_stores=False,
            store_ids=(),
            store_names=[
                requested_name
            ],
        )

        if not matches:
            resolved.append(
                RequestedStoreTarget(
                    requested_name=requested_name,
                    store=None,
                    skip_reason="MAGICERP_STORE_NOT_FOUND",
                )
            )
            continue

        if len(matches) != 1:
            raise ValueError(
                f"MAGICERP_STORE_AMBIGUOUS:"
                f"{requested_name}"
            )

        resolved.append(
            RequestedStoreTarget(
                requested_name=requested_name,
                store=matches[0],
            )
        )

    return resolved


# ====================================================================================================
# MAIN
# ====================================================================================================

def main() -> int:

    parser = build_parser()
    args = parser.parse_args()

    source_path = Path(
        args.template
    ).resolve()

    output_path = Path(
        args.output
    ).resolve()

    if not source_path.exists():
        raise ValueError(
            f"SOURCE_NOT_FOUND:{source_path}"
        )

    if output_path == source_path:
        raise ValueError(
            "OUTPUT_MUST_DIFFER_FROM_SOURCE"
        )

    if output_path.exists():
        output_path.unlink()

    if args.month < 1 or args.month > 12:
        raise ValueError(
            "MONTH_OUT_OF_RANGE"
        )

    if len(args.store_name) < 1:
        raise ValueError(
            "AT_LEAST_ONE_STORE_NAME_REQUIRED"
        )

    period_start = date(
        args.year,
        args.month,
        1,
    )

    period_end = date(
        args.year,
        args.month,
        calendar.monthrange(
            args.year,
            args.month,
        )[1],
    )

    # --------------------------------------------------------------------------------
    # Copy source ONCE.
    # --------------------------------------------------------------------------------

    output_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    shutil.copy2(
        source_path,
        output_path,
    )

    client = None
    workbook = None

    try:

        workbook = load_workbook(
            output_path,
            data_only=False,
        )

        profile = infer_menu_template_profile(
            workbook
        )

        worksheet = workbook[
            profile.sheet_name
        ]

        analysis_sheet = (
            create_analysis_sheet(
                workbook,
                args.month,
            )
        )

        # --------------------------------------------------------------------------------
        # Login ONCE.
        # --------------------------------------------------------------------------------

        # Reuse project environment configuration.
        #
        # ServiceClient(None) follows the same environment-backed configuration
        # only when constructor supports it. If base URL is mandatory, CLI object
        # below provides the environment value.
        from sales_data_collector.cli import (
            DEFAULT_BASE_URL,
            DEFAULT_BRAND_IDX,
            DEFAULT_BRAND_NAME,
        )

        if not DEFAULT_BASE_URL:
            raise ValueError(
                "COLLECTOR_BASE_URL_NOT_SET"
            )

        if not DEFAULT_BRAND_IDX:
            raise ValueError(
                "COLLECTOR_BRAND_IDX_NOT_SET"
            )

        if not DEFAULT_BRAND_NAME:
            raise ValueError(
                "COLLECTOR_BRAND_NAME_NOT_SET"
            )

        client = ServiceClient(
            DEFAULT_BASE_URL
        )

        # login_and_get_stores() requires Namespace-compatible fields.
        class LoginArgs:
            user_id = None
            brand_idx = DEFAULT_BRAND_IDX

        print(
            f"[INFO] Monthly menu batch target: "
            f"{period_start.isoformat()} ~ "
            f"{period_end.isoformat()}"
        )

        available_stores = (
            login_and_get_stores(
                client,
                LoginArgs(),
            )
        )

        target_store_targets = (
            resolve_requested_stores(
                available_stores,
                args.store_name,
            )
        )

        print(
            f"[INFO] TARGET_STORE_COUNT="
            f"{len(target_store_targets)}"
        )

        results = []
        skipped_inactive = 0
        skipped_no_erp_data = 0
        skipped_no_menu_data = 0
        placeholder_failures = []
        failures = []

        # --------------------------------------------------------------------------------
        # STORE LOOP
        #
        # Critical:
        #
        # Plan is built HERE for each store against the CURRENT worksheet.
        # Previous quantity insertions are already included.
        # --------------------------------------------------------------------------------

        for target in target_store_targets:

            store_name = (
                target.store.store_name
                if target.store is not None
                else target.requested_name
            )

            try:
                if target.store is None:
                    skipped_no_erp_data += 1
                    result = apply_skipped_store(
                        worksheet=worksheet,
                        profile=profile,
                        store_name=store_name,
                        reason=target.skip_reason or "MAGICERP_STORE_NOT_FOUND",
                    )
                    results.append(
                        result
                    )
                    continue

                category = menu_store_category(
                    worksheet,
                    profile,
                    store_name,
                )

                if is_inactive_menu_category(
                    category
                ):
                    skipped_inactive += 1
                    result = apply_skipped_store(
                        worksheet=worksheet,
                        profile=profile,
                        store_name=store_name,
                        reason=f"INACTIVE_CATEGORY:{category}",
                    )
                    results.append(
                        result
                    )
                    continue

                result = apply_store(
                    client=client,
                    workbook=workbook,
                    worksheet=worksheet,
                    profile=profile,
                    store=target.store,
                    brand_idx=DEFAULT_BRAND_IDX,
                    brand_name=DEFAULT_BRAND_NAME,
                    period_start=period_start,
                    period_end=period_end,
                    year=args.year,
                    month=args.month,
                    analysis_sheet=analysis_sheet,
                )

                results.append(
                    result
                )

            except Exception as exc:
                reason = str(exc)

                if not _is_placeholder_skip_reason(
                    reason
                ):
                    detail = (
                        f"store={store_name} reason={reason}"
                    )
                    failures.append(
                        detail
                    )
                    print(
                        f"[FAIL] status=FAILED store={store_name} "
                        f"reason={_skip_reason_code(reason)}"
                    )
                    continue

                if "MAGICERP_STORE_NOT_FOUND" in reason:
                    skipped_no_erp_data += 1
                else:
                    skipped_no_menu_data += 1

                try:
                    result = apply_skipped_store(
                        worksheet=worksheet,
                        profile=profile,
                        store_name=store_name,
                        reason=reason,
                    )
                    results.append(
                        result
                    )
                except Exception as placeholder_exc:
                    detail = (
                        f"store={store_name} reason={reason} "
                        f"placeholder_error={placeholder_exc}"
                    )
                    placeholder_failures.append(
                        detail
                    )
                    print(
                        f"[FAIL] {detail}"
                    )

        # --------------------------------------------------------------------------------
        # Finish analysis sheet
        # --------------------------------------------------------------------------------

        if placeholder_failures:
            raise ValueError(
                "MENU_PLACEHOLDER_FAILURES:"
                + "; ".join(
                    placeholder_failures
                )
            )

        if failures:
            raise ValueError(
                "MENU_STORE_FAILURES:"
                + "; ".join(
                    failures
                )
            )

        analysis_sheet.auto_filter.ref = (
            analysis_sheet.dimensions
        )

        for row in analysis_sheet.iter_rows(
            min_row=2
        ):

            for cell in row:

                if cell.column == 7:
                    cell.number_format = "#,##0"

                elif cell.column in {
                    8,
                    9,
                    10,
                }:
                    cell.number_format = "#,##0"

                elif cell.column in {
                    11,
                    12,
                }:
                    cell.number_format = "0.00%"

        # --------------------------------------------------------------------------------
        # Save ONCE after all target stores.
        # --------------------------------------------------------------------------------

        workbook.save(
            output_path
        )

        print(
            f"[WRITE] output={output_path}"
        )

        print(
            f"[SUMMARY] period={period_start:%Y-%m} stores={len(results)} "
            f"skipped_inactive={skipped_inactive} "
            f"skipped_no_erp_data={skipped_no_erp_data} "
            f"skipped_no_menu_data={skipped_no_menu_data} "
            f"analysis_rows={sum(r['analysis_rows'] for r in results)}"
        )

        print(
            "BATCH_RESULT=PASS"
        )

        return 0

    except Exception:

        # Atomic output:
        # any store failure removes generated copy.
        if workbook is not None:

            try:
                workbook.close()
            except Exception:
                pass

        output_path.unlink(
            missing_ok=True
        )

        raise

    finally:

        if client is not None:

            try:
                client.logout()
            except Exception:
                pass


if __name__ == "__main__":

    try:
        raise SystemExit(
            main()
        )

    except Exception as exc:

        print(
            f"[ERROR] {exc}"
        )

        raise SystemExit(1)

