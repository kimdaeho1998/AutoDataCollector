from __future__ import annotations

import argparse
import calendar
import shutil
import sys
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

from sales_data_collector.excel.menu_excel_dry_run import (
    CellPlanStatus,
    build_menu_excel_dry_run_plan,
)

from sales_data_collector.excel.menu_quantity_row import (
    insert_quantity_row,
)

from sales_data_collector.excel.menu_template_resolver import (
    ExcelDisposition,
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

    if plan.ac_plan.status not in {
        CellPlanStatus.READY,
        CellPlanStatus.SAME_VALUE,
    }:
        raise ValueError(
            f"AC_NOT_WRITABLE:"
            f"{plan.store_name}:"
            f"{plan.ac_plan.status.value}"
        )

    allowed = {
        CellPlanStatus.READY,
        CellPlanStatus.SAME_VALUE,
        CellPlanStatus.ZERO_PLACEHOLDER,
    }

    blocked = [
        cell
        for cell in plan.cells
        if cell.status not in allowed
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

    print("")
    print("=" * 120)
    print(
        f"STORE START | {store.store_name}"
    )
    print("=" * 120)

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

    for cell in plan.cells:

        if cell.status in {
            CellPlanStatus.READY,
            CellPlanStatus.ZERO_PLACEHOLDER,
        }:

            worksheet[
                cell.target_cell
            ].value = cell.proposed_value

            written += 1

        elif (
            cell.status
            == CellPlanStatus.SAME_VALUE
        ):

            same += 1

    # AC = monthly source total sales
    if (
        plan.ac_plan.status
        == CellPlanStatus.READY
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
        f"STORE={store.store_name}"
    )

    print(
        f"EXCEL_STORE={plan.excel_store_name}"
    )

    print(
        f"SALES_ROW={quantity_result.sales_row}"
    )

    print(
        f"QUANTITY_ROW={quantity_result.quantity_row}"
    )

    print(
        f"RATIO_ROW={quantity_result.ratio_row}"
    )

    print(
        f"MONTHLY_TOTAL_SALES="
        f"{plan.source_total_sales:,}"
    )

    print(
        f"MONTHLY_TOTAL_QUANTITY="
        f"{quantity_result.source_total_quantity:,}"
    )

    print(
        f"DIRECT_SALES="
        f"{plan.direct_target_sales:,}"
    )

    print(
        f"OTHER_SALES="
        f"{plan.source_other_residual:,}"
    )

    print(
        f"DIRECT_QUANTITY="
        f"{quantity_result.direct_quantity:,}"
    )

    print(
        f"OTHER_QUANTITY="
        f"{quantity_result.other_quantity:,}"
    )

    print(
        f"OPTION_QUANTITY="
        f"{quantity_result.option_quantity:,}"
    )

    print(
        f"AB_QUANTITY="
        f"{quantity_result.ab_quantity:,}"
    )

    print(
        f"WRITTEN_CELLS={written}"
    )

    print(
        f"SAME_VALUE_CELLS={same}"
    )

    print(
        f"ANALYSIS_ROWS={analysis_count}"
    )

    print(
        "SALES_RECONCILIATION=PASS"
    )

    print(
        "QUANTITY_RECONCILIATION=PASS"
    )

    print(
        "STORE_RESULT=PASS"
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


# ====================================================================================================
# STORE RESOLUTION
# ====================================================================================================

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
            raise ValueError(
                f"MAGICERP_STORE_NOT_FOUND:"
                f"{requested_name}"
            )

        if len(matches) != 1:
            raise ValueError(
                f"MAGICERP_STORE_AMBIGUOUS:"
                f"{requested_name}"
            )

        resolved.append(
            matches[0]
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
        raise ValueError(
            f"OUTPUT_ALREADY_EXISTS:{output_path}"
        )

    if args.month < 1 or args.month > 12:
        raise ValueError(
            "MONTH_OUT_OF_RANGE"
        )

    if len(args.store_name) < 2:
        raise ValueError(
            "AT_LEAST_TWO_STORE_NAMES_REQUIRED"
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

        target_stores = (
            resolve_requested_stores(
                available_stores,
                args.store_name,
            )
        )

        print(
            f"[INFO] TARGET_STORE_COUNT="
            f"{len(target_stores)}"
        )

        results = []

        # --------------------------------------------------------------------------------
        # STORE LOOP
        #
        # Critical:
        #
        # Plan is built HERE for each store against the CURRENT worksheet.
        # Previous quantity insertions are already included.
        # --------------------------------------------------------------------------------

        for store in target_stores:

            result = apply_store(
                client=client,
                workbook=workbook,
                worksheet=worksheet,
                profile=profile,
                store=store,
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

        # --------------------------------------------------------------------------------
        # Finish analysis sheet
        # --------------------------------------------------------------------------------

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

        print("")
        print("=" * 120)
        print("MONTHLY MENU MULTI-STORE COMPLETE")
        print("=" * 120)

        print(
            f"OUTPUT={output_path}"
        )

        print(
            f"STORE_COUNT={len(results)}"
        )

        print(
            f"ANALYSIS_ROW_TOTAL="
            f"{sum(r['analysis_rows'] for r in results)}"
        )

        for index, result in enumerate(
            results,
            start=1,
        ):

            print(
                f"STORE_{index}="
                f"{result['store']} | "
                f"SALES_ROW={result['sales_row']} | "
                f"QUANTITY_ROW={result['quantity_row']} | "
                f"RATIO_ROW={result['ratio_row']} | "
                f"SALES={result['sales']:,} | "
                f"QUANTITY={result['quantity']:,}"
            )

        print("")
        print(
            "BATCH_RESULT=PASS"
        )

        print("=" * 120)

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

