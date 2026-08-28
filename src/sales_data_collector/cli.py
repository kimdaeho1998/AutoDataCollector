from __future__ import annotations

import argparse
import os
from datetime import date
from getpass import getpass
from pathlib import Path
from typing import Sequence

from openpyxl import load_workbook

from .client import ServiceClient
from .collector import collect_sales, export_sales
from .exceptions import AuthenticationError, ServiceError
from .models import CollectionMode, ExportFormat, SalesAdminDailyRecord, Store
from .production import SalesAdminDryRun, SingleDaySalesCollector, default_target_date
from .mapping.template_resolver import SalesAdminTemplateResolver
from .mapping.store_normalizer import normalize_store_name
from .writers.single_day_writer import SalesAdminSingleDayWriter
from .utils import date_range, parse_ymd


DEFAULT_BASE_URL = os.environ.get("COLLECTOR_BASE_URL")
DEFAULT_BRAND_IDX = os.environ.get("COLLECTOR_BRAND_IDX")
DEFAULT_BRAND_NAME = os.environ.get("COLLECTOR_BRAND_NAME")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="Sales Data Collector")
    parser.add_argument(
        "--base-url",
        default=DEFAULT_BASE_URL,
        help="Service base URL. Set COLLECTOR_BASE_URL locally or provide this option.",
    )
    parser.add_argument("--user-id", help="Service login ID")
    parser.add_argument("--brand-idx", default=DEFAULT_BRAND_IDX)
    parser.add_argument("--brand-name", default=DEFAULT_BRAND_NAME)
    parser.add_argument("--store-idx", action="append", default=[], help="Specific store ID. Repeatable.")
    parser.add_argument("--store-name", action="append", default=[], help="Specific store name. Repeatable.")
    parser.add_argument("--all-stores", action="store_true", help="Collect all returned stores")
    parser.add_argument("--start-date", help="YYYY-MM-DD (legacy diagnostic mode)")
    parser.add_argument("--end-date", help="YYYY-MM-DD (legacy diagnostic mode)")
    parser.add_argument("--output", help="Output xlsx path (legacy diagnostic mode)")
    parser.add_argument("--production-dry-run", action="store_true", help="Preview one production workbook update without writing.")
    parser.add_argument("--production-write", action="store_true", help="Copy a workbook and write confirmed production updates.")
    parser.add_argument("--date", action="append", default=[], help="YYYY-MM-DD. Repeat for multi-date production dry-run; defaults to yesterday.")
    parser.add_argument("--template", help="Sales-admin workbook path for production dry-run mode.")
    parser.add_argument("--production-output", help="Output xlsx path for production write mode.")
    parser.add_argument("--mode", choices=[m.value for m in CollectionMode], default=CollectionMode.ADMIN.value)
    return parser


def login_and_get_stores(client: ServiceClient, args: argparse.Namespace) -> list[Store]:
    """Require a verified authenticated session before any collection begins."""
    while True:
        user_id = args.user_id or input("Service ID: ").strip()
        password = getpass("Password: ")
        try:
            print("[INFO] Logging in...")
            client.login(user_id, password)
            stores = client.get_stores(args.brand_idx)
            if not stores:
                raise AuthenticationError("login verification failed")
            print("[OK] Login succeeded")
            return stores
        except AuthenticationError:
            print("[WARN] Service ID or password is incorrect. Please try again.")
            cookies = getattr(client.session, "cookies", None)
            if cookies is not None:
                cookies.clear()


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        if not args.base_url or not args.brand_idx or not args.brand_name:
            parser.error(
                "--base-url, --brand-idx, and --brand-name are required. "
                "Set COLLECTOR_BASE_URL, COLLECTOR_BRAND_IDX, and COLLECTOR_BRAND_NAME locally."
            )
        if args.production_dry_run and args.production_write:
            parser.error("--production-dry-run and --production-write cannot be used together")
        if args.production_dry_run or args.production_write:
            return run_production_mode(args, write=args.production_write)
        if not args.start_date or not args.end_date or not args.output:
            parser.error("--start-date, --end-date, and --output are required outside --production-dry-run")
        start_date = parse_ymd(args.start_date)
        end_date = parse_ymd(args.end_date)
        business_dates = date_range(start_date, end_date)

        client = ServiceClient(args.base_url)
        stores = login_and_get_stores(client, args)
        print(f"[OK] Retrieved {len(stores)} stores")

        selected_stores = resolve_stores(
            stores,
            all_stores=args.all_stores or (not args.store_idx and not args.store_name),
            store_ids=args.store_idx,
            store_names=args.store_name,
        )
        print(f"[INFO] Collecting {len(selected_stores)} stores for {len(business_dates)} day(s)")

        results = collect_sales(
            client,
            business_dates=business_dates,
            brand_idx=args.brand_idx,
            brand_name=args.brand_name,
            stores=selected_stores,
            mode=CollectionMode(args.mode),
        )
        output_path = export_sales(results, args.output, mode=CollectionMode(args.mode))
        print(f"[OK] Exported to {output_path}")
        client.logout()
        return 0
    except ServiceError as exc:
        print(f"[ERROR] {exc}")
        return 1
    except Exception as exc:
        print(f"[ERROR] {exc}")
        return 1


def run_production_mode(args: argparse.Namespace, *, write: bool) -> int:
    if not args.template:
        raise ValueError("--template is required for production mode")
    if args.all_stores or len(args.date) > 1 or (args.store_name and not args.store_idx):
        if args.all_stores and (args.store_idx or args.store_name):
            raise ValueError("use either --all-stores or specific store selection")
        if args.store_idx and not args.store_name:
            raise ValueError("--store-idx requires a matching --store-name")
        if not args.all_stores and not args.store_name:
            raise ValueError("select --all-stores or provide --store-name")
        return run_store_batch_production(args, write=write)
    if len(args.store_idx) != 1 or len(args.store_name) != 1:
        raise ValueError("single-store production requires one --store-idx and one --store-name")
    business_date = parse_ymd(args.date[0]) if args.date else default_target_date()
    client = ServiceClient(args.base_url)
    try:
        print(f"[INFO] Production {'write' if write else 'dry-run'} target: {business_date.isoformat()}")
        login_and_get_stores(client, args)
        record = SingleDaySalesCollector(
            client, brand_idx=args.brand_idx, brand_name=args.brand_name
        ).collect(
            store=Store(args.store_idx[0], args.store_name[0]),
            business_date=business_date,
        )
        preview = SalesAdminDryRun(args.template).preview(record)
        print(f"[PREVIEW] status={preview.status.value}")
        print(f"[PREVIEW] store={record.store_name} date={record.business_date.isoformat()}")
        print(f"[PREVIEW] receipt_count={record.receipt_count} gross_sales_amount={record.gross_sales_amount}")
        for change in preview.changes:
            print(f"[PREVIEW] {change.sheet}!{change.cell} {change.metric}: {change.old_value!r} -> {change.new_value}")
        if preview.reason:
            print(f"[PREVIEW] reason={preview.reason}")
        if not write:
            return 0 if preview.status.value in {"READY", "SAME_VALUE"} else 1
        if preview.status.value != "READY":
            return 1
        if input("Proceed? [y/N] ").strip().lower() != "y":
            print("[INFO] Write cancelled")
            return 0
        output = Path(args.production_output) if args.production_output else Path("output") / f"sales_collection_{business_date:%Y%m%d}.xlsx"
        result = SalesAdminSingleDayWriter(args.template).write(preview, output)
        print(f"[WRITE] output={result.output_path}")
        print(f"[WRITE] cells_written={result.changes_written} original_unchanged={result.original_hash_before == result.original_hash_after}")
        return 0
    finally:
        client.logout()


def run_store_batch_production(args: argparse.Namespace, *, write: bool) -> int:
    business_dates = [parse_ymd(value) for value in args.date] or [default_target_date()]
    template = Path(args.template)
    workbook = load_workbook(template, data_only=False)
    resolver = SalesAdminTemplateResolver(workbook)
    client = ServiceClient(args.base_url)
    failures = 0
    skipped_inactive = 0
    skipped_no_erp_data = 0
    ready_previews = []
    failure_details: list[str] = []
    try:
        print(f"[INFO] All-store production {'write' if write else 'dry-run'} dates: {', '.join(value.isoformat() for value in business_dates)}")
        available_stores = login_and_get_stores(client, args)
        stores = resolve_stores(
            available_stores,
            all_stores=args.all_stores,
            store_ids=args.store_idx,
            store_names=args.store_name,
        )
        if not stores:
            raise ValueError("STORE_NOT_FOUND")
        if args.store_name and not args.store_idx and len(stores) > 1:
            raise ValueError("AMBIGUOUS_STORE_NAME")
        print(f"[INFO] Selected {len(stores)} of {len(available_stores)} stores")
        collector = SingleDaySalesCollector(client, brand_idx=args.brand_idx, brand_name=args.brand_name)
        previewer = SalesAdminDryRun(template)
        for business_date in business_dates:
            for store in stores:
                try:
                    category = resolver.store_category(business_date, store.magic_store_id, store.store_name)
                    if resolver.is_inactive_category(category):
                        skipped_inactive += 1
                        # Inactive stores have no ERP lookup, but their template cells
                        # must explicitly show that the store was not operating.
                        record = SalesAdminDailyRecord(
                            business_date=business_date,
                            store_id=store.magic_store_id,
                            store_name=store.store_name,
                            receipt_count="-",
                            gross_sales_amount="-",
                        )
                    else:
                        record = collector.collect(store=store, business_date=business_date)
                    preview = previewer.preview(record)
                    cells = ", ".join(change.cell for change in preview.changes)
                    source = "INACTIVE" if resolver.is_inactive_category(category) else "ERP"
                    print(f"[PREVIEW] status={preview.status.value} source={source} date={business_date.isoformat()} store={store.store_name} receipt={record.receipt_count} gross={record.gross_sales_amount} cells={cells}")
                    if preview.status.value == "READY":
                        ready_previews.append(preview)
                    elif preview.status.value != "SAME_VALUE":
                        failures += 1
                        detail = f"date={business_date.isoformat()} store={store.store_name} status={preview.status.value} reason={preview.reason}"
                        failure_details.append(detail)
                        print(f"[FAIL] {detail}")
                except Exception as exc:
                    if str(exc) == "UNMATCHED":
                        skipped_no_erp_data += 1
                        print(f"[SKIP] status=SKIPPED_NO_ERP_DATA date={business_date.isoformat()} store={store.store_name}")
                        continue
                    failures += 1
                    detail = f"date={business_date.isoformat()} store={store.store_name} status=FAILED reason={exc}"
                    failure_details.append(detail)
                    print(f"[PREVIEW] {detail}")
        print(
            f"[SUMMARY] dates={len(business_dates)} stores={len(stores)} "
            f"skipped_inactive={skipped_inactive} "
            f"skipped_no_erp_data={skipped_no_erp_data} "
            f"ready_updates={len(ready_previews)} failures={failures}"
        )
        if failure_details:
            print("[FAILURE SUMMARY]")
            for detail in failure_details:
                print(f"[FAIL] {detail}")
            report_output = Path(args.production_output) if args.production_output else Path("output") / "sales_collection_failure.log"
            report_path = report_output.with_suffix(".failure.log")
            report_path.parent.mkdir(parents=True, exist_ok=True)
            report_path.write_text("\n".join(failure_details) + "\n", encoding="utf-8")
            print(f"[FAILURE REPORT] {report_path}")
        if failures:
            return 1
        if not write:
            return 0
        if not ready_previews:
            print("[INFO] No new values require writing")
            return 0
        if input(f"Proceed with {len(ready_previews)} updates in one copied workbook? [y/N] ").strip().lower() != "y":
            print("[INFO] Write cancelled")
            return 0
        date_label = "_".join(value.strftime("%Y%m%d") for value in business_dates)
        output = Path(args.production_output) if args.production_output else Path("output") / f"sales_collection_{date_label}.xlsx"
        result = SalesAdminSingleDayWriter(template).write_many(ready_previews, output)
        print(f"[WRITE] output={result.output_path}")
        print(f"[WRITE] cells_written={result.changes_written} original_unchanged={result.original_hash_before == result.original_hash_after}")
        return 0
    finally:
        client.logout()


def resolve_stores(
    stores: Sequence[Store],
    *,
    all_stores: bool,
    store_ids: Sequence[str],
    store_names: Sequence[str],
) -> list[Store]:
    if all_stores:
        return list(stores)
    selected: list[Store] = []
    id_set = {value.strip() for value in store_ids if value.strip()}
    name_set = {normalize_store_name(value) for value in store_names if value.strip()}
    for store in stores:
        if id_set and store.magic_store_id in id_set:
            selected.append(store)
        elif name_set and normalize_store_name(store.store_name) in name_set:
            selected.append(store)
    return selected


if __name__ == "__main__":
    raise SystemExit(main())
