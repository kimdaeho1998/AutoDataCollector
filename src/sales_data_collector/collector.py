from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Iterable, Sequence

from .client import ServiceClient
from .exporters import build_exporter
from .models import CollectionMode, ExportFormat, SalesResult, Store
from .utils import date_range


@dataclass
class CollectorOptions:
    mode: CollectionMode
    export_format: ExportFormat
    include_all_stores: bool = True


def collect_sales(
    client: ServiceClient,
    *,
    business_dates: Sequence[date],
    brand_idx: str,
    brand_name: str,
    stores: Sequence[Store],
    mode: CollectionMode,
) -> list[SalesResult]:
    results: list[SalesResult] = []
    for business_date in business_dates:
        for store in stores:
            results.append(
                client.get_time_sales(
                    business_date=business_date,
                    brand_idx=brand_idx,
                    brand_name=brand_name,
                    store_idx=store.magic_store_id,
                    store_name=store.store_name,
                )
            )
    return results


def export_sales(
    results: Sequence[SalesResult],
    output_path: str | Path,
    *,
    mode: CollectionMode,
) -> Path:
    exporter = build_exporter(mode)
    return exporter.export(output_path, results)

