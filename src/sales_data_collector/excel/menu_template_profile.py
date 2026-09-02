from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class MenuTemplateProfile:
    name: str
    sheet_name: str
    store_header: str
    store_column: int = 3
    sales_marker_column: int = 6
    menu_start_column: int = 7
    direct_menu_end_column: int = 27
    other_column: int = 28


DAEJEON_JULY_PROFILE = MenuTemplateProfile(
    name="daejeon_july",
    sheet_name="07월대전통합본",
    store_header="가맹점명",
)

DAEGU_JULY_PROFILE = MenuTemplateProfile(
    name="daegu_july",
    sheet_name="7월 통합본",
    store_header="가맹점",
)

