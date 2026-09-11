from __future__ import annotations

from dataclasses import dataclass
import re
import unicodedata


# ==================================================================================================
# RESOLUTION MODEL
# ==================================================================================================

@dataclass(frozen=True)
class MenuAliasResolution:
    raw_name: str
    normalized_name: str
    canonical_name: str | None
    canonical_key: str | None
    target_header: str | None
    resolution_type: str
    pack_size: int | None = None


# ==================================================================================================
# TEXT NORMALIZATION
#
# This layer changes only textual representation.
# It must NOT infer business meaning.
# ==================================================================================================

def normalize_menu_alias_text(value: object) -> str:

    if value is None:
        return ""

    text = unicodedata.normalize(
        "NFKC",
        str(value),
    )

    # trim
    text = text.strip()

    # CR/LF/TAB -> space
    text = re.sub(
        r"[\r\n\t]+",
        " ",
        text,
    )

    # repeated spaces
    text = re.sub(
        r"\s+",
        " ",
        text,
    )

    # normalize parentheses
    text = re.sub(
        r"\s*\(\s*",
        "(",
        text,
    )

    text = re.sub(
        r"\s*\)\s*",
        ")",
        text,
    )

    # remove trailing punctuation / spaces
    text = re.sub(
        r"[.\s]+$",
        "",
        text,
    )

    return text.strip()


# ==================================================================================================
# CANONICAL REGISTRY
#
# Direct-menu targets:
#
#   김밥
#       5줄
#       10줄
#       1줄
#       와사비크래마요(4줄)
#       매콤진미꼬마김밥(4줄)
#       유부 꼬마김밥(4줄)
#       불어묵꼬마김밥(4줄)
#
#   어묵탕
#
#   떡볶이
#       떡볶이(순)
#       떡볶이(매)
#
#   쫄면
#       쫄면(순)
#       쫄면(매)
#
#   우동
#       선비우동
#       선비 김치우동
#       우동
#
#   라면
#
#   소스
#       스리라차
#       청양고추
#       마크니커리
#       치즈
#
#   음료
#       식혜
#
# NO fuzzy matching.
# ==================================================================================================

MENU_ALIAS_REGISTRY: dict[str, dict[str, object]] = {


    # ==============================================================================================
    # 01. 꼬마김밥 5줄
    # ==============================================================================================

    "KIMBAP_5": {

        "canonical_name": "꼬마김밥 5줄",
        "target_header": "5줄",
        "pack_size": 5,

        "aliases": (

            "5줄",

            "꼬마김밥5줄",
            "꼬마김밥 5줄",

            "꼬마김밥(5줄)",
            "꼬마김밥 (5줄)",

            "꼬마김밥 소",
            "꼬마김밥소",

            "꼬마김밥(소)",
            "꼬마김밥 (소)",

            "꼬마김밥 소자",
            "꼬마김밥소자",

            "꼬마김밥(소자)",
            "꼬마김밥 (소자)",
        ),
    },


    # ==============================================================================================
    # 02. 꼬마김밥 10줄
    # ==============================================================================================

    "KIMBAP_10": {

        "canonical_name": "꼬마김밥 10줄",
        "target_header": "10줄",
        "pack_size": 10,

        "aliases": (

            "10줄",

            "꼬마김밥10줄",
            "꼬마김밥 10줄",

            "꼬마김밥(10줄)",
            "꼬마김밥 (10줄)",

            "꼬마김밥 대",
            "꼬마김밥대",

            "꼬마김밥(대)",
            "꼬마김밥 (대)",

            "꼬마김밥 대자",
            "꼬마김밥대자",

            "꼬마김밥(대자)",
            "꼬마김밥 (대자)",
        ),
    },


    # ==============================================================================================
    # 03. 꼬마김밥 1줄
    # ==============================================================================================

    "KIMBAP_1": {

        "canonical_name": "꼬마김밥 1줄",
        "target_header": "1줄",
        "pack_size": 1,

        "aliases": (

            "1줄",

            "꼬마김밥1줄",
            "꼬마김밥 1줄",

            "꼬마김밥(1줄)",
            "꼬마김밥 (1줄)",

            "꼬마김밥 한줄",
            "꼬마김밥 한 줄",

            "꼬마김밥 낱줄",
            "꼬마김밥 낱개",
        ),
    },


    # ==============================================================================================
    # 04. 와사비크래마요 4줄
    #
    # IMPORTANT
    #   Explicit 1줄 variant is NOT included.
    # ==============================================================================================

    "WASABI_CRAB_4": {

        "canonical_name": "와사비크래마요(4줄)",
        "target_header": "와사비크래마요(4줄)",
        "pack_size": 4,

        "aliases": (

            "와사비크래마요",
            "와사비 크래마요",

            "와사비크래마요김밥",
            "와사비크래마요 김밥",

            "와사비크래마요꼬마김밥",
            "와사비크래마요 꼬마김밥",

            "와사비 크래마요꼬마김밥",
            "와사비 크래마요 꼬마김밥",

            "와사비크래마요4줄",
            "와사비크래마요 4줄",

            "와사비크래마요(4줄)",
            "와사비크래마요 (4줄)",

            "와사비크래마요꼬마김밥4줄",
            "와사비크래마요꼬마김밥 4줄",

            "와사비크래마요 꼬마김밥4줄",
            "와사비크래마요 꼬마김밥 4줄",

            "와사비크래마요꼬마김밥(4줄)",
            "와사비크래마요 꼬마김밥(4줄)",
        ),
    },


    # ==============================================================================================
    # 05. 매콤진미 꼬마김밥 4줄
    # ==============================================================================================

    "SPICY_JINMI_4": {

        "canonical_name": "매콤진미꼬마김밥(4줄)",
        "target_header": "매콤진미꼬마김밥(4줄)",
        "pack_size": 4,

        "aliases": (

            "매콤진미",

            "매콤진미김밥",
            "매콤진미 김밥",

            "매콤진미꼬마김밥",
            "매콤진미 꼬마김밥",

            "매콤진미4줄",
            "매콤진미 4줄",

            "매콤진미김밥4줄",
            "매콤진미 김밥 4줄",

            "매콤진미꼬마김밥4줄",
            "매콤진미꼬마김밥 4줄",

            "매콤진미 꼬마김밥4줄",
            "매콤진미 꼬마김밥 4줄",

            "매콤진미꼬마김밥(4줄)",
            "매콤진미 꼬마김밥(4줄)",
        ),
    },


    # ==============================================================================================
    # 06. 유부 꼬마김밥 4줄
    # ==============================================================================================

    "YUBU_4": {

        "canonical_name": "유부 꼬마김밥(4줄)",
        "target_header": "유부 꼬마김밥(4줄)",
        "pack_size": 4,

        "aliases": (

            "유부",

            "유부김밥",
            "유부 김밥",

            "유부꼬마김밥",
            "유부 꼬마김밥",

            "유부4줄",
            "유부 4줄",

            "유부김밥4줄",
            "유부 김밥 4줄",

            "유부꼬마김밥4줄",
            "유부꼬마김밥 4줄",

            "유부 꼬마김밥4줄",
            "유부 꼬마김밥 4줄",

            "유부꼬마김밥(4줄)",
            "유부 꼬마김밥(4줄)",
        ),
    },


    # ==============================================================================================
    # 07. 불어묵 꼬마김밥 4줄
    # ==============================================================================================

    "BUL_EOMUK_4": {

        "canonical_name": "불어묵꼬마김밥(4줄)",
        "target_header": "불어묵꼬마김밥(4줄)",
        "pack_size": 4,

        "aliases": (

            "불어묵",

            "불어묵김밥",
            "불어묵 김밥",

            "불어묵꼬마김밥",
            "불어묵 꼬마김밥",

            "불어묵4줄",
            "불어묵 4줄",

            "불어묵김밥4줄",
            "불어묵 김밥 4줄",

            "불어묵꼬마김밥4줄",
            "불어묵꼬마김밥 4줄",

            "불어묵 꼬마김밥4줄",
            "불어묵 꼬마김밥 4줄",

            "불어묵꼬마김밥(4줄)",
            "불어묵 꼬마김밥(4줄)",
        ),
    },


    # ==============================================================================================
    # 08. 어묵탕
    # ==============================================================================================

    "FISH_CAKE_SOUP": {

        "canonical_name": "어묵탕",
        "target_header": "어묵탕",
        "pack_size": None,

        "aliases": (

            "어묵탕",
            "어묵 탕",

            "선비어묵탕",
            "선비 어묵탕",

            "꼬마어묵탕",
            "꼬마 어묵탕",
        ),
    },


    # ==============================================================================================
    # 09. 떡볶이(순)
    #
    # Includes known typo:
    #   떡볶기
    # ==============================================================================================

    "TTEOKBOKKI_MILD": {

        "canonical_name": "떡볶이(순)",
        "target_header": "떡볶이(순)",
        "pack_size": None,

        "aliases": (

            # --------------------------------------------------------------------------------------
            # Correct spelling
            # --------------------------------------------------------------------------------------

            "떡볶이(순)",
            "떡볶이 (순)",
            "떡볶이순",
            "떡볶이 순",

            "순한떡볶이",
            "순한 떡볶이",

            "순한맛떡볶이",
            "순한맛 떡볶이",

            "국물떡볶이(순)",
            "국물떡볶이(순한맛)",
            "국물 떡볶이(순)",

            "국물떡볶이 순",
            "국물 떡볶이 순",

            "순한국물떡볶이",
            "순한 국물떡볶이",

            "순한맛국물떡볶이",
            "순한맛 국물떡볶이",


            # --------------------------------------------------------------------------------------
            # Known typo: 떡볶기
            # --------------------------------------------------------------------------------------

            "떡볶기(순)",
            "떡볶기 (순)",
            "떡볶기순",
            "떡볶기 순",

            "순한떡볶기",
            "순한 떡볶기",

            "순한맛떡볶기",
            "순한맛 떡볶기",

            "국물떡볶기(순)",
            "국물 떡볶기(순)",

            "국물떡볶기 순",
            "국물 떡볶기 순",

            "순한국물떡볶기",
            "순한 국물떡볶기",

            "순한맛국물떡볶기",
            "순한맛 국물떡볶기",
        ),
    },


    # ==============================================================================================
    # 10. 떡볶이(매)
    #
    # Includes known typo:
    #   떡볶기
    # ==============================================================================================

    "TTEOKBOKKI_SPICY": {

        "canonical_name": "떡볶이(매)",
        "target_header": "떡볶이(매)",
        "pack_size": None,

        "aliases": (

            # --------------------------------------------------------------------------------------
            # Correct spelling
            # --------------------------------------------------------------------------------------

            "떡볶이(매)",
            "떡볶이 (매)",
            "떡볶이매",
            "떡볶이 매",

            "매운떡볶이",
            "매운 떡볶이",

            "매운맛떡볶이",
            "매운맛 떡볶이",

            "국물떡볶이(매)",
            "국물떡볶이(매운맛)",
            "국물 떡볶이(매)",

            "국물떡볶이 매",
            "국물 떡볶이 매",

            "매운국물떡볶이",
            "매운 국물떡볶이",

            "매운맛국물떡볶이",
            "매운맛 국물떡볶이",


            # --------------------------------------------------------------------------------------
            # Known typo: 떡볶기
            # --------------------------------------------------------------------------------------

            "떡볶기(매)",
            "떡볶기 (매)",
            "떡볶기매",
            "떡볶기 매",

            "매운떡볶기",
            "매운 떡볶기",

            "매운맛떡볶기",
            "매운맛 떡볶기",

            "국물떡볶기(매)",
            "국물 떡볶기(매)",

            "국물떡볶기 매",
            "국물 떡볶기 매",

            "매운국물떡볶기",
            "매운 국물떡볶기",

            "매운맛국물떡볶기",
            "매운맛 국물떡볶기",
        ),
    },


    # ==============================================================================================
    # 11. 쫄면(순)
    # ==============================================================================================

    "JJOLMYEON_MILD": {

        "canonical_name": "쫄면(순)",
        "target_header": "쫄면(순)",
        "pack_size": None,

        "aliases": (

            "쫄면(순)",
            "쫄면(순한맛)",
            "쫄면 (순)",
            "쫄면순",
            "쫄면 순",

            "순한쫄면",
            "순한 쫄면",

            "순한맛쫄면",
            "순한맛 쫄면",
        ),
    },


    # ==============================================================================================
    # 12. 쫄면(매)
    # ==============================================================================================

    "JJOLMYEON_SPICY": {

        "canonical_name": "쫄면(매)",
        "target_header": "쫄면(매)",
        "pack_size": None,

        "aliases": (

            "쫄면(매)",
            "쫄면(매운맛)",
            "쫄면 (매)",
            "쫄면매",
            "쫄면 매",

            "매운쫄면",
            "매운 쫄면",

            "매운맛쫄면",
            "매운맛 쫄면",
        ),
    },


    # ==============================================================================================
    # 13. 선비우동
    # ==============================================================================================

    "SEONBI_UDON": {

        "canonical_name": "선비우동",
        "target_header": "선비우동",
        "pack_size": None,

        "aliases": (

            "선비우동",
            "선비 우동",

            "선비꼬마우동",
            "선비 꼬마우동",
        ),
    },


    # ==============================================================================================
    # 14. 선비 김치우동
    # ==============================================================================================

    "SEONBI_KIMCHI_UDON": {

        "canonical_name": "선비 김치우동",
        "target_header": "선비 김치우동",
        "pack_size": None,

        "aliases": (

            "선비김치우동",
            "선비 김치우동",

            "선비김치 우동",
            "선비 김치 우동",

            "김치선비우동",
            "김치 선비우동",
        ),
    },


    # ==============================================================================================
    # 15. 우동
    # ==============================================================================================

    "UDON": {

        "canonical_name": "우동",
        "target_header": "우동",
        "pack_size": None,

        "aliases": (

            "우동",

            "일반우동",
            "일반 우동",

            "기본우동",
            "기본 우동",
        ),
    },


    # ==============================================================================================
    # 16. 라면
    # ==============================================================================================

    "RAMEN": {

        "canonical_name": "라면",
        "target_header": "라면",
        "pack_size": None,

        "aliases": (

            "라면",

            "일반라면",
            "일반 라면",

            "기본라면",
            "기본 라면",
        ),
    },


    # ==============================================================================================
    # 17. 스리라차
    # ==============================================================================================

    "SRIRACHA": {

        "canonical_name": "스리라차",
        "target_header": "스리라차",
        "pack_size": None,

        "aliases": (

            "스리라차",
            "스리 라차",

            "스리라차소스",
            "스리라차 소스",

            "스리라챠",
            "스리 라챠",

            "스리라챠소스",
            "스리라챠 소스",
        ),
    },


    # ==============================================================================================
    # 18. 청양고추
    # ==============================================================================================

    "CHEONGYANG": {

        "canonical_name": "청양고추",
        "target_header": "청양고추",
        "pack_size": None,

        "aliases": (

            "청양고추",
            "청양 고추",

            "청양고추소스",
            "청양고추 소스",

            "청양소스",
            "청양 소스",
        ),
    },


    # ==============================================================================================
    # 19. 마크니커리
    # ==============================================================================================

    "MAKHANI_CURRY": {

        "canonical_name": "마크니커리",
        "target_header": "마크니커리",
        "pack_size": None,

        "aliases": (

            "마크니커리",
            "마크니 커리",

            "마크니카레",
            "마크니 카레",

            "마크니커리소스",
            "마크니커리 소스",

            "마크니 커리소스",
            "마크니 커리 소스",

            "마크니카레소스",
            "마크니 카레 소스",
        ),
    },


    # ==============================================================================================
    # 20. 치즈
    # ==============================================================================================

    "CHEESE": {

        "canonical_name": "치즈",
        "target_header": "치즈",
        "pack_size": None,

        "aliases": (

            "치즈",

            "치즈소스",
            "치즈 소스",

            "치즈추가",
            "치즈 추가",
        ),
    },


    # ==============================================================================================
    # 21. 식혜
    #
    # Only dedicated drink target.
    # ==============================================================================================

    "SIKHYE": {

        "canonical_name": "식혜",
        "target_header": "식혜",
        "pack_size": None,

        "aliases": (

            "식혜",

            "선비식혜",
            "선비 식혜",

            "전통식혜",
            "전통 식혜",

            "수제식혜",
            "수제 식혜",
        
            '(HACCP 인증) 선비식혜 150ml',
        ),
    },
}


# ==================================================================================================
# REVIEW-REQUIRED
#
# These MUST NOT be automatically mapped.
#
# Reasons:
#   - missing package/variant information
#   - different selling unit
#   - potentially different product
# ==================================================================================================

REVIEW_REQUIRED_PATTERNS: tuple[re.Pattern[str], ...] = (

    # ----------------------------------------------------------------------------------------------
    # Base 꼬마김밥 without package size.
    # Could be 1 / 5 / 10 line depending on store/POS configuration.
    # ----------------------------------------------------------------------------------------------

    re.compile(
        r"^꼬마김밥$"
    ),


    # ----------------------------------------------------------------------------------------------
    # Explicit 1-line flavored products.
    #
    # Current Excel direct columns represent the 4-line flavored products.
    # Do not convert their product quantity into 4-line product quantity automatically.
    # ----------------------------------------------------------------------------------------------

    re.compile(
        r"^와사비\s*크래마요.*1\s*줄$"
    ),

    re.compile(
        r"^매콤진미.*1\s*줄$"
    ),

    re.compile(
        r"^유부.*1\s*줄$"
    ),

    re.compile(
        r"^불어묵.*1\s*줄$"
    ),


    # ----------------------------------------------------------------------------------------------
    # 떡볶이 / 떡볶기 without mild/spicy information.
    # ----------------------------------------------------------------------------------------------

    re.compile(
        r"^(?:국물\s*)?떡볶[이기]$"
    ),


    # ----------------------------------------------------------------------------------------------
    # 쫄면 without mild/spicy information.
    # ----------------------------------------------------------------------------------------------

    re.compile(
        r"^쫄면$"
    ),
)


# ==================================================================================================
# ALIAS INDEX
# ==================================================================================================

def _build_alias_index() -> dict[str, str]:

    result: dict[str, str] = {}

    for canonical_key, spec in MENU_ALIAS_REGISTRY.items():

        candidates = [
            str(spec["canonical_name"]),
            str(spec["target_header"]),
            *[
                str(alias)
                for alias in spec["aliases"]
            ],
        ]

        for candidate in candidates:

            normalized = normalize_menu_alias_text(
                candidate
            )

            if not normalized:
                continue

            existing = result.get(
                normalized
            )

            if (
                existing is not None
                and existing != canonical_key
            ):

                raise RuntimeError(
                    "MENU_ALIAS_COLLISION:"
                    f"ALIAS={normalized}:"
                    f"EXISTING={existing}:"
                    f"NEW={canonical_key}"
                )

            result[normalized] = canonical_key

    return result


ALIAS_TO_CANONICAL_KEY = _build_alias_index()


# ==================================================================================================
# PACK SIZE DETECTION
# ==================================================================================================

def detect_pack_size(
    value: object,
) -> int | None:

    text = normalize_menu_alias_text(
        value
    )

    if not text:
        return None


    # ----------------------------------------------------------------------------------------------
    # Explicit numeric size.
    # ----------------------------------------------------------------------------------------------

    numeric_match = re.search(
        r"(?<!\d)(1|4|5|10)\s*줄(?!\d)",
        text,
    )

    if numeric_match:

        return int(
            numeric_match.group(1)
        )


    # ----------------------------------------------------------------------------------------------
    # Base 꼬마김밥 size aliases.
    #
    # 소 -> 5
    # 대 -> 10
    # ----------------------------------------------------------------------------------------------

    compact = re.sub(
        r"\s+",
        "",
        text,
    )


    if "꼬마김밥" in compact:

        small_patterns = (
            r"\(소\)",
            r"\(소자\)",
            r"꼬마김밥소$",
            r"꼬마김밥소자$",
        )

        for pattern in small_patterns:

            if re.search(
                pattern,
                compact,
            ):
                return 5


        large_patterns = (
            r"\(대\)",
            r"\(대자\)",
            r"꼬마김밥대$",
            r"꼬마김밥대자$",
        )

        for pattern in large_patterns:

            if re.search(
                pattern,
                compact,
            ):
                return 10


    return None


# ==================================================================================================
# REVIEW CHECK
# ==================================================================================================

def is_review_required(
    value: object,
) -> bool:

    text = normalize_menu_alias_text(
        value
    )

    if not text:
        return False

    return any(
        pattern.search(text)
        for pattern in REVIEW_REQUIRED_PATTERNS
    )


# ==================================================================================================
# RESOLVER
# ==================================================================================================

def resolve_menu_alias(
    raw_name: object,
) -> MenuAliasResolution:

    raw = (
        ""
        if raw_name is None
        else str(raw_name)
    )

    normalized = normalize_menu_alias_text(
        raw
    )


    # ----------------------------------------------------------------------------------------------
    # EMPTY
    # ----------------------------------------------------------------------------------------------

    if not normalized:

        return MenuAliasResolution(
            raw_name=raw,
            normalized_name="",
            canonical_name=None,
            canonical_key=None,
            target_header=None,
            resolution_type="EMPTY",
            pack_size=None,
        )


    # ----------------------------------------------------------------------------------------------
    # REVIEW REQUIRED
    #
    # This happens BEFORE alias lookup.
    # ----------------------------------------------------------------------------------------------

    if is_review_required(
        normalized
    ):

        return MenuAliasResolution(
            raw_name=raw,
            normalized_name=normalized,
            canonical_name=None,
            canonical_key=None,
            target_header=None,
            resolution_type="REVIEW_REQUIRED",
            pack_size=detect_pack_size(
                normalized
            ),
        )


    # ----------------------------------------------------------------------------------------------
    # EXACT APPROVED ALIAS LOOKUP
    # ----------------------------------------------------------------------------------------------

    canonical_key = ALIAS_TO_CANONICAL_KEY.get(
        normalized
    )


    if canonical_key is None:

        return MenuAliasResolution(
            raw_name=raw,
            normalized_name=normalized,
            canonical_name=None,
            canonical_key=None,
            target_header=None,
            resolution_type="UNMAPPED",
            pack_size=detect_pack_size(
                normalized
            ),
        )


    spec = MENU_ALIAS_REGISTRY[
        canonical_key
    ]


    pack_size = (
        int(spec["pack_size"])
        if spec["pack_size"] is not None
        else detect_pack_size(
            normalized
        )
    )


    return MenuAliasResolution(
        raw_name=raw,
        normalized_name=normalized,
        canonical_name=str(
            spec["canonical_name"]
        ),
        canonical_key=canonical_key,
        target_header=str(
            spec["target_header"]
        ),
        resolution_type="ALIAS",
        pack_size=pack_size,
    )


# ==================================================================================================
# PUBLIC CONVENIENCE FUNCTIONS
# ==================================================================================================

def canonicalize_menu_name(
    raw_name: object,
) -> str:

    """
    Convert an approved alias into its canonical business name.

    Unknown or review-required values are NOT force-mapped.

    This behavior allows the existing production mapper to continue
    routing unknown values to OTHER_RESIDUAL.
    """

    resolution = resolve_menu_alias(
        raw_name
    )

    if resolution.canonical_name is not None:

        return resolution.canonical_name

    return resolution.normalized_name


def canonical_target_header(
    raw_name: object,
) -> str | None:

    resolution = resolve_menu_alias(
        raw_name
    )

    return resolution.target_header


def canonical_key_for_menu(
    raw_name: object,
) -> str | None:

    resolution = resolve_menu_alias(
        raw_name
    )

    return resolution.canonical_key


def registry_summary() -> dict[str, int]:

    return {

        "canonical_count": len(
            MENU_ALIAS_REGISTRY
        ),

        "alias_count": len(
            ALIAS_TO_CANONICAL_KEY
        ),

        "review_pattern_count": len(
            REVIEW_REQUIRED_PATTERNS
        ),
    }


# F11_SAFE_VARIANT_INDEX_EXTENSION_V2
#
# Deterministic orthographic alias expansion.
#
# This extension DOES NOT modify resolve_menu_alias().
# It extends the existing ALIAS_TO_CANONICAL_KEY index at module import time.
#
# Rules:
#   - no fuzzy matching
#   - no price inference
#   - no pack inference
#   - collisions are forbidden
#   - ambiguous menu names are intentionally excluded


def _f11_expand_safe_orthographic_forms(
    value: str,
) -> set[str]:
    import re
    import unicodedata

    normalized = unicodedata.normalize(
        "NFKC",
        str(value),
    )

    normalized = normalized.replace(
        "\u00a0",
        " ",
    )

    normalized = normalized.replace(
        "\u3000",
        " ",
    )

    normalized = re.sub(
        r"\s+",
        " ",
        normalized,
    ).strip()

    normalized = re.sub(
        r"\s*\(\s*",
        "(",
        normalized,
    )

    normalized = re.sub(
        r"\s*\)\s*",
        ")",
        normalized,
    )

    result: set[str] = {
        normalized,
    }

    compact = re.sub(
        r"\s+",
        "",
        normalized,
    )

    if compact:
        result.add(
            compact
        )

    # Approved lexical typo equivalence:
    #
    #   ??? <-> ???
    #
    current = tuple(result)

    for candidate in current:

        if "\ub5a1\ubcf6\uc774" in candidate:
            result.add(
                candidate.replace(
                    "\ub5a1\ubcf6\uc774",
                    "\ub5a1\ubcf6\uae30",
                )
            )

        if "\ub5a1\ubcf6\uae30" in candidate:
            result.add(
                candidate.replace(
                    "\ub5a1\ubcf6\uae30",
                    "\ub5a1\ubcf6\uc774",
                )
            )

    return {
        item
        for item in result
        if item
    }


# ================================================================================================
# Approved semantic seeds.
#
# These are business-equivalent labels only.
# No ambiguous names are included.
# ================================================================================================

F11_SAFE_VARIANT_SEEDS_V2: dict[
    str,
    tuple[str, ...],
] = {

    # --------------------------------------------------------------------------------------------
    # 1. ???? 5?
    # --------------------------------------------------------------------------------------------
    "KIMBAP_5": (
        "\uaf2c\ub9c8\uae40\ubc25 5\uc904",
        "\uaf2c\ub9c8\uae40\ubc255\uc904",
        "\uaf2c\ub9c8\uae40\ubc25 (\uc18c)",
        "\uaf2c\ub9c8\uae40\ubc25(\uc18c)",
        "기본5줄",
    ),

    # --------------------------------------------------------------------------------------------
    # 2. ???? 10?
    # --------------------------------------------------------------------------------------------
    "KIMBAP_10": (
        "\uaf2c\ub9c8\uae40\ubc25 10\uc904",
        "\uaf2c\ub9c8\uae40\ubc2510\uc904",
        "\uaf2c\ub9c8\uae40\ubc25 (\ub300)",
        "\uaf2c\ub9c8\uae40\ubc25(\ub300)",
        "기본10줄",
    ),

    # --------------------------------------------------------------------------------------------
    # 3. ???? 1?
    # --------------------------------------------------------------------------------------------
    "KIMBAP_1": (
        "\uaf2c\ub9c8\uae40\ubc25 1\uc904",
        "\uaf2c\ub9c8\uae40\ubc251\uc904",
        "기본1줄",
        "꼬마깁밥 1줄",
    ),

    # --------------------------------------------------------------------------------------------
    # 4. ??????? ???? 4?
    # --------------------------------------------------------------------------------------------
    "WASABI_CRAB_4": (
        "\uc640\uc0ac\ube44\ud06c\ub798\ub9c8\uc694 \uaf2c\ub9c8\uae40\ubc25",
        "\uc640\uc0ac\ube44\ud06c\ub798\ub9c8\uc694\uaf2c\ub9c8\uae40\ubc25",
        "\uc640\uc0ac\ube44\ud06c\ub798\ub9c8\uc694 \uaf2c\ub9c8\uae40\ubc25 4\uc904",
        "\uc640\uc0ac\ube44\ud06c\ub798\ub9c8\uc694\uaf2c\ub9c8\uae40\ubc254\uc904",
        "\uc640\uc0ac\ube44\ud06c\ub798\ub9c8\uc694 (4\uc904)",
        "\uc640\uc0ac\ube44\ud06c\ub798\ub9c8\uc694(4\uc904)",
    ),

    # --------------------------------------------------------------------------------------------
    # 5. ???? ???? 4?
    # CURRENT REPORTED CASE INCLUDED HERE.
    # --------------------------------------------------------------------------------------------
    "SPICY_JINMI_4": (
        "\ub9e4\ucf64\uc9c4\ubbf8 \uaf2c\ub9c8\uae40\ubc25",
        "\ub9e4\ucf64\uc9c4\ubbf8\uaf2c\ub9c8\uae40\ubc25",
        "\ub9e4\ucf64 \uc9c4\ubbf8 \uaf2c\ub9c8\uae40\ubc25",
        "\ub9e4\ucf64 \uc9c4\ubbf8\uaf2c\ub9c8\uae40\ubc25",

        "\ub9e4\ucf64\uc9c4\ubbf8 \uaf2c\ub9c8\uae40\ubc25 4\uc904",
        "\ub9e4\ucf64\uc9c4\ubbf8\uaf2c\ub9c8\uae40\ubc254\uc904",

        "\ub9e4\ucf64 \uc9c4\ubbf8 \uaf2c\ub9c8\uae40\ubc25 4\uc904",
        "\ub9e4\ucf64 \uc9c4\ubbf8\uaf2c\ub9c8\uae40\ubc254\uc904",
        "진미4줄",
    ),

    # --------------------------------------------------------------------------------------------
    # 6. ?? ???? 4?
    # --------------------------------------------------------------------------------------------
    "YUBU_4": (
        "\uc720\ubd80 \uaf2c\ub9c8\uae40\ubc25",
        "\uc720\ubd80\uaf2c\ub9c8\uae40\ubc25",
        "\uc720\ubd80 \uaf2c\ub9c8\uae40\ubc25 4\uc904",
        "\uc720\ubd80\uaf2c\ub9c8\uae40\ubc254\uc904",
        "\uc720\ubd80 \uaf2c\ub9c8\uae40\ubc25 (4\uc904)",
        "\uc720\ubd80\uaf2c\ub9c8\uae40\ubc25(4\uc904)",
    ),

    # --------------------------------------------------------------------------------------------
    # 7. ??? ???? 4?
    # --------------------------------------------------------------------------------------------
    "BUL_EOMUK_4": (
        "\ubd88\uc5b4\ubb35 \uaf2c\ub9c8\uae40\ubc25",
        "\ubd88\uc5b4\ubb35\uaf2c\ub9c8\uae40\ubc25",
        "\ubd88\uc5b4\ubb35 \uaf2c\ub9c8\uae40\ubc25 4\uc904",
        "\ubd88\uc5b4\ubb35\uaf2c\ub9c8\uae40\ubc254\uc904",
        "\ubd88\uc5b4\ubb35 \uaf2c\ub9c8\uae40\ubc25 (4\uc904)",
        "\ubd88\uc5b4\ubb35\uaf2c\ub9c8\uae40\ubc25(4\uc904)",
    ),

    # --------------------------------------------------------------------------------------------
    # 8. ???
    # --------------------------------------------------------------------------------------------
    "FISH_CAKE_SOUP": (
        "\uc5b4\ubb35\ud0d5",
        "\uc5b4\ubb35 \ud0d5",
    ),

    # --------------------------------------------------------------------------------------------
    # 9. ??? ???
    # --------------------------------------------------------------------------------------------
    "TTEOKBOKKI_MILD": (
        "\uc21c\ud55c\ub9db \ub5a1\ubcf6\uc774",
        "\uc21c\ud55c\ub9db\ub5a1\ubcf6\uc774",
        "\uc21c\ud55c\ub9db \ub5a1\ubcf6\uae30",
        "\uc21c\ud55c\ub9db\ub5a1\ubcf6\uae30",
        "\ub5a1\ubcf6\uc774(\uc21c)",
        "\ub5a1\ubcf6\uc774 (\uc21c)",
    ),

    # --------------------------------------------------------------------------------------------
    # 10. ??? ???
    # --------------------------------------------------------------------------------------------
    "TTEOKBOKKI_SPICY": (
        "\ub9e4\uc6b4\ub9db \ub5a1\ubcf6\uc774",
        "\ub9e4\uc6b4\ub9db\ub5a1\ubcf6\uc774",
        "\ub9e4\uc6b4\ub9db \ub5a1\ubcf6\uae30",
        "\ub9e4\uc6b4\ub9db\ub5a1\ubcf6\uae30",
        "\ub5a1\ubcf6\uc774(\ub9e4)",
        "\ub5a1\ubcf6\uc774 (\ub9e4)",
    ),

    # --------------------------------------------------------------------------------------------
    # 11. ?? ???
    # --------------------------------------------------------------------------------------------
    "JJOLMYEON_MILD": (
        "\uc21c\ud55c\ub9db \ucac4\uba74",
        "\uc21c\ud55c\ub9db\ucac4\uba74",
        "\ucac4\uba74(\uc21c)",
        "\ucac4\uba74 (\uc21c)",
    ),

    # --------------------------------------------------------------------------------------------
    # 12. ?? ???
    # --------------------------------------------------------------------------------------------
    "JJOLMYEON_SPICY": (
        "\ub9e4\uc6b4\ub9db \ucac4\uba74",
        "\ub9e4\uc6b4\ub9db\ucac4\uba74",
        "\ucac4\uba74(\ub9e4)",
        "\ucac4\uba74 (\ub9e4)",
    ),

    # --------------------------------------------------------------------------------------------
    # 13. ????
    # --------------------------------------------------------------------------------------------
    "SEONBI_UDON": (
        "\uc120\ube44\uc6b0\ub3d9",
        "\uc120\ube44 \uc6b0\ub3d9",
    ),

    # --------------------------------------------------------------------------------------------
    # 14. ?? ????
    # --------------------------------------------------------------------------------------------
    "SEONBI_KIMCHI_UDON": (
        "\uc120\ube44\uae40\uce58\uc6b0\ub3d9",
        "\uc120\ube44 \uae40\uce58\uc6b0\ub3d9",
        "\uc120\ube44\uae40\uce58 \uc6b0\ub3d9",
        "\uc120\ube44 \uae40\uce58 \uc6b0\ub3d9",
        "김치우동",
        "김치우동.",
    ),

    # --------------------------------------------------------------------------------------------
    # 15. ??
    # --------------------------------------------------------------------------------------------
    "UDON": (
        "\uc6b0\ub3d9",
    ),

    # --------------------------------------------------------------------------------------------
    # 16. ??
    # --------------------------------------------------------------------------------------------
    "RAMEN": (
        "\ub77c\uba74",
    ),

    # --------------------------------------------------------------------------------------------
    # 17. ????
    # --------------------------------------------------------------------------------------------
    "SRIRACHA": (
        "\uc2a4\ub9ac\ub77c\ucc28",
        "\uc2a4\ub9ac\ub77c\ucc28\uc18c\uc2a4",
        "\uc2a4\ub9ac\ub77c\ucc28 \uc18c\uc2a4",
        "\uc2a4\ub9ac\ub77c\ucc28\ub9c8\uc694\uc18c\uc2a4",
        "\uc2a4\ub9ac\ub77c\ucc28 \ub9c8\uc694\uc18c\uc2a4",
    ),

    # --------------------------------------------------------------------------------------------
    # 18. ????
    # --------------------------------------------------------------------------------------------
    "CHEONGYANG": (
        "\uccad\uc591\uace0\ucd94",
        "\uccad\uc591 \uace0\ucd94",
        "\uccad\uc591\uace0\ucd94\uc18c\uc2a4",
        "\uccad\uc591\uace0\ucd94 \uc18c\uc2a4",
        "\uccad\uc591\uc18c\uc2a4",
        "\uccad\uc591 \uc18c\uc2a4",
    ),

    # --------------------------------------------------------------------------------------------
    # 19. ?????
    # --------------------------------------------------------------------------------------------
    "MAKHANI_CURRY": (
        "\ub9c8\ud06c\ub2c8\ucee4\ub9ac",
        "\ub9c8\ud06c\ub2c8 \ucee4\ub9ac",
        "\ub9c8\ud06c\ub2c8\ucee4\ub9ac\uc18c\uc2a4",
        "\ub9c8\ud06c\ub2c8\ucee4\ub9ac \uc18c\uc2a4",
        "\ub9c8\ud06c\ub2c8 \ucee4\ub9ac \uc18c\uc2a4",
    ),

    # --------------------------------------------------------------------------------------------
    # 20. ??
    # --------------------------------------------------------------------------------------------
    "CHEESE": (
        "\uce58\uc988",
        "\uce58\uc988\uc18c\uc2a4",
        "\uce58\uc988 \uc18c\uc2a4",
    ),

    # --------------------------------------------------------------------------------------------
    # 21. ??
    # --------------------------------------------------------------------------------------------
    "SIKHYE": (
        "\uc2dd\ud61c",
        "\uc120\ube44\uc2dd\ud61c",
        "\uc120\ube44 \uc2dd\ud61c",
        "\uc218\uc81c\uc2dd\ud61c",
        "\uc218\uc81c \uc2dd\ud61c",

        "(HACCP \uc778\uc99d) \uc120\ube44\uc2dd\ud61c 150ml",
        "(HACCP \uc778\uc99d)\uc120\ube44\uc2dd\ud61c 150ml",
    ),
}


def _f11_register_safe_alias_v2(
    alias: str,
    canonical_key: str,
) -> None:

    normalized = normalize_menu_alias_text(
        alias
    )

    previous = ALIAS_TO_CANONICAL_KEY.get(
        normalized
    )

    if (
        previous is not None
        and previous != canonical_key
    ):
        raise RuntimeError(
            "F11_SAFE_ALIAS_COLLISION:"
            f"{normalized!r}:"
            f"{previous}:"
            f"{canonical_key}"
        )

    ALIAS_TO_CANONICAL_KEY[
        normalized
    ] = canonical_key


def _f11_install_safe_variant_index_v2() -> None:

    expected_keys = {
        "KIMBAP_5",
        "KIMBAP_10",
        "KIMBAP_1",
        "WASABI_CRAB_4",
        "SPICY_JINMI_4",
        "YUBU_4",
        "BUL_EOMUK_4",
        "FISH_CAKE_SOUP",
        "TTEOKBOKKI_MILD",
        "TTEOKBOKKI_SPICY",
        "JJOLMYEON_MILD",
        "JJOLMYEON_SPICY",
        "SEONBI_UDON",
        "SEONBI_KIMCHI_UDON",
        "UDON",
        "RAMEN",
        "SRIRACHA",
        "CHEONGYANG",
        "MAKHANI_CURRY",
        "CHEESE",
        "SIKHYE",
    }

    actual_keys = set(
        F11_SAFE_VARIANT_SEEDS_V2
    )

    if actual_keys != expected_keys:

        missing = (
            expected_keys
            - actual_keys
        )

        extra = (
            actual_keys
            - expected_keys
        )

        raise RuntimeError(
            "F11_SAFE_VARIANT_CANONICAL_SET_MISMATCH:"
            f"missing={sorted(missing)}:"
            f"extra={sorted(extra)}"
        )

    # --------------------------------------------------------------------------------------------
    # A. Expand all aliases already approved in MENU_ALIAS_REGISTRY.
    # --------------------------------------------------------------------------------------------

    for canonical_key, definition in MENU_ALIAS_REGISTRY.items():

        canonical_name = definition.get(
            "canonical_name"
        )

        source_aliases: list[str] = []

        if canonical_name:
            source_aliases.append(
                canonical_name
            )

        source_aliases.extend(
            definition.get(
                "aliases",
                (),
            )
        )

        for source_alias in source_aliases:

            for candidate in _f11_expand_safe_orthographic_forms(
                source_alias
            ):
                _f11_register_safe_alias_v2(
                    candidate,
                    canonical_key,
                )

    # --------------------------------------------------------------------------------------------
    # B. Add explicitly-approved semantic seeds for all 21 canonical groups.
    # --------------------------------------------------------------------------------------------

    for canonical_key, seeds in F11_SAFE_VARIANT_SEEDS_V2.items():

        if canonical_key not in MENU_ALIAS_REGISTRY:

            raise RuntimeError(
                "F11_SAFE_VARIANT_UNKNOWN_CANONICAL:"
                f"{canonical_key}"
            )

        for seed in seeds:

            for candidate in _f11_expand_safe_orthographic_forms(
                seed
            ):
                _f11_register_safe_alias_v2(
                    candidate,
                    canonical_key,
                )


_f11_install_safe_variant_index_v2()


