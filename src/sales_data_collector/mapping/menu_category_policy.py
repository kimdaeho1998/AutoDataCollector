from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from enum import Enum


# ==================================================================================================
# MENU_MONTHLY CATEGORY POLICY
# ==================================================================================================
#
# MagicERP [??] field is a SOURCE-OWNERSHIP boundary.
#
# Only classifications that are positively identified as menu-oriented are
# allowed into MENU_MONTHLY.
#
# Unknown / delivery / online / platform / miscellaneous classifications
# are fail-closed and excluded from MENU_MONTHLY.
#
# Excluded records must NOT enter:
#
#   - canonical menu sales
#   - canonical menu quantity
#   - OTHER_RESIDUAL
#   - OTHER_NEW_MENU
#   - SOURCE_TOTAL
#   - SOURCE_TOTAL_QUANTITY
#
# ==================================================================================================


class MenuCategoryDecision(str, Enum):
    INCLUDE = "INCLUDE"
    EXCLUDE = "EXCLUDE"


@dataclass(frozen=True)
class MenuCategoryResolution:
    raw_value: str
    normalized_value: str
    canonical_tokens: tuple[str, ...]
    decision: MenuCategoryDecision
    reason: str


# ==================================================================================================
# Canonical category tokens
# ==================================================================================================

TOKEN_MENU = "\uba54\ub274"
TOKEN_KIMBAP = "\uae40\ubc25"
TOKEN_SIDE = "\uc0ac\uc774\ub4dc"
TOKEN_SAUCE = "\uc18c\uc2a4"
TOKEN_DRINK = "\uc74c\ub8cc"
TOKEN_BUNSIK = "\ubd84\uc2dd"
TOKEN_MEAL = "\uc2dd\uc0ac"


ALLOWED_CANONICAL_TOKENS: frozenset[str] = frozenset(
    {
        TOKEN_MENU,
        TOKEN_KIMBAP,
        TOKEN_SIDE,
        TOKEN_SAUCE,
        TOKEN_DRINK,
        TOKEN_BUNSIK,
        TOKEN_MEAL,
    }
)


# ==================================================================================================
# Token aliases
#
# Left side  = normalized raw token
# Right side = canonical category token
# ==================================================================================================

CATEGORY_TOKEN_ALIASES: dict[str, str] = {

    # ----------------------------------------------------------------------------------------------
    # ??
    # ----------------------------------------------------------------------------------------------

    "\uba54\ub274": TOKEN_MENU,
    "\uba54\ub274\ub958": TOKEN_MENU,
    "\uc77c\ubc18\uba54\ub274": TOKEN_MENU,
    "\uc77c\ubc18 \uba54\ub274": TOKEN_MENU,

    # ----------------------------------------------------------------------------------------------
    # ??
    # ----------------------------------------------------------------------------------------------

    "\uae40\ubc25": TOKEN_KIMBAP,
    "\uae40\ubc25\ub958": TOKEN_KIMBAP,
    "\uaf2c\ub9c8\uae40\ubc25": TOKEN_KIMBAP,
    "\uaf2c\ub9c8 \uae40\ubc25": TOKEN_KIMBAP,
    "\uaf2c\ub9c8\uae40\ubc25\ub958": TOKEN_KIMBAP,
    "\uaf2c\ub9c8 \uae40\ubc25\ub958": TOKEN_KIMBAP,

    # ----------------------------------------------------------------------------------------------
    # ???
    # ----------------------------------------------------------------------------------------------

    "\uc0ac\uc774\ub4dc": TOKEN_SIDE,
    "\uc0ac\uc774\ub4dc\ub958": TOKEN_SIDE,
    "\uc0ac\uc774\ub4dc\uba54\ub274": TOKEN_SIDE,
    "\uc0ac\uc774\ub4dc \uba54\ub274": TOKEN_SIDE,
    "\uc0ac\uc774\ub4dc\uba54\ub274\ub958": TOKEN_SIDE,
    "\uc0ac\uc774\ub4dc \uba54\ub274\ub958": TOKEN_SIDE,

    # ----------------------------------------------------------------------------------------------
    # ??
    # ----------------------------------------------------------------------------------------------

    "\uc18c\uc2a4": TOKEN_SAUCE,
    "\uc18c\uc2a4\ub958": TOKEN_SAUCE,
    "\uc18c\uc2a4 \ub958": TOKEN_SAUCE,

    # ----------------------------------------------------------------------------------------------
    # ??
    # ----------------------------------------------------------------------------------------------

    "\uc74c\ub8cc": TOKEN_DRINK,
    "\uc74c\ub8cc\ub958": TOKEN_DRINK,
    "\uc74c\ub8cc \ub958": TOKEN_DRINK,
    "\uc74c\ub8cc\uc218": TOKEN_DRINK,
    "\uc74c\ub8cc\uc218\ub958": TOKEN_DRINK,
    "\uc74c\ub8cc\uc218 \ub958": TOKEN_DRINK,

    # ----------------------------------------------------------------------------------------------
    # ??
    # ----------------------------------------------------------------------------------------------

    "\ubd84\uc2dd": TOKEN_BUNSIK,
    "\ubd84\uc2dd\ub958": TOKEN_BUNSIK,
    "\ubd84\uc2dd \ub958": TOKEN_BUNSIK,
    "\ubd84\uc2dd\uba54\ub274": TOKEN_BUNSIK,
    "\ubd84\uc2dd \uba54\ub274": TOKEN_BUNSIK,
    "\ubd84\uc2dd\uba54\ub274\ub958": TOKEN_BUNSIK,

    # ----------------------------------------------------------------------------------------------
    # ??
    # ----------------------------------------------------------------------------------------------

    "\uc2dd\uc0ac": TOKEN_MEAL,
    "\uc2dd\uc0ac\ub958": TOKEN_MEAL,
    "\uc2dd\uc0ac \ub958": TOKEN_MEAL,
    "\uc2dd\uc0ac\uba54\ub274": TOKEN_MEAL,
    "\uc2dd\uc0ac \uba54\ub274": TOKEN_MEAL,
}


# ==================================================================================================
# Explicitly confirmed MENU_MONTHLY category strings
#
# These are kept as an audit/reference contract even though the tokenizer can
# recognize their equivalent normalized forms.
# ==================================================================================================

CONFIRMED_MENU_CATEGORIES: frozenset[str] = frozenset(
    {
        "\uba54\ub274",
        "\uba54\ub274.",

        "\uae40\ubc25 / \uc0ac\uc774\ub4dc / \uc18c\uc2a4\ub958",
        "\uae40\ubc25 & \uc2dd\uc0ac",
        "\uae40\ubc25/\uc0ac\uc774\ub4dc",
        "\uae40\ubc25/\uc18c\uc2a4\ub958",
        "\uae40\ubc25/\uc18c\uc2a4",
        "\uae40\ubc25 \ubc0f \uc18c\uc2a4",
        "\uaf2c\ub9c8\uae40\ubc25",
        "\uae40\ubc25 \ubc0f \uc18c\uc2a4\ub958",

        "\uc0ac\uc774\ub4dc",

        "\uc74c\ub8cc\ub958",
        "\uc74c\ub8cc",

        "\uc0ac\uc774\ub4dc \ubc0f \uc74c\ub8cc",

        "\uc18c\uc2a4/\uc74c\ub8cc",
        "\uc18c\uc2a4",
        "\uc18c\uc2a4\ub958",

        "\uc0ac\uc774\ub4dc\uba54\ub274",
        "\uc0ac\uc774\ub4dc\ub958",
        "\uc0ac\uc774\ub4dc/\uc74c\ub8cc",

        "\ubd84\uc2dd\ub958",
    }
)


# ==================================================================================================
# Approved semantic compositions
#
# IMPORTANT:
#
# We intentionally do NOT allow every mathematically possible combination of
# category tokens.
#
# Only menu-domain combinations that make operational sense are allowed.
#
# This avoids accidentally admitting unexpected source categories.
# ==================================================================================================

ALLOWED_TOKEN_COMBINATIONS: frozenset[frozenset[str]] = frozenset(
    {

        # Single category
        frozenset({TOKEN_MENU}),
        frozenset({TOKEN_KIMBAP}),
        frozenset({TOKEN_SIDE}),
        frozenset({TOKEN_SAUCE}),
        frozenset({TOKEN_DRINK}),
        frozenset({TOKEN_BUNSIK}),
        frozenset({TOKEN_MEAL}),

        # Two-category combinations
        frozenset({TOKEN_KIMBAP, TOKEN_SIDE}),
        frozenset({TOKEN_KIMBAP, TOKEN_SAUCE}),
        frozenset({TOKEN_KIMBAP, TOKEN_MEAL}),
        frozenset({TOKEN_KIMBAP, TOKEN_DRINK}),
        frozenset({TOKEN_KIMBAP, TOKEN_BUNSIK}),

        frozenset({TOKEN_SIDE, TOKEN_SAUCE}),
        frozenset({TOKEN_SIDE, TOKEN_DRINK}),
        frozenset({TOKEN_SIDE, TOKEN_BUNSIK}),
        frozenset({TOKEN_SIDE, TOKEN_MEAL}),

        frozenset({TOKEN_SAUCE, TOKEN_DRINK}),

        frozenset({TOKEN_BUNSIK, TOKEN_DRINK}),
        frozenset({TOKEN_BUNSIK, TOKEN_MEAL}),

        # Three-category combinations
        frozenset(
            {
                TOKEN_KIMBAP,
                TOKEN_SIDE,
                TOKEN_SAUCE,
            }
        ),

        frozenset(
            {
                TOKEN_KIMBAP,
                TOKEN_SIDE,
                TOKEN_DRINK,
            }
        ),

        frozenset(
            {
                TOKEN_KIMBAP,
                TOKEN_SIDE,
                TOKEN_MEAL,
            }
        ),

        frozenset(
            {
                TOKEN_KIMBAP,
                TOKEN_SAUCE,
                TOKEN_DRINK,
            }
        ),

        frozenset(
            {
                TOKEN_SIDE,
                TOKEN_SAUCE,
                TOKEN_DRINK,
            }
        ),

        frozenset(
            {
                TOKEN_KIMBAP,
                TOKEN_BUNSIK,
                TOKEN_DRINK,
            }
        ),

        # Four-category menu grouping
        frozenset(
            {
                TOKEN_KIMBAP,
                TOKEN_SIDE,
                TOKEN_SAUCE,
                TOKEN_DRINK,
            }
        ),
    }
)


# ==================================================================================================
# Normalization
# ==================================================================================================


def normalize_menu_category(
    value: str | None,
) -> str:
    """Normalize MagicERP [??] text without changing its business meaning."""

    if value is None:
        return ""

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

    # ERP-entry punctuation variation.
    normalized = normalized.rstrip(
        "."
    ).strip()

    # Normalize equivalent category separators.
    normalized = re.sub(
        r"\s*/\s*",
        "/",
        normalized,
    )

    normalized = re.sub(
        r"\s*&\s*",
        "/",
        normalized,
    )

    normalized = re.sub(
        r"\s*\ubc0f\s*",
        "/",
        normalized,
    )

    # Re-collapse duplicate separators.
    normalized = re.sub(
        r"/+",
        "/",
        normalized,
    )

    normalized = normalized.strip(
        "/ "
    )

    return normalized


def _normalize_category_token(
    token: str,
) -> str:
    normalized = unicodedata.normalize(
        "NFKC",
        token,
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

    return normalized


# ==================================================================================================
# Classification
# ==================================================================================================


def resolve_menu_category(
    value: str | None,
) -> MenuCategoryResolution:

    raw_value = (
        ""
        if value is None
        else str(value)
    )

    normalized = normalize_menu_category(
        value
    )

    # Fail closed for blank category.
    if not normalized:

        return MenuCategoryResolution(
            raw_value=raw_value,
            normalized_value=normalized,
            canonical_tokens=(),
            decision=MenuCategoryDecision.EXCLUDE,
            reason="EMPTY_CATEGORY",
        )

    raw_tokens = tuple(
        token
        for token in normalized.split("/")
        if token
    )

    if not raw_tokens:

        return MenuCategoryResolution(
            raw_value=raw_value,
            normalized_value=normalized,
            canonical_tokens=(),
            decision=MenuCategoryDecision.EXCLUDE,
            reason="EMPTY_CATEGORY",
        )

    canonical_tokens: list[str] = []

    for raw_token in raw_tokens:

        normalized_token = _normalize_category_token(
            raw_token
        )

        canonical = CATEGORY_TOKEN_ALIASES.get(
            normalized_token
        )

        if canonical is None:

            return MenuCategoryResolution(
                raw_value=raw_value,
                normalized_value=normalized,
                canonical_tokens=tuple(
                    canonical_tokens
                ),
                decision=MenuCategoryDecision.EXCLUDE,
                reason=(
                    "UNKNOWN_CATEGORY_TOKEN:"
                    + normalized_token
                ),
            )

        canonical_tokens.append(
            canonical
        )

    # Remove duplicates while preserving order.
    canonical_tokens = list(
        dict.fromkeys(
            canonical_tokens
        )
    )

    token_set = frozenset(
        canonical_tokens
    )

    if token_set not in ALLOWED_TOKEN_COMBINATIONS:

        return MenuCategoryResolution(
            raw_value=raw_value,
            normalized_value=normalized,
            canonical_tokens=tuple(
                canonical_tokens
            ),
            decision=MenuCategoryDecision.EXCLUDE,
            reason="UNAPPROVED_CATEGORY_COMBINATION",
        )

    return MenuCategoryResolution(
        raw_value=raw_value,
        normalized_value=normalized,
        canonical_tokens=tuple(
            canonical_tokens
        ),
        decision=MenuCategoryDecision.INCLUDE,
        reason="MENU_CATEGORY_WHITELIST",
    )


def is_allowed_menu_category(
    value: str | None,
) -> bool:

    return (
        resolve_menu_category(
            value
        ).decision
        == MenuCategoryDecision.INCLUDE
    )
