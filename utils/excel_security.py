# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

"""Security policy for values written to generated workbooks."""

import re
from typing import Any


_FORMULA_PREFIXES = ("=", "+", "-", "@", "\t", "\r", "\n")
_BLOCKED_FORMULA_TOKENS = re.compile(
    r"(?:\[|\]|\||\b(?:DDE|FILTERXML|HYPERLINK|RTD|WEBSERVICE)\s*\()",
    re.IGNORECASE,
)


def sanitize_cell_value(value: Any) -> Any:
    """Store formula-like untrusted strings as literal text."""
    if isinstance(value, str) and value.startswith(_FORMULA_PREFIXES):
        return "'" + value
    return value


def validate_formula(formula: str) -> str:
    """Validate an explicitly requested local Excel formula."""
    if not isinstance(formula, str) or not formula.strip():
        raise ValueError("Formula must be a non-empty string")
    normalized = formula if formula.startswith("=") else "=" + formula
    if len(normalized) > 8192 or any(character in normalized for character in "\r\n\0"):
        raise ValueError("Formula exceeds the allowed size or contains control characters")
    if _BLOCKED_FORMULA_TOKENS.search(normalized):
        raise ValueError("External, network, and link formulas are not allowed")
    return normalized