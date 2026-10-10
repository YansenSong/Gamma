"""Lossless-enough text and unit normalization for evidence indexing."""

from __future__ import annotations

import re
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation


_FULLWIDTH = str.maketrans({chr(0xFF10 + i): str(i) for i in range(10)} | {
    "％": "%", "．": ".", "，": ",", "：": ":", "（": "(", "）": ")",
    "－": "-", "−": "-", "＋": "+", "＝": "=", "＞": ">", "＜": "<",
})
_UNITS = re.compile(r"(?<![A-Za-z])(?:km/h|m/s²|m/s2|dB\(A\)|kHz|Hz|ms|lx|dB|s|m|g|%)(?![A-Za-z])")


@dataclass(frozen=True)
class UnitToken:
    raw: str
    normalized: str
    factor: str


def normalize_text(raw: str) -> str:
    """Keep comparison operators intact and never replace source evidence."""
    value = raw.translate(_FULLWIDTH).replace("\u00a0", " ").replace("\u200b", "")
    return re.sub(r"\s+", " ", value).strip()


def parse_decimal(raw: str) -> Decimal | None:
    token = normalize_text(raw).replace(",", "")
    token = re.sub(r"\s*(?:分|次|条)$", "", token)
    try:
        if token.endswith("%"):
            return Decimal(token[:-1].strip()) / Decimal(100)
        return Decimal(token)
    except (InvalidOperation, ValueError):
        return None


def unit_tokens(raw: str) -> tuple[UnitToken, ...]:
    tokens = []
    for match in _UNITS.finditer(normalize_text(raw)):
        unit = match.group()
        normalized = {"m/s2": "m/s²", "kHz": "Hz", "%": "ratio"}.get(unit, unit)
        factor = "1000" if unit == "kHz" else "0.01" if unit == "%" else "1"
        tokens.append(UnitToken(unit, normalized, factor))
    return tuple(tokens)
