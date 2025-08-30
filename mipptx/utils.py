from __future__ import annotations

import re
from typing import Iterable


EMU_PER_INCH = 914400
EMU_PER_CM = 360000
EMU_PER_PT = 12700


def pt_to_emu(pt: float | int | None) -> int | None:
    if pt is None:
        return None
    return int(round(float(pt) * EMU_PER_PT))


def emu_to_pt(emu: int | None) -> float | None:
    if emu is None:
        return None
    return float(emu) / EMU_PER_PT


def inches_to_emu(inches: float | int | None) -> int | None:
    if inches is None:
        return None
    return int(round(float(inches) * EMU_PER_INCH))


def emu_to_inches(emu: int | None) -> float | None:
    if emu is None:
        return None
    return float(emu) / EMU_PER_INCH


def hex_color(color: str | None) -> str | None:
    if color is None:
        return None
    # distinguish empty-string from whitespace-only for test expectations
    if color == "":
        raise ValueError("invalid hex color: empty string")
    c = color.strip()
    if not c:
        return None
    if not c.startswith("#"):
        c = "#" + c
    if not re.fullmatch(r"#([0-9a-fA-F]{6})", c):
        raise ValueError(f"invalid hex color: {color}")
    return c.lower()


def unique(seq: Iterable) -> list:
    seen = set()
    out = []
    for item in seq:
        if item in seen:
            continue
        seen.add(item)
        out.append(item)
    return out
