"""Читання клітинок Excel так, щоб нечислове ніколи мовчки не ставало нулем.

Кожна клітинка повертається разом з адресою і видом значення. Суми рахуються
в Decimal, тому «копійка в копійку» означає саме це, без шуму від float.
"""

from __future__ import annotations

import re
from collections.abc import Iterator
from dataclasses import dataclass
from decimal import Decimal
from enum import Enum
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, range_boundaries


class Kind(Enum):
    NUMBER = "число"
    EMPTY = "порожньо"
    TEXT = "текст"
    TEXT_NUMBER = "число, записане текстом"
    FORMULA_NO_VALUE = "формула без збереженого значення"


class RefError(ValueError):
    """Посилання в правилах не відповідає файлу або записане з помилкою."""


_TEXT_NUMBER = re.compile(r"-?\d+(?:[.,]\d+)?")
_SPACES = re.compile(r"[\s  ]")
_REF = re.compile(r"(?P<sheet>.+)!(?P<first>[A-Z]{1,3}[1-9]\d*)(?::(?P<last>[A-Z]{1,3}[1-9]\d*))?")


@dataclass(frozen=True)
class Cell:
    sheet: str
    addr: str
    kind: Kind
    value: Decimal | None
    raw: object

    @property
    def ref(self) -> str:
        return f"{self.sheet}!{self.addr}"


def parse_ref(ref: str) -> tuple[str, list[str]]:
    """«Аркуш!C4» або «Аркуш!C4:C10» -> (аркуш, список адрес по рядках)."""
    match = _REF.fullmatch(ref.strip())
    if match is None:
        raise RefError(
            f"Не можу прочитати посилання «{ref}». Очікую «Аркуш!C4» або «Аркуш!C4:C10»."
        )
    first, last = match["first"], match["last"] or match["first"]
    min_col, min_row, max_col, max_row = range_boundaries(f"{first}:{last}")
    addrs = [
        f"{get_column_letter(col)}{row}"
        for row in range(min_row, max_row + 1)
        for col in range(min_col, max_col + 1)
    ]
    return match["sheet"], addrs


def classify(value: object, raw: object = None) -> tuple[Kind, Decimal | None]:
    """Вид значення клітинки і число, якщо воно там є."""
    if value is None:
        if isinstance(raw, str) and raw.startswith("="):
            return Kind.FORMULA_NO_VALUE, None
        return Kind.EMPTY, None
    if isinstance(value, bool):  # bool це підклас int, перевіряємо раніше
        return Kind.TEXT, None
    if isinstance(value, int):
        return Kind.NUMBER, Decimal(value)
    if isinstance(value, float):
        if value != value or value in (float("inf"), float("-inf")):
            return Kind.TEXT, None
        return Kind.NUMBER, Decimal(repr(value))
    if isinstance(value, str):
        compact = _SPACES.sub("", value)
        if not compact:
            return Kind.EMPTY, None
        if _TEXT_NUMBER.fullmatch(compact):
            return Kind.TEXT_NUMBER, Decimal(compact.replace(",", "."))
        return Kind.TEXT, None
    return Kind.TEXT, None  # дати та інше: для суми це не число


class Book:
    """Файл .xlsx, відкритий лише для читання."""

    def __init__(self, path: str | Path) -> None:
        self.path = Path(path)
        self._values = load_workbook(self.path, data_only=True)
        self._formulas = load_workbook(self.path, data_only=False)

    @property
    def sheets(self) -> list[str]:
        return list(self._values.sheetnames)

    def _require(self, sheet: str) -> None:
        if sheet not in self._values.sheetnames:
            known = ", ".join(f"«{name}»" for name in self._values.sheetnames)
            raise RefError(f"Аркуша «{sheet}» немає у файлі {self.path.name}. Є: {known}.")

    def cell(self, sheet: str, addr: str) -> Cell:
        self._require(sheet)
        value = self._values[sheet][addr].value
        raw = self._formulas[sheet][addr].value
        kind, number = classify(value, raw)
        return Cell(sheet, addr, kind, number, value if value is not None else raw)

    def cells(self, ref: str) -> list[Cell]:
        sheet, addrs = parse_ref(ref)
        return [self.cell(sheet, addr) for addr in addrs]

    def max_row(self, sheet: str) -> int:
        self._require(sheet)
        return int(self._values[sheet].max_row)

    def numeric_cells(self) -> Iterator[Cell]:
        """Усі клітинки, де стоїть число (також число, записане текстом)."""
        for sheet in self._values.sheetnames:
            for row in self._values[sheet].iter_rows():
                for item in row:
                    if item.value is None:
                        continue
                    kind, number = classify(item.value)
                    if kind in (Kind.NUMBER, Kind.TEXT_NUMBER):
                        yield Cell(sheet, item.coordinate, kind, number, item.value)
