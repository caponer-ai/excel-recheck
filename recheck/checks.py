"""Перевірки: підсумок у файлі проти незалежного перерахунку з деталей."""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from enum import Enum
from pathlib import Path

from recheck.book import Book, Cell, Kind, RefError


class Status(Enum):
    OK = "зійшлось"
    DIFF = "розбіжність"
    SKIPPED = "не перевірено"


class SpecError(ValueError):
    """Файл правил записано з помилкою або він не підходить до цього звіту."""


@dataclass(frozen=True)
class Check:
    id: str
    group: str
    name: str
    total: str
    plus: tuple[str, ...]
    minus: tuple[str, ...] = ()
    scale: Decimal = Decimal(1)
    tol: Decimal = Decimal(0)
    note: str = ""


@dataclass(frozen=True)
class Result:
    check: Check
    status: Status
    theirs: Decimal | None
    mine: Decimal | None
    reason: str
    cells: tuple[Cell, ...]
    empty_parts: int = 0

    @property
    def diff(self) -> Decimal | None:
        if self.theirs is None or self.mine is None:
            return None
        return self.mine - self.theirs


@dataclass(frozen=True)
class Unrecognized:
    """Рядок таблиці з числами, який не є ні деталлю, ні групою, ні підсумком."""

    sheet: str
    row: int
    label: str


@dataclass(frozen=True)
class Table:
    id: str
    sheet: str
    first_row: int
    label_col: str
    total_label: str
    detail_regex: str
    sum_cols: tuple[str, ...]
    group_regex: str = ""
    rule_plus: tuple[str, ...] = ()
    rule_minus: tuple[str, ...] = ()
    rule_equals: str = ""


@dataclass
class Spec:
    checks: list[Check] = field(default_factory=list)
    tables: list[Table] = field(default_factory=list)


_BAD_KINDS = (Kind.TEXT, Kind.TEXT_NUMBER, Kind.FORMULA_NO_VALUE)


def evaluate(book: Book, check: Check) -> Result:
    """Одна перевірка. Усе, що не число, дає «не перевірено», а не нуль."""
    total_cells = book.cells(check.total)
    if len(total_cells) != 1:
        raise SpecError(f"{check.id}: підсумок має бути однією клітинкою, а не «{check.total}».")
    total = total_cells[0]
    plus = [cell for ref in check.plus for cell in book.cells(ref)]
    minus = [cell for ref in check.minus for cell in book.cells(ref)]
    parts = plus + minus
    used = (total, *parts)

    def skipped(reason: str) -> Result:
        return Result(check, Status.SKIPPED, total.value, None, reason, used)

    bad = [cell for cell in used if cell.kind in _BAD_KINDS]
    if bad:
        listed = "; ".join(f"{cell.ref}: {cell.kind.value} ({cell.raw!r})" for cell in bad)
        return skipped(listed)
    if total.kind is Kind.EMPTY:
        return skipped(f"{total.ref}: підсумкова клітинка порожня")
    if not parts:
        return skipped("немає жодного доданка")
    empty = sum(1 for cell in parts if cell.kind is Kind.EMPTY)
    if empty == len(parts):
        return skipped("усі доданки порожні, порівнювати нема з чим")

    def total_of(cells: list[Cell]) -> Decimal:
        return sum((cell.value for cell in cells if cell.value is not None), Decimal(0))

    mine = (total_of(plus) - total_of(minus)) * check.scale
    assert total.value is not None
    status = Status.OK if abs(mine - total.value) <= check.tol else Status.DIFF
    return Result(check, status, total.value, mine, "", used, empty)


def _label(cell: Cell) -> str:
    if cell.kind is Kind.EMPTY:
        return ""
    if cell.kind is Kind.NUMBER and cell.value is not None:
        return format(cell.value.normalize(), "f")
    return str(cell.raw).strip()


def expand_table(book: Book, table: Table) -> tuple[list[Check], list[Unrecognized]]:
    """Розкладає таблицю на перевірки рядків, груп і загального підсумку.

    Рядок-деталь і рядок-група розпізнаються за кодом, а не за назвою. Рядок з
    числами, який не підійшов під жодне правило, не пропускається мовчки.
    """
    detail = re.compile(table.detail_regex)
    group = re.compile(table.group_regex) if table.group_regex else None
    sheet = table.sheet
    checks: list[Check] = []
    unknown: list[Unrecognized] = []
    all_details: list[int] = []
    open_group: tuple[int, str, list[int]] | None = None
    groups: list[tuple[int, str, list[int]]] = []
    total_row: int | None = None

    def rule(row: int, label: str, kind: str) -> None:
        if not table.rule_equals:
            return
        left = " + ".join(table.rule_plus)
        right = "".join(f" − {col}" for col in table.rule_minus)
        checks.append(
            Check(
                id=f"{table.id}.rule.total" if kind == "total" else f"{table.id}.rule.{kind}.{row}",
                group=f"{sheet}: рух у рядку",
                name=f"р. {row} «{label}»: {left}{right} = {table.rule_equals}",
                total=f"{sheet}!{table.rule_equals}{row}",
                plus=tuple(f"{sheet}!{col}{row}" for col in table.rule_plus),
                minus=tuple(f"{sheet}!{col}{row}" for col in table.rule_minus),
            )
        )

    for row in range(table.first_row, book.max_row(sheet) + 1):
        label = _label(book.cell(sheet, f"{table.label_col}{row}"))
        if label == table.total_label:
            total_row = row
            break
        if detail.fullmatch(label):
            all_details.append(row)
            if open_group is not None:
                open_group[2].append(row)
            rule(row, label, "row")
        elif group is not None and group.fullmatch(label):
            open_group = (row, label, [])
            groups.append(open_group)
            rule(row, label, "group")
        else:
            has_numbers = any(
                book.cell(sheet, f"{col}{row}").kind in (Kind.NUMBER, Kind.TEXT_NUMBER)
                for col in table.sum_cols
            )
            if has_numbers:
                unknown.append(Unrecognized(sheet, row, label))

    if total_row is None:
        raise SpecError(
            f"{table.id}: на аркуші «{sheet}» не знайдено рядок «{table.total_label}» "
            f"у стовпчику {table.label_col}."
        )
    rule(total_row, table.total_label, "total")

    for row, label, members in groups:
        for col in table.sum_cols:
            checks.append(
                Check(
                    id=f"{table.id}.group.{label}.{col}",
                    group=f"{sheet}: підсумки груп",
                    name=f"Група «{label}», стовпчик {col} = сума рядків групи",
                    total=f"{sheet}!{col}{row}",
                    plus=tuple(f"{sheet}!{col}{member}" for member in members),
                )
            )
    for col in table.sum_cols:
        checks.append(
            Check(
                id=f"{table.id}.total.{col}",
                group=f"{sheet}: загальний підсумок",
                name=f"«{table.total_label}», стовпчик {col} = сума всіх рядків-деталей",
                total=f"{sheet}!{col}{total_row}",
                plus=tuple(f"{sheet}!{col}{member}" for member in all_details),
            )
        )
    return checks, unknown


_CHECK_KEYS = {"id", "group", "name", "total", "plus", "minus", "scale", "tol", "note"}
_TABLE_KEYS = {
    "id",
    "sheet",
    "first_row",
    "label_col",
    "total_label",
    "detail_regex",
    "group_regex",
    "sum_cols",
    "row_rule",
}
_RULE_KEYS = {"plus", "minus", "equals"}


def _decimal(value: object, where: str) -> Decimal:
    try:
        return Decimal(str(value))
    except InvalidOperation:
        raise SpecError(f"{where}: «{value}» не є числом.") from None


def _keys(item: dict[str, object], allowed: set[str], required: set[str], where: str) -> None:
    extra = set(item) - allowed
    if extra:
        raise SpecError(f"{where}: невідомі поля {sorted(extra)}.")
    missing = required - set(item)
    if missing:
        raise SpecError(f"{where}: бракує полів {sorted(missing)}.")


def parse_spec(data: object) -> Spec:
    if not isinstance(data, dict) or set(data) - {"checks", "tables"}:
        raise SpecError("Файл правил: очікую об'єкт лише з полями «checks» і «tables».")
    spec = Spec()
    seen: set[str] = set()

    def unique(ident: str) -> str:
        if ident in seen:
            raise SpecError(f"Файл правил: id «{ident}» повторюється.")
        seen.add(ident)
        return ident

    for number, item in enumerate(data.get("checks", []), start=1):
        where = f"checks[{number}]"
        _keys(item, _CHECK_KEYS, {"id", "group", "name", "total", "plus"}, where)
        spec.checks.append(
            Check(
                id=unique(str(item["id"])),
                group=str(item["group"]),
                name=str(item["name"]),
                total=str(item["total"]),
                plus=tuple(item["plus"]),
                minus=tuple(item.get("minus", ())),
                scale=_decimal(item.get("scale", "1"), where),
                tol=_decimal(item.get("tol", "0"), where),
                note=str(item.get("note", "")),
            )
        )
    for number, item in enumerate(data.get("tables", []), start=1):
        where = f"tables[{number}]"
        required = _TABLE_KEYS - {"group_regex", "row_rule"}
        _keys(item, _TABLE_KEYS, required, where)
        row_rule = item.get("row_rule") or {}
        if row_rule:
            _keys(row_rule, _RULE_KEYS, {"plus", "equals"}, f"{where}.row_rule")
        spec.tables.append(
            Table(
                id=unique(str(item["id"])),
                sheet=str(item["sheet"]),
                first_row=int(item["first_row"]),
                label_col=str(item["label_col"]),
                total_label=str(item["total_label"]),
                detail_regex=str(item["detail_regex"]),
                group_regex=str(item.get("group_regex", "")),
                sum_cols=tuple(item["sum_cols"]),
                rule_plus=tuple(row_rule.get("plus", ())),
                rule_minus=tuple(row_rule.get("minus", ())),
                rule_equals=str(row_rule.get("equals", "")),
            )
        )
    if not spec.checks and not spec.tables:
        raise SpecError("Файл правил порожній: немає жодної перевірки.")
    return spec


def load_spec(path: str | Path) -> Spec:
    try:
        data = json.loads(Path(path).read_text(encoding="utf-8"))
    except json.JSONDecodeError as error:
        raise SpecError(f"Файл правил {path}: це не JSON ({error}).") from None
    return parse_spec(data)


@dataclass
class Run:
    results: list[Result]
    unrecognized: list[Unrecognized]
    uncovered: dict[str, list[Cell]]
    numeric_total: int


def run(book: Book, spec: Spec) -> Run:
    """Виконує всі перевірки і рахує, яких чисел у файлі вони не торкнулись."""
    checks = list(spec.checks)
    unrecognized: list[Unrecognized] = []
    try:
        for table in spec.tables:
            table_checks, unknown = expand_table(book, table)
            checks.extend(table_checks)
            unrecognized.extend(unknown)
        results = [evaluate(book, check) for check in checks]
    except RefError as error:
        raise SpecError(str(error)) from None

    touched = {cell.ref for result in results for cell in result.cells}
    labels: set[str] = set()  # коди рядків це підписи, а не суми
    for table in spec.tables:
        for row in range(table.first_row, book.max_row(table.sheet) + 1):
            labels.add(f"{table.sheet}!{table.label_col}{row}")
    uncovered: dict[str, list[Cell]] = {}
    numeric_total = 0
    for cell in book.numeric_cells():
        if cell.ref in labels:
            continue
        numeric_total += 1
        if cell.ref not in touched:
            uncovered.setdefault(cell.sheet, []).append(cell)
    return Run(results, unrecognized, uncovered, numeric_total)
