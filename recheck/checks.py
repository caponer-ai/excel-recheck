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
    reason: str = "не підходить під жодне правило"


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
    if total.ref in {cell.ref for cell in parts}:
        raise SpecError(f"{check.id}: підсумок {total.ref} стоїть серед власних доданків.")

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
    largest = max(abs(mine), abs(total.value))
    if check.tol > 0 and 0 < largest <= check.tol:
        # допуск «1e30» скінченний, але з ним зійдеться що завгодно
        return skipped(
            f"допуск {check.tol:f} не менший за самі числа, порівняння нічого не доводить"
        )
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

    watched = dict.fromkeys(
        (*table.sum_cols, *table.rule_plus, *table.rule_minus, table.rule_equals)
    )
    watched.pop("", None)

    def has_content(row: int) -> bool:
        # будь-що непорожнє: текст «(1 234,56)» чи формула без значення теж не зникають мовчки
        return any(book.cell(sheet, f"{col}{row}").kind is not Kind.EMPTY for col in watched)

    for row in range(table.first_row, book.max_row(sheet) + 1):
        label = _label(book.cell(sheet, f"{table.label_col}{row}"))
        if total_row is not None:
            # усе, що йде після підсумку, у підсумок не входить: мовчки не пропускаємо
            if label or has_content(row):
                unknown.append(Unrecognized(sheet, row, label, "стоїть після підсумкового рядка"))
        elif label == table.total_label:
            total_row = row
        elif detail.fullmatch(label):
            all_details.append(row)
            if open_group is not None:
                open_group[2].append(row)
                tail = label.removeprefix(open_group[1])
                if tail == label or tail[:1].isalnum():  # «2015.002» не належить групі «201»
                    reason = f"код не починається з коду групи «{open_group[1]}», у якій стоїть"
                    unknown.append(Unrecognized(sheet, row, label, reason))
            elif group is not None:
                unknown.append(
                    Unrecognized(sheet, row, label, "рядок-деталь стоїть до першої групи")
                )
            rule(row, label, "row")
        elif group is not None and group.fullmatch(label):
            open_group = (row, label, [])
            groups.append(open_group)
            rule(row, label, "group")
        elif has_content(row):
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
        number = Decimal(str(value))
    except InvalidOperation:
        raise SpecError(f"{where}: «{value}» не є числом.") from None
    if not number.is_finite():
        raise SpecError(f"{where}: «{value}» не є скінченним числом.")
    return number


def _objects(value: object, where: str) -> list[dict[str, object]]:
    if not isinstance(value, list) or not all(isinstance(item, dict) for item in value):
        raise SpecError(f"{where}: очікую список об'єктів.")
    return value


def _strings(value: object, where: str) -> tuple[str, ...]:
    if not isinstance(value, list) or not all(isinstance(item, str) for item in value):
        raise SpecError(f"{where}: очікую список рядків.")
    return tuple(value)


_COLUMN = re.compile(r"[A-Z]{1,3}")


def _column(value: object, where: str) -> str:
    if not isinstance(value, str) or not _COLUMN.fullmatch(value):
        raise SpecError(f"{where}: «{value}» не є літерою стовпчика. Очікую A, B або AA.")
    return value


def _columns(value: object, where: str) -> tuple[str, ...]:
    return tuple(_column(item, where) for item in _strings(value, where))


def _regex(value: object, where: str) -> str:
    try:
        re.compile(str(value))
    except re.error as error:
        raise SpecError(f"{where}: «{value}» не є регулярним виразом ({error}).") from None
    return str(value)


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

    for number, item in enumerate(_objects(data.get("checks", []), "checks"), start=1):
        where = f"checks[{number}]"
        _keys(item, _CHECK_KEYS, {"id", "group", "name", "total", "plus"}, where)
        if _decimal(item.get("tol", "0"), where) < 0:
            raise SpecError(f"{where}: допуск не може бути від'ємним.")
        if _decimal(item.get("scale", "1"), where) == 0:
            raise SpecError(f"{where}: множник не може бути нулем.")
        spec.checks.append(
            Check(
                id=unique(str(item["id"])),
                group=str(item["group"]),
                name=str(item["name"]),
                total=str(item["total"]),
                plus=_strings(item["plus"], f"{where}.plus"),
                minus=_strings(item.get("minus", []), f"{where}.minus"),
                scale=_decimal(item.get("scale", "1"), where),
                tol=_decimal(item.get("tol", "0"), where),
                note=str(item.get("note", "")),
            )
        )
    for number, item in enumerate(_objects(data.get("tables", []), "tables"), start=1):
        where = f"tables[{number}]"
        required = _TABLE_KEYS - {"group_regex", "row_rule"}
        _keys(item, _TABLE_KEYS, required, where)
        row_rule = item.get("row_rule") or {}
        if not isinstance(row_rule, dict):
            raise SpecError(f"{where}.row_rule: очікую об'єкт.")
        if row_rule:
            _keys(row_rule, _RULE_KEYS, {"plus", "equals"}, f"{where}.row_rule")
        first_row = item["first_row"]
        if isinstance(first_row, bool) or not isinstance(first_row, int) or first_row < 1:
            raise SpecError(f"{where}.first_row: очікую номер рядка, ціле число від 1.")
        sum_cols = _columns(item["sum_cols"], f"{where}.sum_cols")
        rule_equals = _column(row_rule["equals"], f"{where}.row_rule.equals") if row_rule else ""
        if not sum_cols and not row_rule:
            raise SpecError(f"{where}: немає ні «sum_cols», ні «row_rule», перевіряти нічого.")
        spec.tables.append(
            Table(
                id=unique(str(item["id"])),
                sheet=str(item["sheet"]),
                first_row=first_row,
                label_col=_column(item["label_col"], f"{where}.label_col"),
                total_label=str(item["total_label"]),
                detail_regex=_regex(item["detail_regex"], f"{where}.detail_regex"),
                group_regex=_regex(item.get("group_regex", ""), f"{where}.group_regex"),
                sum_cols=sum_cols,
                rule_plus=_columns(row_rule.get("plus", []), f"{where}.row_rule.plus"),
                rule_minus=_columns(row_rule.get("minus", []), f"{where}.row_rule.minus"),
                rule_equals=rule_equals,
            )
        )
    if not spec.checks and not spec.tables:
        raise SpecError("Файл правил порожній: немає жодної перевірки.")
    return spec


def load_spec(path: str | Path) -> Spec:
    try:
        data = json.loads(Path(path).read_text(encoding="utf-8"))
    except UnicodeDecodeError:
        raise SpecError(f"Файл правил {path}: кодування не UTF-8.") from None
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
    if not results:
        raise SpecError("Правила не дали жодної перевірки: нуль перевірок це не «усе зійшлось».")

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
