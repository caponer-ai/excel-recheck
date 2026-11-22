"""Порядкова звірка двох документів: акт проти еталонного акта.

Рядки зіставляються за номером позиції, а не за місцем у таблиці. Для кожної
позиції окремо звіряються одиниця, кількість, ціна і вартість, а в акті ще й
арифметика рядка (кількість × ціна). Позиція, якої немає в одному з документів,
дубль номера позиції і текст замість числа ніколи не стають «зійшлось».

Запуск: python -m recheck.acts АКТ.xlsx --reference ЕТАЛОН.xlsx --spec ПРАВИЛА.json
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import time
from collections.abc import Sequence
from dataclasses import dataclass, field, replace
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from pathlib import Path
from zipfile import BadZipFile

from openpyxl.utils.exceptions import InvalidFileException

from recheck.book import Book, Cell, Kind, RefError
from recheck.checks import SpecError, Status
from recheck.report import money

FIELDS = ("unit", "qty", "price", "amount")
NUMERIC = ("qty", "price", "amount")
FIELD_NAMES = {
    "presence": "наявність позиції",
    "unit": "одиниця виміру",
    "qty": "кількість",
    "price": "ціна за одиницю",
    "amount": "вартість",
    "calc": "вартість = кількість × ціна",
}
_COL = re.compile(r"[A-Z]{1,3}")
_SUPERSCRIPT = str.maketrans({"²": "2", "³": "3"})
_CENT = Decimal("0.01")


@dataclass(frozen=True)
class Layout:
    """Де в документі лежать стовпчики і де закінчується таблиця."""

    sheet: str
    first_row: int
    key: str
    unit: str
    qty: str
    price: str
    amount: str
    total_label_col: str
    total_label: str
    name: str = ""


@dataclass(frozen=True)
class Spec:
    act: Layout
    reference: Layout
    tol: Decimal = Decimal(0)


@dataclass(frozen=True)
class Line:
    """Рядок документа з номером позиції."""

    row: int
    key: str
    name: str
    cells: dict[str, Cell]


@dataclass(frozen=True)
class Stray:
    """Рядок із вмістом, який не є ні позицією, ні підсумком."""

    doc: str
    row: int
    label: str
    reason: str


@dataclass(frozen=True)
class Doc:
    title: str
    layout: Layout
    lines: list[Line]
    total: Cell | None
    strays: list[Stray]


@dataclass(frozen=True)
class Compared:
    id: str
    key: str
    field: str
    status: Status
    act: Decimal | str | None
    reference: Decimal | str | None
    reason: str = ""
    refs: tuple[str, ...] = ()

    @property
    def diff(self) -> Decimal | None:
        if isinstance(self.act, Decimal) and isinstance(self.reference, Decimal):
            return self.act - self.reference
        return None


@dataclass
class Outcome:
    results: list[Compared] = field(default_factory=list)
    strays: list[Stray] = field(default_factory=list)
    positions_act: int = 0
    positions_reference: int = 0


def _column(raw: object, where: str) -> str:
    if not isinstance(raw, str) or _COL.fullmatch(raw) is None:
        raise SpecError(f"{where}: стовпчик має бути літерою, як «C», а стоїть {raw!r}.")
    return raw


def _layout(raw: object, where: str) -> Layout:
    if not isinstance(raw, dict):
        raise SpecError(f"Розділ «{where}» має бути об'єктом з описом таблиці.")
    columns = raw.get("columns")
    if not isinstance(columns, dict):
        raise SpecError(f"{where}: бракує «columns» з літерами стовпчиків.")
    missing = [name for name in ("key", *FIELDS) if name not in columns]
    if missing:
        raise SpecError(f"{where}: у «columns» бракує {', '.join(missing)}.")
    first_row = raw.get("first_row")
    if not isinstance(first_row, int) or isinstance(first_row, bool) or first_row < 1:
        raise SpecError(f"{where}: «first_row» має бути номером рядка, від 1.")
    sheet, label = raw.get("sheet"), raw.get("total_label")
    if not isinstance(sheet, str) or not sheet:
        raise SpecError(f"{where}: бракує назви аркуша «sheet».")
    if not isinstance(label, str) or not label.strip():
        raise SpecError(f"{where}: бракує «total_label», підпису підсумкового рядка.")
    used = {name: _column(columns[name], f"{where}.columns.{name}") for name in ("key", *FIELDS)}
    if len(set(used.values())) != len(used):
        raise SpecError(f"{where}: два поля вказують на один стовпчик: {used}.")
    name_col = columns.get("name", "")
    return Layout(
        sheet=sheet,
        first_row=first_row,
        key=used["key"],
        unit=used["unit"],
        qty=used["qty"],
        price=used["price"],
        amount=used["amount"],
        total_label_col=_column(raw.get("total_label_col"), f"{where}.total_label_col"),
        total_label=label.strip(),
        name=_column(name_col, f"{where}.columns.name") if name_col else "",
    )


def load_spec(path: str | Path) -> Spec:
    try:
        raw = json.loads(Path(path).read_text(encoding="utf-8"))
    except json.JSONDecodeError as error:
        raise SpecError(f"{Path(path).name} не читається як JSON ({error}).") from None
    if not isinstance(raw, dict):
        raise SpecError("Файл правил має бути об'єктом з розділами «act» і «reference».")
    try:
        tol = Decimal(str(raw.get("tolerance", "0")))
    except InvalidOperation:
        raise SpecError(f"Допуск {raw.get('tolerance')!r} не є числом.") from None
    if not tol.is_finite() or tol < 0:
        raise SpecError(f"Допуск має бути скінченним і не від'ємним, а стоїть {tol}.")
    return Spec(_layout(raw.get("act"), "act"), _layout(raw.get("reference"), "reference"), tol)


def _key(cell: Cell) -> str:
    """Номер позиції як текст: 12, 12.0 і «12» це одна позиція."""
    if cell.kind in (Kind.NUMBER, Kind.TEXT_NUMBER) and cell.value is not None:
        value = cell.value.normalize()
        return format(value, "f") if value != value.to_integral() else str(int(value))
    if cell.kind is Kind.EMPTY:
        return ""
    return str(cell.raw).strip()


def _label(book: Book, sheet: str, addr: str) -> str:
    raw = book.cell(sheet, addr).raw
    return "" if raw is None else str(raw).strip()


def _tagged(cell: Cell, title: str) -> Cell:
    """Обидва файли можуть мати аркуш з тією самою назвою: адреса каже, з якого документа."""
    return replace(cell, sheet=f"{title}: {cell.sheet}")


def read_doc(book: Book, layout: Layout, title: str) -> Doc:
    lines: list[Line] = []
    strays: list[Stray] = []
    total: Cell | None = None
    data_cols = (layout.key, layout.unit, layout.qty, layout.price, layout.amount)
    for row in range(layout.first_row, book.max_row(layout.sheet) + 1):
        label = _label(book, layout.sheet, f"{layout.total_label_col}{row}")
        filled = [
            c for c in data_cols if book.cell(layout.sheet, f"{c}{row}").kind is not Kind.EMPTY
        ]
        if total is not None:
            if filled or label:
                strays.append(Stray(title, row, label, "стоїть після підсумкового рядка"))
            continue
        if label.casefold() == layout.total_label.casefold():
            total = _tagged(book.cell(layout.sheet, f"{layout.amount}{row}"), title)
            continue
        if not filled:
            continue
        key = _key(book.cell(layout.sheet, f"{layout.key}{row}"))
        if not key:
            strays.append(Stray(title, row, label, "є числа, але немає номера позиції"))
            continue
        cells = {
            name: _tagged(book.cell(layout.sheet, f"{getattr(layout, name)}{row}"), title)
            for name in FIELDS
        }
        name = _label(book, layout.sheet, f"{layout.name}{row}") if layout.name else ""
        lines.append(Line(row, key, name, cells))
    return Doc(title, layout, lines, total, strays)


def _unit(cell: Cell) -> str:
    return (
        ""
        if cell.raw is None
        else re.sub(r"[\s.]", "", str(cell.raw)).translate(_SUPERSCRIPT).casefold()
    )


def _why_not_number(cell: Cell) -> str:
    return f"{cell.ref}: {cell.kind.value}" + (f" («{cell.raw}»)" if cell.raw is not None else "")


def _compare_number(ident: str, key: str, name: str, a: Cell, r: Cell, tol: Decimal) -> Compared:
    refs = (a.ref, r.ref)
    bad = [c for c in (a, r) if c.kind is not Kind.NUMBER]
    if bad:
        reason = "не число: " + "; ".join(_why_not_number(c) for c in bad)
        return Compared(ident, key, name, Status.SKIPPED, a.value, r.value, reason, refs)
    assert a.value is not None and r.value is not None
    status = Status.OK if abs(a.value - r.value) <= tol else Status.DIFF
    return Compared(ident, key, name, status, a.value, r.value, "", refs)


def _compare_line(a: Line, r: Line, tol: Decimal) -> list[Compared]:
    out: list[Compared] = []
    key = a.key
    ua, ur = a.cells["unit"], r.cells["unit"]
    refs = (ua.ref, ur.ref)
    if not _unit(ua) or not _unit(ur):
        empty = ", ".join(c.ref for c in (ua, ur) if not _unit(c))
        out.append(
            Compared(
                f"{key}:unit",
                key,
                "unit",
                Status.SKIPPED,
                None,
                None,
                f"одиниця не вказана: {empty}",
                refs,
            )
        )
    else:
        status = Status.OK if _unit(ua) == _unit(ur) else Status.DIFF
        out.append(Compared(f"{key}:unit", key, "unit", status, str(ua.raw), str(ur.raw), "", refs))
    for name in NUMERIC:
        out.append(_compare_number(f"{key}:{name}", key, name, a.cells[name], r.cells[name], tol))
    qty, price, amount = a.cells["qty"], a.cells["price"], a.cells["amount"]
    refs = (amount.ref, qty.ref, price.ref)
    bad = [c for c in (qty, price, amount) if c.kind is not Kind.NUMBER]
    if bad:
        reason = "не число: " + "; ".join(_why_not_number(c) for c in bad)
        out.append(
            Compared(f"{key}:calc", key, "calc", Status.SKIPPED, amount.value, None, reason, refs)
        )
    else:
        assert qty.value is not None and price.value is not None and amount.value is not None
        mine = (qty.value * price.value).quantize(_CENT, ROUND_HALF_UP)
        status = Status.OK if abs(amount.value - mine) <= tol else Status.DIFF
        out.append(Compared(f"{key}:calc", key, "calc", status, amount.value, mine, "", refs))
    return out


def _by_key(doc: Doc) -> tuple[dict[str, Line], dict[str, list[int]]]:
    seen: dict[str, list[Line]] = {}
    for line in doc.lines:
        seen.setdefault(line.key, []).append(line)
    unique = {key: rows[0] for key, rows in seen.items() if len(rows) == 1}
    dupes = {key: [line.row for line in rows] for key, rows in seen.items() if len(rows) > 1}
    return unique, dupes


def _sum_rows(doc: Doc) -> tuple[Decimal | None, str]:
    bad = [
        line.cells["amount"] for line in doc.lines if line.cells["amount"].kind is not Kind.NUMBER
    ]
    if bad:
        return None, "вартість не число: " + "; ".join(_why_not_number(c) for c in bad)
    return sum((line.cells["amount"].value or Decimal(0) for line in doc.lines), Decimal(0)), ""


def _total_check(
    ident: str, doc: Doc, other: Decimal | None, why: str, tol: Decimal, refs: tuple[str, ...]
) -> Compared:
    total = doc.total
    if total is None:
        return Compared(
            ident,
            "Разом",
            "total",
            Status.SKIPPED,
            None,
            other,
            f"у документі «{doc.title}» немає рядка «{doc.layout.total_label}»",
            refs,
        )
    if total.kind is not Kind.NUMBER:
        return Compared(
            ident,
            "Разом",
            "total",
            Status.SKIPPED,
            None,
            other,
            "підсумок не число: " + _why_not_number(total),
            (total.ref, *refs),
        )
    if other is None:
        return Compared(
            ident, "Разом", "total", Status.SKIPPED, total.value, None, why, (total.ref, *refs)
        )
    assert total.value is not None
    status = Status.OK if abs(total.value - other) <= tol else Status.DIFF
    return Compared(ident, "Разом", "total", status, total.value, other, "", (total.ref, *refs))


def compare(act: Doc, reference: Doc, tol: Decimal = Decimal(0)) -> Outcome:
    outcome = Outcome(strays=[*act.strays, *reference.strays])
    act_lines, act_dupes = _by_key(act)
    ref_lines, ref_dupes = _by_key(reference)
    outcome.positions_act = len(act_lines) + len(act_dupes)
    outcome.positions_reference = len(ref_lines) + len(ref_dupes)

    order = [line.key for line in reference.lines] + [line.key for line in act.lines]
    keys = list(dict.fromkeys(order))
    for key in keys:
        if key in act_dupes or key in ref_dupes:
            parts = []
            if key in act_dupes:
                parts.append(f"в акті рядки {', '.join(map(str, act_dupes[key]))}")
            if key in ref_dupes:
                parts.append(f"в еталоні рядки {', '.join(map(str, ref_dupes[key]))}")
            reason = "номер позиції повторюється: " + "; ".join(parts)
            outcome.results.append(
                Compared(f"{key}:presence", key, "presence", Status.SKIPPED, None, None, reason)
            )
            continue
        a, r = act_lines.get(key), ref_lines.get(key)
        if a is None and r is not None:
            outcome.results.append(
                Compared(
                    f"{key}:presence",
                    key,
                    "presence",
                    Status.DIFF,
                    "немає",
                    "є",
                    f"позиції немає в акті (еталон, рядок {r.row})",
                    (r.cells["amount"].ref,),
                )
            )
            continue
        if r is None and a is not None:
            outcome.results.append(
                Compared(
                    f"{key}:presence",
                    key,
                    "presence",
                    Status.DIFF,
                    "є",
                    "немає",
                    f"зайва позиція в акті (рядок {a.row})",
                    (a.cells["amount"].ref,),
                )
            )
            continue
        assert a is not None and r is not None
        outcome.results.extend(_compare_line(a, r, tol))

    act_sum, act_why = _sum_rows(act)
    ref_sum, ref_why = _sum_rows(reference)
    outcome.results.append(_total_check("total:act_rows", act, act_sum, act_why, tol, ()))
    outcome.results.append(
        _total_check("total:reference_rows", reference, ref_sum, ref_why, tol, ())
    )
    other = (
        reference.total.value if reference.total and reference.total.kind is Kind.NUMBER else None
    )
    why = "підсумок еталона не прочитано" if other is None else ""
    ref_refs = (reference.total.ref,) if reference.total else ()
    outcome.results.append(_total_check("total:act_vs_reference", act, other, why, tol, ref_refs))
    return outcome


def run(act_book: Book, reference_book: Book, spec: Spec) -> Outcome:
    act = read_doc(act_book, spec.act, "акт")
    reference = read_doc(reference_book, spec.reference, "еталон")
    return compare(act, reference, spec.tol)


def exit_code(outcome: Outcome) -> int:
    statuses = {r.status for r in outcome.results}
    if Status.DIFF in statuses:
        return 1
    if Status.SKIPPED in statuses or outcome.strays:
        return 2
    return 0


def _tally(outcome: Outcome) -> dict[Status, int]:
    return {s: sum(1 for r in outcome.results if r.status is s) for s in Status}


def summary(outcome: Outcome) -> str:
    t = _tally(outcome)
    return (
        f"Позицій в акті: {outcome.positions_act}, в еталоні: {outcome.positions_reference}. "
        f"Перевірок: {len(outcome.results)}. Зійшлось: {t[Status.OK]}. "
        f"Розбіжностей: {t[Status.DIFF]}. Не перевірено: {t[Status.SKIPPED]}. "
        f"Рядків поза таблицею: {len(outcome.strays)}."
    )


def _show(value: Decimal | str | None) -> str:
    if isinstance(value, Decimal):
        return money(value)
    return "немає" if value is None else str(value)


def _what(result: Compared) -> str:
    if result.field == "total":
        return {
            "total:act_rows": "«Разом» акта проти суми його рядків",
            "total:reference_rows": "«Разом» еталона проти суми його рядків",
            "total:act_vs_reference": "«Разом» акта проти «Разом» еталона",
        }[result.id]
    return FIELD_NAMES[result.field]


def render(outcome: Outcome, act_name: str, reference_name: str, spec_name: str) -> str:
    """Markdown-звіт. Часу прогону тут немає, щоб звіт був відтворюваним."""
    lines = [
        f"# Порядкова звірка: {act_name} проти {reference_name}",
        "",
        f"Правила: `{spec_name}`. Рядки зіставлено за номером позиції, не за місцем у таблиці.",
        "",
        summary(outcome),
        "",
        "## Розбіжності",
        "",
    ]
    diffs = [r for r in outcome.results if r.status is Status.DIFF]
    if diffs:
        lines += [
            "| Позиція | Що звіряли | В акті | В еталоні | Різниця | Клітинки |",
            "|---|---|---:|---:|---:|---|",
        ]
        for r in diffs:
            what = r.reason if r.field == "presence" else _what(r)
            cells = ", ".join(f"`{ref}`" for ref in r.refs)
            reference = _show(r.reference)
            if r.field == "calc":
                reference = f"{reference} (перерахунок)"
            lines.append(
                f"| {r.key} | {what} | {_show(r.act)} | {reference} "
                f"| {_show(r.diff) if r.diff is not None else ''} | {cells} |"
            )
    else:
        lines.append("Немає.")
    lines += ["", "## Що не перевірено і чому", ""]
    skipped = [r for r in outcome.results if r.status is Status.SKIPPED]
    if skipped:
        lines += [f"- {r.key}, {_what(r)}: {r.reason}" for r in skipped]
    else:
        lines.append("Немає.")
    lines += ["", "### Рядки, які не є ні позицією, ні підсумком", ""]
    if outcome.strays:
        lines += [f"- {s.doc}, р. {s.row}: «{s.label}», {s.reason}" for s in outcome.strays]
    else:
        lines.append("Немає.")
    lines += ["", "## Що зійшлось", ""]
    by_key: dict[str, list[Status]] = {}
    for r in outcome.results:
        if r.field != "total":
            by_key.setdefault(r.key, []).append(r.status)
    clean = [key for key, got in by_key.items() if set(got) == {Status.OK}]
    lines.append(
        f"Позицій, де зійшлось усе (одиниця, кількість, ціна, вартість, арифметика): "
        f"{len(clean)} з {len(by_key)}."
    )
    totals = [r for r in outcome.results if r.field == "total" and r.status is Status.OK]
    lines += [f"- {_what(r)}: {_show(r.act)}" for r in totals]
    lines += ["", "## Звідки числа", ""]
    for r in outcome.results:
        cells = ", ".join(f"`{ref}`" for ref in r.refs) or "немає клітинок"
        lines.append(f"- [{r.status.value}] {r.key}, {_what(r)}: {cells}")
    lines.append("")
    return "\n".join(lines)


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="recheck.acts",
        description="Порядкова звірка акта з еталонним актом за номерами позицій.",
    )
    parser.add_argument("act", type=Path, help="акт, який перевіряємо (.xlsx)")
    parser.add_argument("--reference", type=Path, required=True, help="еталонний акт (.xlsx)")
    parser.add_argument("--spec", type=Path, required=True, help="опис стовпчиків .json")
    parser.add_argument("--out", type=Path, help="куди записати звіт у Markdown")
    args = parser.parse_args(argv)

    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")

    started = time.perf_counter()
    try:
        spec = load_spec(args.spec)
        books = []
        for path in (args.act, args.reference):
            try:
                books.append(Book(path))
            except (InvalidFileException, BadZipFile, KeyError) as error:
                raise SpecError(f"{path.name} не читається як .xlsx ({error})") from None
        outcome = run(books[0], books[1], spec)
    except (SpecError, RefError, OSError) as error:
        print(f"Не можу почати звірку: {error}", file=sys.stderr)
        return 3
    elapsed = time.perf_counter() - started

    print(summary(outcome))
    for r in outcome.results:
        if r.status is Status.DIFF and r.field == "presence":
            print(f"РОЗБІЖНІСТЬ  позиція {r.key}: {r.reason}")
        elif r.status is Status.DIFF:
            other = "перерахунок" if r.field == "calc" else "в еталоні"
            where = "підсумок" if r.field == "total" else f"позиція {r.key}"
            print(
                f"РОЗБІЖНІСТЬ  {where}, {_what(r)}: "
                f"в акті {_show(r.act)}, {other} {_show(r.reference)}"
            )
        elif r.status is Status.SKIPPED:
            print(f"НЕ ПЕРЕВІРЕНО  позиція {r.key}, {_what(r)}: {r.reason}")
    for s in outcome.strays:
        print(f"ПОЗА ТАБЛИЦЕЮ  {s.doc}, р. {s.row}: «{s.label}», {s.reason}")
    print(f"Час прогону: {elapsed:.2f} с")

    if args.out:
        text = render(outcome, args.act.name, args.reference.name, args.spec.name)
        try:
            args.out.write_text(text, encoding="utf-8", newline="\n")
        except OSError as error:
            print(f"Звірку виконано, але звіт не записано: {error}", file=sys.stderr)
            return 3
        print(f"Звіт: {args.out}")
    return exit_code(outcome)


if __name__ == "__main__":
    raise SystemExit(main())
