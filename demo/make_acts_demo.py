"""Вигаданий еталонний акт, чистий акт і акт із вісьмома підкладеними помилками.

Усі позиції, обсяги й ціни вигадані. Стовпчики умовні: справжнього вивантаження
з кошторисної програми тут немає, тому розкладка задається у файлі правил.
Запуск: python -m demo.make_acts_demo
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

from openpyxl import Workbook

HERE = Path(__file__).parent
REFERENCE = HERE / "acts_reference.xlsx"
CLEAN = HERE / "acts_clean.xlsx"
WITH_DEFECTS = HERE / "acts_with_defects.xlsx"
SPEC = HERE / "acts.json"
SAMPLE_REPORT = HERE / "report_acts.md"
SHEET = "Акт"


@dataclass
class Row:
    key: str
    name: str
    unit: str
    qty: Decimal | str
    price: Decimal
    amount: Decimal

    @classmethod
    def of(cls, key: str, name: str, unit: str, qty: str, price: str) -> Row:
        q, p = Decimal(qty), Decimal(price)
        return cls(key, name, unit, q, p, (q * p).quantize(Decimal("0.01"), ROUND_HALF_UP))


SECTIONS = {
    "1": "Розділ 1. Земляні роботи",
    "4": "Розділ 2. Фундаменти",
    "8": "Розділ 3. Стіни і фасад",
    "12": "Розділ 4. Інше",
}


def positions() -> list[Row]:
    return [
        Row.of("1", "Розробка ґрунту екскаватором з навантаженням", "м3", "412.5", "86.37"),
        Row.of("2", "Ущільнення ґрунту віброкатком", "м3", "380", "22.15"),
        Row.of("3", "Влаштування піщаної основи", "м3", "64.2", "512.80"),
        Row.of("4", "Влаштування бетонної підготовки", "м3", "38.6", "3940.00"),
        Row.of("5", "Армування фундаментної плити", "т", "12.84", "38250.00"),
        Row.of("6", "Бетонування фундаментної плити", "м3", "214", "4615.50"),
        Row.of("7", "Гідроізоляція обмазувальна", "м2", "520", "188.30"),
        Row.of("8", "Мурування стін з газоблоку", "м3", "96.4", "3210.00"),
        Row.of("9", "Монтаж залізобетонних перемичок", "шт", "42", "865.00"),
        Row.of("10", "Утеплення фасаду мінеральною ватою", "м2", "310", "742.60"),
        Row.of("11", "Монтаж віконних блоків", "м2", "48.6", "6120.00"),
        Row.of("12", "Вивезення ґрунту", "т", "650", "74.00"),
    ]


def _find(rows: list[Row], key: str) -> int:
    return next(i for i, row in enumerate(rows) if row.key == key)


def _neighbour_volume(rows: list[Row]) -> None:
    row = rows[_find(rows, "3")]
    row.qty = rows[_find(rows, "2")].qty
    row.amount = (row.qty * row.price).quantize(Decimal("0.01"), ROUND_HALF_UP)


def _dropped(rows: list[Row]) -> None:
    del rows[_find(rows, "9")]


def _retry_duplicate(rows: list[Row]) -> None:
    i = _find(rows, "7")
    rows.insert(i + 1, Row(**vars(rows[i])))


def _old_price(rows: list[Row]) -> None:
    row = rows[_find(rows, "10")]
    row.price = Decimal("724.60")
    row.amount = (row.qty * row.price).quantize(Decimal("0.01"), ROUND_HALF_UP)


def _wrong_unit(rows: list[Row]) -> None:
    rows[_find(rows, "8")].unit = "м2"


def _qty_as_text(rows: list[Row]) -> None:
    rows[_find(rows, "6")].qty = "214,0"


def _truncated_kopeck(rows: list[Row]) -> None:
    row = rows[_find(rows, "1")]
    assert isinstance(row.qty, Decimal)
    row.amount = (row.qty * row.price).quantize(Decimal("0.01"), rounding="ROUND_DOWN")


def _extra_position(rows: list[Row]) -> None:
    rows.append(Row.of("13", "Прибирання будівельного сміття", "т", "18", "410.00"))


@dataclass(frozen=True)
class Defect:
    id: str
    title: str
    apply: Callable[[list[Row]], None]
    expect: dict[str, str]


TOTAL = {"total:act_vs_reference": "DIFF"}
DEFECTS = [
    Defect(
        "A1",
        "Обсяг узято з сусідньої позиції",
        _neighbour_volume,
        {"3:qty": "DIFF", "3:amount": "DIFF", **TOTAL},
    ),
    Defect("A2", "Позицію пропущено", _dropped, {"9:presence": "DIFF", **TOTAL}),
    Defect(
        "A3",
        "Рядок задвоївся після повторної спроби",
        _retry_duplicate,
        {"7:presence": "SKIPPED", **TOTAL},
    ),
    Defect(
        "A4",
        "Ціна зі старої редакції, переставлені цифри",
        _old_price,
        {"10:price": "DIFF", "10:amount": "DIFF", **TOTAL},
    ),
    Defect("A5", "Одиниця виміру не та, суми ті самі", _wrong_unit, {"8:unit": "DIFF"}),
    Defect(
        "A6", "Кількість набрана текстом", _qty_as_text, {"6:qty": "SKIPPED", "6:calc": "SKIPPED"}
    ),
    Defect(
        "A7",
        "Копійку відкинуто замість округлення",
        _truncated_kopeck,
        {"1:amount": "DIFF", "1:calc": "DIFF", **TOTAL},
    ),
    Defect(
        "A8",
        "Зайва позиція, якої немає в еталоні",
        _extra_position,
        {"13:presence": "DIFF", **TOTAL},
    ),
]


def _cell(value: Decimal | str) -> float | str:
    return float(value) if isinstance(value, Decimal) else value


def _write(rows: list[Row], path: Path, *, reference: bool) -> None:
    """Еталон і акт мають різну розкладку стовпчиків, як файли з двох різних рук."""
    book = Workbook()
    sheet = book.active
    assert sheet is not None
    sheet.title = SHEET
    if reference:
        sheet["A1"] = "Еталонний акт (вигадані дані, умовний формат)"
        head = [
            "№ з/п",
            "Позиція кошторису",
            "Найменування робіт",
            "Од. вим.",
            "Кількість",
            "Ціна за од., грн",
            "Вартість, грн",
        ]
    else:
        sheet["A1"] = "Акт, підготовлений агентом (вигадані дані, умовний формат)"
        head = [
            "Позиція кошторису",
            "Найменування робіт",
            "Од. вим.",
            "Кількість",
            "Ціна за од., грн",
            "Вартість, грн",
        ]
    for col, text in enumerate(head, start=1):
        sheet.cell(row=3, column=col, value=text)
    at = 4
    for number, row in enumerate(rows, start=1):
        if reference and row.key in SECTIONS:
            sheet.cell(row=at, column=3, value=SECTIONS[row.key])
            at += 1
        values: list[object] = [
            row.key,
            row.name,
            row.unit,
            _cell(row.qty),
            float(row.price),
            float(row.amount),
        ]
        if reference:
            values = [number, *values]
        for col, value in enumerate(values, start=1):
            sheet.cell(row=at, column=col, value=value)
        at += 1
    total = sum((row.amount for row in rows), Decimal(0))
    label_col, amount_col = (3, 7) if reference else (2, 6)
    sheet.cell(row=at, column=label_col, value="Разом")
    sheet.cell(row=at, column=amount_col, value=float(total))
    book.save(path)


def act_rows(ids: list[str] | None = None, *, reorder: bool = True) -> list[Row]:
    rows = positions()
    if reorder:  # агент видав позиції 2 і 3 в іншому порядку: звірка за номером це витримує
        i = _find(rows, "2")
        rows[i], rows[i + 1] = rows[i + 1], rows[i]
    for defect in DEFECTS:
        if ids is None or defect.id in ids:
            defect.apply(rows)
    return rows


def build_reference(path: Path = REFERENCE) -> None:
    _write(positions(), path, reference=True)


def build_act(ids: list[str] | None, path: Path) -> None:
    _write(act_rows(ids), path, reference=False)


def main() -> None:
    from recheck import acts
    from recheck.book import Book

    build_reference()
    build_act([], CLEAN)
    build_act(None, WITH_DEFECTS)
    outcome = acts.run(Book(WITH_DEFECTS), Book(REFERENCE), acts.load_spec(SPEC))
    text = acts.render(outcome, WITH_DEFECTS.name, REFERENCE.name, SPEC.name)
    SAMPLE_REPORT.write_text(text, encoding="utf-8", newline="\n")
    print(f"Готово: {REFERENCE.name}, {CLEAN.name}, {WITH_DEFECTS.name}, {SAMPLE_REPORT.name}")


if __name__ == "__main__":
    main()
