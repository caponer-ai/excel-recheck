"""Вигаданий квартальний звіт і його копія з п'ятьма підкладеними помилками.

Усі назви і суми згенеровано, жодного реального документа тут немає.
Запуск: python -m demo.make_demo
"""

from __future__ import annotations

import random
from collections.abc import Callable
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

from openpyxl import Workbook, load_workbook

HERE = Path(__file__).parent
CLEAN = HERE / "clean.xlsx"
WITH_DEFECTS = HERE / "with_defects.xlsx"
SPEC = HERE / "checks.json"
SAMPLE_REPORT = HERE / "report_with_defects.md"

INCOME = "Доходи"
EXPENSE = "Витрати"
SUMMARY = "Зведення"
STOCK = "Оборот складу"

INCOME_ITEMS = [
    "Медичні послуги за договором",
    "Платні послуги",
    "Страхові компанії",
    "Оренда приміщень",
    "Благодійна допомога цільова",
    "Благодійна допомога нецільова",
    "Інші надходження",
]
EXPENSE_ITEMS = [
    "Оплата праці",
    "Нарахування на оплату праці",
    "Медикаменти",
    "Паливо",
    "Комунальні послуги",
    "Амортизація",
    "Інші витрати",
]
STOCK_GROUPS = {
    "201": (
        "Медикаменти",
        ["Розчин А", "Розчин Б", "Таблетки В", "Ампули Г", "Мазь Д", "Спирт"],
    ),
    "203": ("Паливо", ["Бензин А-95", "Дизельне пальне", "Газ"]),
    "207": ("Запчастини", ["Шини", "Фільтри", "Акумулятори", "Гальмівні колодки"]),
}
FIRST_ITEM_ROW = 4
TOTAL_ROW = FIRST_ITEM_ROW + len(INCOME_ITEMS)  # 11
SOURCES_ROW = 10
STOCK_FIRST_ROW = 4


def _hryvnias(kopecks: int) -> float:
    return kopecks / 100


def _thousands(kopecks: int) -> Decimal:
    return (Decimal(kopecks) / Decimal(100_000)).quantize(Decimal("0.1"), ROUND_HALF_UP)


def build_clean(path: Path = CLEAN) -> None:
    rng = random.Random(2026)
    book = Workbook()

    totals: dict[str, int] = {}
    income_now: list[int] = []
    for title, items in ((INCOME, INCOME_ITEMS), (EXPENSE, EXPENSE_ITEMS)):
        sheet = book.active if title == INCOME else book.create_sheet()
        sheet.title = title
        sheet["A1"] = f"{title} закладу «Приклад», перший квартал (вигадані дані)"
        sheet["A3"], sheet["B3"], sheet["C3"] = (
            "Стаття",
            "Минулий рік, грн",
            "Звітний квартал, грн",
        )
        past_total = now_total = 0
        for offset, item in enumerate(items):
            row = FIRST_ITEM_ROW + offset
            past = rng.randrange(20_000_000, 900_000_000)
            now = rng.randrange(20_000_000, 900_000_000)
            sheet[f"A{row}"], sheet[f"B{row}"], sheet[f"C{row}"] = (
                item,
                _hryvnias(past),
                _hryvnias(now),
            )
            past_total += past
            now_total += now
            if title == INCOME:
                income_now.append(now)
        sheet[f"A{TOTAL_ROW}"] = f"{title}, всього"
        sheet[f"B{TOTAL_ROW}"] = _hryvnias(past_total)
        sheet[f"C{TOTAL_ROW}"] = _hryvnias(now_total)
        totals[title] = now_total

    summary = book.create_sheet(SUMMARY)
    summary["A1"] = "Зведення, тис. грн (вигадані дані)"
    summary["A3"], summary["B3"] = "Показник", "тис. грн"
    income_k, expense_k = _thousands(totals[INCOME]), _thousands(totals[EXPENSE])
    summary["A4"], summary["B4"] = "Дохід, всього", float(income_k)
    summary["A5"], summary["B5"] = "Витрати, всього", float(expense_k)
    summary["A6"], summary["B6"] = "Результат", float(income_k - expense_k)
    summary["A8"] = "Надходження за джерелами, тис. грн"
    summary["A9"], summary["B9"] = "Показник", "Всього"
    summary[f"A{SOURCES_ROW}"] = "Надходження"
    parts = [_thousands(value) for value in income_now]
    for index, (item, part) in enumerate(zip(INCOME_ITEMS, parts, strict=True)):
        column = "CDEFGHI"[index]
        summary[f"{column}9"] = item
        summary[f"{column}{SOURCES_ROW}"] = float(part)
    summary[f"B{SOURCES_ROW}"] = float(sum(parts, Decimal(0)))

    stock = book.create_sheet(STOCK)
    stock["A1"] = "Оборотна відомість складу, грн (вигадані дані)"
    for column, head in zip(
        "ABCDEFG",
        [
            "Код",
            "Назва",
            "Одиниця",
            "Залишок на початок",
            "Прихід",
            "Розхід",
            "Залишок на кінець",
        ],
        strict=True,
    ):
        stock[f"{column}3"] = head
    row = STOCK_FIRST_ROW
    grand = [0, 0, 0, 0]
    for code, (name, items) in STOCK_GROUPS.items():
        group_row = row
        stock[f"A{row}"], stock[f"B{row}"] = code, name
        row += 1
        subtotal = [0, 0, 0, 0]
        for number, item in enumerate(items, start=1):
            opening = rng.randrange(500_000, 60_000_000)
            incoming = rng.randrange(0, 40_000_000)
            outgoing = rng.randrange(0, opening + incoming)
            values = [opening, incoming, outgoing, opening + incoming - outgoing]
            stock[f"A{row}"], stock[f"B{row}"] = f"{code}.{number:03d}", item
            stock[f"C{row}"] = rng.choice(["шт", "л", "упак"])
            for column, value in zip("DEFG", values, strict=True):
                stock[f"{column}{row}"] = _hryvnias(value)
            subtotal = [a + b for a, b in zip(subtotal, values, strict=True)]
            row += 1
        for column, value in zip("DEFG", subtotal, strict=True):
            stock[f"{column}{group_row}"] = _hryvnias(value)
        grand = [a + b for a, b in zip(grand, subtotal, strict=True)]
    stock[f"A{row}"] = "Разом"
    for column, value in zip("DEFG", grand, strict=True):
        stock[f"{column}{row}"] = _hryvnias(value)

    book.save(path)


@dataclass(frozen=True)
class Defect:
    id: str
    title: str
    apply: Callable[[Workbook], None]
    expect: dict[str, str]  # id перевірки -> назва статусу, який вона має отримати


def _lost_part(book: Workbook) -> None:
    sheet = book[SUMMARY]
    sheet[f"B{SOURCES_ROW}"] = round(
        sheet[f"B{SOURCES_ROW}"].value - sheet[f"H{SOURCES_ROW}"].value, 1
    )


def _number_as_text(book: Workbook) -> None:
    cell = book[EXPENSE]["C6"]
    cell.value = f"{cell.value:,.2f}".replace(",", " ").replace(".", ",")


def _swapped_digits(book: Workbook) -> None:
    cell = book[STOCK]["G6"]
    whole, _, fraction = f"{cell.value:.2f}".partition(".")
    for index in range(len(whole) - 1):
        if whole[index] != whole[index + 1]:
            swapped = whole[:index] + whole[index + 1] + whole[index] + whole[index + 2 :]
            cell.value = float(f"{swapped}.{fraction}")
            return
    raise AssertionError("у числі немає двох різних сусідніх цифр")


def _one_kopeck(book: Workbook) -> None:
    cell = book[INCOME][f"C{TOTAL_ROW}"]
    cell.value = round(cell.value + 0.01, 2)


def _double_count(book: Workbook) -> None:
    sheet = book[STOCK]
    total_row = next(
        row for row in range(1, sheet.max_row + 1) if sheet[f"A{row}"].value == "Разом"
    )
    sheet[f"E{total_row}"] = round(sheet[f"E{total_row}"].value + sheet["E4"].value, 2)


DEFECTS = [
    Defect(
        "D1",
        "Підсумок губить один доданок: «Всього» у зведенні не враховує нецільову благодійну",
        _lost_part,
        {"sum.sources_total": "DIFF", "sum.sources_vs_income": "DIFF"},
    ),
    Defect(
        "D2",
        "Число записане текстом: сума статті у витратах набрана як «1 234,56»",
        _number_as_text,
        {"exp.total": "SKIPPED"},
    ),
    Defect(
        "D3",
        "Переставлені сусідні цифри в залишку на кінець одного рядка складу",
        _swapped_digits,
        {
            "stock.rule.row.6": "DIFF",
            "stock.group.201.G": "DIFF",
            "stock.total.G": "DIFF",
        },
    ),
    Defect(
        "D4",
        "Одна копійка: підсумок доходів більший за суму статей на 0,01",
        _one_kopeck,
        {"inc.total": "DIFF"},
    ),
    Defect(
        "D5",
        "Подвійний рахунок: «Разом» по приходу містить ще й підсумок групи",
        _double_count,
        {"stock.total.E": "DIFF", "stock.rule.total": "DIFF"},
    ),
]


def build_with_defects(
    ids: list[str] | None = None, source: Path = CLEAN, target: Path = WITH_DEFECTS
) -> None:
    book = load_workbook(source)
    for defect in DEFECTS:
        if ids is None or defect.id in ids:
            defect.apply(book)
    book.save(target)


def main() -> None:
    from recheck.book import Book
    from recheck.checks import load_spec, run
    from recheck.report import render

    build_clean()
    build_with_defects()
    outcome = run(Book(WITH_DEFECTS), load_spec(SPEC))
    text = render(outcome, WITH_DEFECTS.name, SPEC.name)
    SAMPLE_REPORT.write_text(text, encoding="utf-8", newline="\n")
    print(f"Готово: {CLEAN.name}, {WITH_DEFECTS.name}, {SAMPLE_REPORT.name}")


if __name__ == "__main__":
    main()
