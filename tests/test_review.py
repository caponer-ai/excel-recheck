"""Діри, знайдені рев'ю після публікації: кожна тут має тест, який на старому коді падав."""

from __future__ import annotations

import json
from decimal import Decimal
from pathlib import Path

import pytest
from openpyxl import Workbook

from recheck.__main__ import main
from recheck.book import Book, Kind, classify
from recheck.checks import Run, Spec, SpecError, Status, parse_spec, run
from recheck.report import exit_code

TABLE: dict[str, object] = {
    "id": "t",
    "sheet": "Аркуш",
    "first_row": 2,
    "label_col": "A",
    "total_label": "Разом",
    "detail_regex": r"\d+\.\d+",
    "group_regex": r"\d+",
    "sum_cols": ["B"],
}
CHECK: dict[str, object] = {
    "id": "x",
    "group": "Г",
    "name": "н",
    "total": "Аркуш!B6",
    "plus": ["Аркуш!B3:B4"],
}


def save(tmp_path: Path, rows: list[tuple[object, object]]) -> Path:
    """Таблиця з кодом у стовпчику A і сумою в B, перший рядок це шапка."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Аркуш"
    sheet["A1"], sheet["B1"] = "Код", "Сума"
    for number, (label, amount) in enumerate(rows, start=2):
        sheet[f"A{number}"], sheet[f"B{number}"] = label, amount
    path = tmp_path / "table.xlsx"
    workbook.save(path)
    return path


def check_table(path: Path, **overrides: object) -> Run:
    return run(Book(path), parse_spec({"tables": [{**TABLE, **overrides}]}))


@pytest.mark.parametrize("amount", ["(5,00)", "−5,00", "=1+1", "уточнити"])
def test_row_without_plain_numbers_is_not_dropped_silently(tmp_path: Path, amount: str) -> None:
    # суми сходяться, а зайвий рядок з від'ємним у дужках, формулою чи текстом стоїть у таблиці
    rows = [("201", 30), ("201.001", 10), ("201.002", 20), ("Примітка", amount), ("Разом", 30)]
    outcome = check_table(save(tmp_path, rows))
    assert {result.status for result in outcome.results} == {Status.OK}
    assert [(row.row, row.label) for row in outcome.unrecognized] == [(5, "Примітка")]
    assert exit_code(outcome) == 2


def test_code_that_only_shares_digits_with_the_group_is_reported(tmp_path: Path) -> None:
    # «2015.002» починається з «201», але це код іншої групи
    rows = [("201", 30), ("201.001", 10), ("2015.002", 20), ("Разом", 30)]
    outcome = check_table(save(tmp_path, rows))
    assert [(row.row, row.label) for row in outcome.unrecognized] == [(4, "2015.002")]
    assert "групи «201»" in outcome.unrecognized[0].reason


@pytest.mark.parametrize(
    ("text", "number"),
    [
        ("−500,00", "-500.00"),  # мінус U+2212: саме так його друкує наш власний звіт
        ("(1 234,56)", "-1234.56"),
        ("1 234,56", "1234.56"),
        ("1 234,56", "1234.56"),
    ],
)
def test_negative_and_spaced_text_numbers_are_recognised(text: str, number: str) -> None:
    assert classify(text) == (Kind.TEXT_NUMBER, Decimal(number))


@pytest.mark.parametrize(
    "data",
    [
        {"checks": 5},
        {"checks": [5]},
        {"checks": [{**CHECK, "plus": "Аркуш!B3"}]},
        {"checks": [{**CHECK, "minus": [1]}]},
        {"tables": "x"},
        {"tables": [{**TABLE, "detail_regex": "("}]},
        {"tables": [{**TABLE, "group_regex": "["}]},
        {"tables": [{**TABLE, "first_row": "x"}]},
        {"tables": [{**TABLE, "first_row": 0}]},
        {"tables": [{**TABLE, "first_row": True}]},
        {"tables": [{**TABLE, "sum_cols": "B"}]},
        {"tables": [{**TABLE, "sum_cols": []}]},
        {"tables": [{**TABLE, "row_rule": 5}]},
    ],
)
def test_malformed_rules_are_a_spec_error_not_a_crash(data: object) -> None:
    with pytest.raises(SpecError):
        parse_spec(data)


def test_zero_checks_is_never_a_pass(tmp_path: Path) -> None:
    path = save(tmp_path, [("201", 30), ("201.001", 30), ("Разом", 30)])
    with pytest.raises(SpecError, match="жодної перевірки"):
        run(Book(path), Spec())


@pytest.mark.parametrize(
    "rules",
    [
        json.dumps({"tables": [{**TABLE, "detail_regex": "("}]}).encode(),
        json.dumps({"tables": [{**TABLE, "sum_cols": []}]}).encode(),
        json.dumps({"checks": 5}).encode(),
        '{"checks": [], "примітка": "кирилиця"}'.encode("cp1251"),
        b"not json",
    ],
)
def test_broken_rules_file_gives_exit_code_3(tmp_path: Path, rules: bytes) -> None:
    # код 1 означає «є розбіжності», тому зламані правила не мають права його повертати
    book = save(tmp_path, [("201", 30), ("201.001", 30), ("Разом", 30)])
    spec = tmp_path / "rules.json"
    spec.write_bytes(rules)
    assert main([str(book), "--checks", str(spec)]) == 3
