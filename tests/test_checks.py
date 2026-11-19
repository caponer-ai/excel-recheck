"""Правила однієї перевірки на маленьких файлах, зібраних у тесті."""

from __future__ import annotations

from decimal import Decimal
from pathlib import Path

import pytest
from openpyxl import Workbook

from recheck.book import Book, Kind, RefError, classify, parse_ref
from recheck.checks import Check, SpecError, Status, evaluate, parse_spec, run


def make_book(tmp_path: Path, cells: dict[str, object]) -> Book:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Аркуш"
    for addr, value in cells.items():
        sheet[addr] = value
    path = tmp_path / "book.xlsx"
    workbook.save(path)
    return Book(path)


def total_check(**overrides: object) -> Check:
    fields: dict[str, object] = {
        "id": "t",
        "group": "Г",
        "name": "підсумок = сума",
        "total": "Аркуш!A4",
        "plus": ("Аркуш!A1:A3",),
    }
    fields.update(overrides)
    return Check(**fields)


def test_exact_match_to_the_kopeck(tmp_path: Path) -> None:
    book = make_book(tmp_path, {"A1": 100.10, "A2": 200.20, "A3": 0.03, "A4": 300.33})
    result = evaluate(book, total_check())
    assert result.status is Status.OK
    assert result.mine == Decimal("300.33")


def test_one_kopeck_is_a_discrepancy(tmp_path: Path) -> None:
    book = make_book(tmp_path, {"A1": 100.10, "A2": 200.20, "A3": 0.03, "A4": 300.34})
    result = evaluate(book, total_check())
    assert result.status is Status.DIFF
    assert result.diff == Decimal("-0.01")


def test_float_noise_is_not_a_discrepancy(tmp_path: Path) -> None:
    # у float 0.1 + 0.2 != 0.3; на цьому ламається наївна звірка
    assert 0.1 + 0.2 != 0.3
    book = make_book(tmp_path, {"A1": 0.1, "A2": 0.2, "A4": 0.3})
    assert evaluate(book, total_check()).status is Status.OK


def test_number_typed_as_text_is_not_verified(tmp_path: Path) -> None:
    book = make_book(tmp_path, {"A1": 100, "A2": "1 250,50", "A3": 5, "A4": 1355.5})
    result = evaluate(book, total_check())
    assert result.status is Status.SKIPPED
    assert "Аркуш!A2" in result.reason
    assert "число, записане текстом" in result.reason


def test_text_in_a_sum_is_not_verified_even_if_the_rest_adds_up(tmp_path: Path) -> None:
    book = make_book(tmp_path, {"A1": 100, "A2": "див. примітку", "A3": 5, "A4": 105})
    assert evaluate(book, total_check()).status is Status.SKIPPED


def test_empty_total_is_not_verified(tmp_path: Path) -> None:
    book = make_book(tmp_path, {"A1": 1, "A2": 2, "A3": 3})
    result = evaluate(book, total_check())
    assert result.status is Status.SKIPPED
    assert "порожня" in result.reason


def test_all_parts_empty_is_not_a_pass(tmp_path: Path) -> None:
    # нуль дорівнює нулю, але перевіряти тут нічого: це не «зійшлось»
    book = make_book(tmp_path, {"A4": 0})
    assert evaluate(book, total_check()).status is Status.SKIPPED


def test_empty_parts_are_counted_and_shown(tmp_path: Path) -> None:
    book = make_book(tmp_path, {"A1": 10, "A3": 5, "A4": 15})
    result = evaluate(book, total_check())
    assert result.status is Status.OK
    assert result.empty_parts == 1


def test_formula_without_saved_value_is_not_verified(tmp_path: Path) -> None:
    # файл, збережений без перерахунку: формула є, значення немає
    book = make_book(tmp_path, {"A1": 1, "A2": 2, "A3": 3, "A4": "=SUM(A1:A3)"})
    result = evaluate(book, total_check())
    assert result.status is Status.SKIPPED
    assert "формула без збереженого значення" in result.reason


def test_minus_scale_and_tolerance(tmp_path: Path) -> None:
    book = make_book(tmp_path, {"A1": 1_500_000, "A2": 250_000, "A4": 1250.04})
    check = total_check(
        plus=("Аркуш!A1",),
        minus=("Аркуш!A2",),
        scale=Decimal("0.001"),
        tol=Decimal("0.05"),
    )
    assert evaluate(book, check).status is Status.OK
    strict = total_check(plus=("Аркуш!A1",), minus=("Аркуш!A2",), scale=Decimal("0.001"))
    assert evaluate(book, strict).status is Status.DIFF


def test_total_must_be_one_cell(tmp_path: Path) -> None:
    book = make_book(tmp_path, {"A1": 1})
    with pytest.raises(SpecError):
        evaluate(book, total_check(total="Аркуш!A1:A2"))


@pytest.mark.parametrize(
    ("value", "kind"),
    [
        (None, Kind.EMPTY),
        ("   ", Kind.EMPTY),
        (True, Kind.TEXT),
        (12, Kind.NUMBER),
        (12.5, Kind.NUMBER),
        ("12,5", Kind.TEXT_NUMBER),
        ("1 234.50", Kind.TEXT_NUMBER),
        ("-7", Kind.TEXT_NUMBER),
        ("12 шт", Kind.TEXT),
        ("1.2.3", Kind.TEXT),
    ],
)
def test_classify(value: object, kind: Kind) -> None:
    assert classify(value)[0] is kind


def test_parse_ref_range_and_errors() -> None:
    assert parse_ref("Оборот складу!B2:C3") == (
        "Оборот складу",
        ["B2", "C2", "B3", "C3"],
    )
    assert parse_ref("Аркуш!A1") == ("Аркуш", ["A1"])
    for bad in ("A1", "Аркуш!", "Аркуш!a1", "Аркуш!A0"):
        with pytest.raises(RefError):
            parse_ref(bad)


def test_missing_sheet_is_a_clear_error(tmp_path: Path) -> None:
    book = make_book(tmp_path, {"A1": 1})
    spec = parse_spec(
        {
            "checks": [
                {
                    "id": "x",
                    "group": "Г",
                    "name": "н",
                    "total": "Нема!A1",
                    "plus": ["Аркуш!A1"],
                }
            ]
        }
    )
    with pytest.raises(SpecError, match="Аркуша «Нема» немає"):
        run(book, spec)


@pytest.mark.parametrize(
    "data",
    [
        {},
        {"checks": []},
        {"other": 1},
        {"checks": [{"id": "x", "group": "Г", "name": "н", "total": "А!A1"}]},
        {
            "checks": [
                {
                    "id": "x",
                    "group": "Г",
                    "name": "н",
                    "total": "А!A1",
                    "plus": [],
                    "typo": 1,
                }
            ]
        },
        {
            "checks": [
                {
                    "id": "x",
                    "group": "Г",
                    "name": "н",
                    "total": "А!A1",
                    "plus": ["А!A2"],
                },
                {
                    "id": "x",
                    "group": "Г",
                    "name": "н",
                    "total": "А!A1",
                    "plus": ["А!A2"],
                },
            ]
        },
        {
            "checks": [
                {
                    "id": "x",
                    "group": "Г",
                    "name": "н",
                    "total": "А!A1",
                    "plus": [],
                    "tol": "багато",
                }
            ]
        },
    ],
)
def test_bad_spec_is_rejected(data: object) -> None:
    with pytest.raises(SpecError):
        parse_spec(data)


def test_binary_noise_stored_in_the_file_is_not_a_discrepancy() -> None:
    # результат формули 0.1 + 0.2 лежить у файлі як 0.30000000000000004
    assert classify(0.1 + 0.2) == (Kind.NUMBER, Decimal("0.3"))
    assert classify(1234567.89) == (Kind.NUMBER, Decimal("1234567.89"))
    assert classify(0.01) == (Kind.NUMBER, Decimal("0.01"))


def test_total_among_its_own_parts_is_a_spec_error(tmp_path: Path) -> None:
    book = make_book(tmp_path, {"A1": 1, "A2": 2, "A3": 3, "A4": 6})
    with pytest.raises(SpecError, match="серед власних доданків"):
        evaluate(book, total_check(plus=("Аркуш!A1:A4",)))


@pytest.mark.parametrize(
    "extra",
    [{"tol": "Infinity"}, {"tol": "-0.01"}, {"tol": "NaN"}, {"scale": "0"}, {"scale": "Infinity"}],
)
def test_tolerance_and_scale_that_switch_the_check_off_are_rejected(extra: dict[str, str]) -> None:
    item = {"id": "x", "group": "Г", "name": "н", "total": "А!A1", "plus": ["А!A2"], **extra}
    with pytest.raises(SpecError):
        parse_spec({"checks": [item]})
