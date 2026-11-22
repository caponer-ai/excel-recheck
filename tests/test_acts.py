"""Порядкова звірка двох актів: кожна підкладена помилка мусить бути спіймана."""

from __future__ import annotations

import json
from decimal import Decimal
from pathlib import Path

import pytest
from openpyxl import load_workbook

from demo import make_acts_demo as demo
from recheck import acts
from recheck.book import Book
from recheck.checks import SpecError, Status


@pytest.fixture(scope="module")
def reference(tmp_path_factory: pytest.TempPathFactory) -> Path:
    path = tmp_path_factory.mktemp("acts") / "reference.xlsx"
    demo.build_reference(path)
    return path


def act_file(tmp_path: Path, ids: list[str] | None, name: str = "act.xlsx") -> Path:
    path = tmp_path / name
    demo.build_act(ids, path)
    return path


def check(act: Path, reference: Path) -> acts.Outcome:
    return acts.run(Book(act), Book(reference), acts.load_spec(demo.SPEC))


def not_ok(outcome: acts.Outcome) -> dict[str, str]:
    return {r.id: r.status.name for r in outcome.results if r.status is not Status.OK}


def test_clean_act_passes_without_a_single_remark(reference: Path, tmp_path: Path) -> None:
    outcome = check(act_file(tmp_path, []), reference)
    # 12 позицій × 5 перевірок (одиниця, кількість, ціна, вартість, арифметика) + 3 підсумки
    assert len(outcome.results) == 63
    assert not_ok(outcome) == {}
    assert outcome.strays == []
    assert acts.exit_code(outcome) == 0


def test_row_order_does_not_matter(reference: Path, tmp_path: Path) -> None:
    # чистий акт уже має переставлені позиції 2 і 3; звіримо й без перестановки
    path = tmp_path / "same_order.xlsx"
    demo._write(demo.act_rows([], reorder=False), path, reference=False)
    assert not_ok(check(path, reference)) == {}
    rows = demo.act_rows([])
    assert [row.key for row in rows[:3]] == ["1", "3", "2"]


@pytest.mark.parametrize("defect", demo.DEFECTS, ids=lambda defect: defect.id)
def test_each_defect_flips_exactly_its_checks(
    reference: Path, tmp_path: Path, defect: demo.Defect
) -> None:
    assert not_ok(check(act_file(tmp_path, [defect.id]), reference)) == defect.expect


def test_all_defects_together_do_not_hide_each_other(reference: Path, tmp_path: Path) -> None:
    expected: dict[str, str] = {}
    for defect in demo.DEFECTS:
        expected.update(defect.expect)
    outcome = check(act_file(tmp_path, None), reference)
    assert not_ok(outcome) == expected
    assert acts.exit_code(outcome) == 1


def test_wrong_unit_is_caught_although_every_sum_matches(reference: Path, tmp_path: Path) -> None:
    # суми й підсумок сходяться; звірка лише сум цієї помилки не побачила б
    outcome = check(act_file(tmp_path, ["A5"]), reference)
    totals = {r.id: r.status for r in outcome.results if r.field == "total"}
    assert set(totals.values()) == {Status.OK}
    assert not_ok(outcome) == {"8:unit": "DIFF"}


def test_kopeck_cut_off_is_a_discrepancy(reference: Path, tmp_path: Path) -> None:
    outcome = check(act_file(tmp_path, ["A7"]), reference)
    calc = next(r for r in outcome.results if r.id == "1:calc")
    # 412,5 × 86,37 = 35 627,625: округлення дає ,63, відкинута копійка ,62
    assert (calc.act, calc.reference) == (Decimal("35627.62"), Decimal("35627.63"))


def test_empty_act_reports_every_position_as_missing(reference: Path, tmp_path: Path) -> None:
    path = tmp_path / "empty.xlsx"
    demo._write([], path, reference=False)
    outcome = check(path, reference)
    missing = [r.key for r in outcome.results if r.field == "presence"]
    assert missing == [str(n) for n in range(1, 13)]
    assert {r.status for r in outcome.results if r.field == "presence"} == {Status.DIFF}
    assert not_ok(outcome)["total:act_vs_reference"] == "DIFF"


def test_single_row_act(reference: Path, tmp_path: Path) -> None:
    path = tmp_path / "one.xlsx"
    demo._write(demo.positions()[:1], path, reference=False)
    outcome = check(path, reference)
    assert {r.id: r.status for r in outcome.results if r.key == "1"} == {
        f"1:{name}": Status.OK for name in ("unit", "qty", "price", "amount", "calc")
    }
    assert sum(1 for r in outcome.results if r.field == "presence") == 11


def edit(path: Path, tmp_path: Path, change: dict[str, object], name: str = "e.xlsx") -> Path:
    book = load_workbook(path)
    sheet = book[demo.SHEET]
    for addr, value in change.items():
        sheet[addr] = value
    out = tmp_path / name
    book.save(out)
    return out


def total_row(path: Path, label_col: str) -> int:
    sheet = load_workbook(path)[demo.SHEET]
    return next(r for r in range(1, sheet.max_row + 1) if sheet[f"{label_col}{r}"].value == "Разом")


def test_row_without_position_number_is_reported(reference: Path, tmp_path: Path) -> None:
    clean = act_file(tmp_path, [])
    row = total_row(clean, "B")
    book = load_workbook(clean)
    book[demo.SHEET].insert_rows(row)
    book[demo.SHEET][f"B{row}"], book[demo.SHEET][f"F{row}"] = "Додаткові роботи", 1000.0
    path = tmp_path / "stray.xlsx"
    book.save(path)
    outcome = check(path, reference)
    assert [(s.row, s.reason) for s in outcome.strays] == [
        (row, "є числа, але немає номера позиції")
    ]
    assert acts.exit_code(outcome) == 2


def test_row_after_the_total_is_reported(reference: Path, tmp_path: Path) -> None:
    clean = act_file(tmp_path, [])
    row = total_row(clean, "B") + 1
    path = edit(clean, tmp_path, {f"A{row}": "14", f"F{row}": 500.0})
    outcome = check(path, reference)
    assert [s.reason for s in outcome.strays] == ["стоїть після підсумкового рядка"]
    assert acts.exit_code(outcome) == 2


def test_missing_total_row_is_not_a_pass(reference: Path, tmp_path: Path) -> None:
    clean = act_file(tmp_path, [])
    row = total_row(clean, "B")
    outcome = check(edit(clean, tmp_path, {f"B{row}": "Всього"}), reference)
    # без «Разом» рядок із сумою стає рядком без номера позиції, а підсумки не перевірено
    assert not_ok(outcome) == {"total:act_rows": "SKIPPED", "total:act_vs_reference": "SKIPPED"}
    assert acts.exit_code(outcome) == 2


def test_duplicate_in_the_reference_is_not_verified(reference: Path, tmp_path: Path) -> None:
    path = edit(reference, tmp_path, {"B6": "1"}, "ref_dup.xlsx")  # рядок 6 це позиція 2
    outcome = check(act_file(tmp_path, []), path)
    result = next(r for r in outcome.results if r.id == "1:presence")
    assert result.status is Status.SKIPPED
    assert "в еталоні рядки 5, 6" in result.reason


def test_unit_spelling_and_numeric_key_are_normalised(reference: Path, tmp_path: Path) -> None:
    clean = act_file(tmp_path, [])
    # у чистому акті позиція 1 стоїть у рядку 4: одиниця «м³», номер як число 1.0
    outcome = check(edit(clean, tmp_path, {"C4": "м³", "A4": 1.0}), reference)
    assert not_ok(outcome) == {}


def test_empty_unit_is_not_verified(reference: Path, tmp_path: Path) -> None:
    outcome = check(edit(act_file(tmp_path, []), tmp_path, {"C4": None}), reference)
    assert not_ok(outcome) == {"1:unit": "SKIPPED"}


def test_tolerance_hides_only_what_it_should(reference: Path, tmp_path: Path) -> None:
    spec = json.loads(demo.SPEC.read_text(encoding="utf-8"))
    spec["tolerance"] = "0.01"
    path = tmp_path / "tol.json"
    path.write_text(json.dumps(spec, ensure_ascii=False), encoding="utf-8")
    outcome = acts.run(Book(act_file(tmp_path, ["A7"])), Book(reference), acts.load_spec(path))
    assert not_ok(outcome) == {}


@pytest.mark.parametrize(
    ("change", "message"),
    [
        ({"tolerance": "-1"}, "не від'ємним"),
        ({"tolerance": "Infinity"}, "скінченним"),
        ({"tolerance": "abc"}, "не є числом"),
        ({"act.columns.qty": "4"}, "літерою"),
        ({"act.columns.qty": "A"}, "один стовпчик"),
        ({"act.first_row": 0}, "first_row"),
        ({"reference.total_label": " "}, "total_label"),
        ({"reference.columns": None}, "columns"),
    ],
)
def test_broken_spec_is_rejected(tmp_path: Path, change: dict[str, object], message: str) -> None:
    spec = json.loads(demo.SPEC.read_text(encoding="utf-8"))
    for dotted, value in change.items():
        *path, last = dotted.split(".")
        node = spec
        for part in path:
            node = node[part]
        node[last] = value
    file = tmp_path / "spec.json"
    file.write_text(json.dumps(spec, ensure_ascii=False), encoding="utf-8")
    with pytest.raises(SpecError, match=message):
        acts.load_spec(file)


def test_missing_column_is_rejected(tmp_path: Path) -> None:
    spec = json.loads(demo.SPEC.read_text(encoding="utf-8"))
    del spec["act"]["columns"]["price"]
    file = tmp_path / "spec.json"
    file.write_text(json.dumps(spec, ensure_ascii=False), encoding="utf-8")
    with pytest.raises(SpecError, match="price"):
        acts.load_spec(file)


def test_cli_exit_codes(
    reference: Path, tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    spec, ref = str(demo.SPEC), str(reference)
    assert (
        acts.main([str(act_file(tmp_path, [], "c.xlsx")), "--reference", ref, "--spec", spec]) == 0
    )
    assert (
        acts.main([str(act_file(tmp_path, ["A6"], "s.xlsx")), "--reference", ref, "--spec", spec])
        == 2
    )
    broken = act_file(tmp_path, ["A2"], "b.xlsx")
    report = tmp_path / "report.md"
    args = [str(broken), "--reference", ref, "--spec", spec, "--out", str(report)]
    assert acts.main(args) == 1
    text = report.read_text(encoding="utf-8")
    assert "позиції немає в акті" in text
    assert "## Звідки числа" in text
    assert acts.main([str(tmp_path / "нема.xlsx"), "--reference", ref, "--spec", spec]) == 3
    bad = tmp_path / "bad.xlsx"
    bad.write_bytes(b"not a spreadsheet")
    assert acts.main([str(bad), "--reference", ref, "--spec", spec]) == 3
    nowhere = tmp_path / "немає" / "report.md"
    assert acts.main([str(broken), "--reference", ref, "--spec", spec, "--out", str(nowhere)]) == 3
    assert "не записано" in capsys.readouterr().err


def test_committed_files_match_the_generator(tmp_path: Path) -> None:
    # README посилається на зразок звіту; він не має розійтися з кодом
    fresh_ref, fresh_act = tmp_path / "ref.xlsx", tmp_path / "act.xlsx"
    demo.build_reference(fresh_ref)
    demo.build_act(None, fresh_act)
    fresh = acts.render(
        check(fresh_act, fresh_ref), demo.WITH_DEFECTS.name, demo.REFERENCE.name, demo.SPEC.name
    )
    assert demo.SAMPLE_REPORT.read_text(encoding="utf-8") == fresh
    committed = acts.render(
        check(demo.WITH_DEFECTS, demo.REFERENCE),
        demo.WITH_DEFECTS.name,
        demo.REFERENCE.name,
        demo.SPEC.name,
    )
    assert committed == fresh
