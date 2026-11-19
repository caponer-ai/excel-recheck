"""Кожна підкладена помилка мусить бути спіймана, а чистий файл мусить пройти чисто."""

from __future__ import annotations

from pathlib import Path

import pytest

from demo import make_demo
from recheck.__main__ import main
from recheck.book import Book
from recheck.checks import Run, Status, load_spec, run
from recheck.report import exit_code, render


@pytest.fixture(scope="module")
def clean(tmp_path_factory: pytest.TempPathFactory) -> Path:
    path = tmp_path_factory.mktemp("demo") / "clean.xlsx"
    make_demo.build_clean(path)
    return path


def check(path: Path) -> Run:
    return run(Book(path), load_spec(make_demo.SPEC))


def statuses(outcome: Run) -> dict[str, str]:
    return {result.check.id: result.status.name for result in outcome.results}


def test_clean_file_passes_without_a_single_remark(clean: Path) -> None:
    outcome = check(clean)
    assert len(outcome.results) == 40
    assert {result.status for result in outcome.results} == {Status.OK}
    assert outcome.unrecognized == []
    assert exit_code(outcome) == 0


@pytest.mark.parametrize("defect", make_demo.DEFECTS, ids=lambda defect: defect.id)
def test_each_defect_flips_exactly_its_checks(
    clean: Path, tmp_path: Path, defect: make_demo.Defect
) -> None:
    broken = tmp_path / "broken.xlsx"
    make_demo.build_with_defects([defect.id], source=clean, target=broken)
    changed = {ident: name for ident, name in statuses(check(broken)).items() if name != "OK"}
    assert changed == defect.expect


def test_all_defects_together_do_not_hide_each_other(clean: Path, tmp_path: Path) -> None:
    broken = tmp_path / "broken.xlsx"
    make_demo.build_with_defects(source=clean, target=broken)
    expected: dict[str, str] = {}
    for defect in make_demo.DEFECTS:
        expected.update(defect.expect)
    outcome = check(broken)
    changed = {ident: name for ident, name in statuses(outcome).items() if name != "OK"}
    assert changed == expected
    assert exit_code(outcome) == 1


def test_one_kopeck_is_invisible_to_the_rounded_check(clean: Path, tmp_path: Path) -> None:
    # перевірка в тисячах з допуском копійку не бачить; тому точна перевірка обов'язкова
    broken = tmp_path / "broken.xlsx"
    make_demo.build_with_defects(["D4"], source=clean, target=broken)
    result = statuses(check(broken))
    assert result["sum.income_total"] == "OK"
    assert result["inc.total"] == "DIFF"


def test_uncovered_numbers_are_listed_not_hidden(clean: Path) -> None:
    outcome = check(clean)
    assert outcome.numeric_total == 111
    assert {sheet: len(cells) for sheet, cells in outcome.uncovered.items()} == {
        "Доходи": 8,
        "Витрати": 8,
    }
    assert outcome.uncovered["Доходи"][0].ref == "Доходи!B4"


def test_row_that_fits_no_rule_is_reported(clean: Path, tmp_path: Path) -> None:
    from openpyxl import load_workbook

    workbook = load_workbook(clean)
    sheet = workbook[make_demo.STOCK]
    sheet.insert_rows(6)
    sheet["A6"], sheet["D6"] = "у т.ч. прострочені", 100.0
    path = tmp_path / "odd.xlsx"
    workbook.save(path)
    outcome = check(path)
    assert [(row.row, row.label) for row in outcome.unrecognized] == [(6, "у т.ч. прострочені")]
    assert exit_code(outcome) == 2


def test_committed_files_match_the_generator(tmp_path: Path) -> None:
    # README посилається на зразок звіту; він не має розійтися з кодом
    fresh_clean, fresh_broken = tmp_path / "clean.xlsx", tmp_path / "broken.xlsx"
    make_demo.build_clean(fresh_clean)
    make_demo.build_with_defects(source=fresh_clean, target=fresh_broken)
    fresh = render(check(fresh_broken), make_demo.WITH_DEFECTS.name, make_demo.SPEC.name)
    assert make_demo.SAMPLE_REPORT.read_text(encoding="utf-8") == fresh
    committed = render(
        check(make_demo.WITH_DEFECTS), make_demo.WITH_DEFECTS.name, make_demo.SPEC.name
    )
    assert committed == fresh


def test_cli_exit_codes(clean: Path, tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    spec = str(make_demo.SPEC)
    assert main([str(clean), "--checks", spec]) == 0

    broken = tmp_path / "broken.xlsx"
    make_demo.build_with_defects(["D2"], source=clean, target=broken)
    assert main([str(broken), "--checks", spec]) == 2

    make_demo.build_with_defects(["D2", "D4"], source=clean, target=broken)
    report = tmp_path / "report.md"
    assert main([str(broken), "--checks", spec, "--out", str(report)]) == 1
    text = report.read_text(encoding="utf-8")
    assert "## Що не перевірено і чому" in text
    assert "## Звідки числа" in text

    assert main([str(tmp_path / "нема.xlsx"), "--checks", spec]) == 3
    bad_spec = tmp_path / "bad.json"
    bad_spec.write_text("{", encoding="utf-8")
    assert main([str(clean), "--checks", str(bad_spec)]) == 3
    assert "Не можу почати перевірку" in capsys.readouterr().err
