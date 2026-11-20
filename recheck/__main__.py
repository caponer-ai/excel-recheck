"""Командний рядок: python -m recheck ФАЙЛ.xlsx --checks ПРАВИЛА.json [--out ЗВІТ.md]"""

from __future__ import annotations

import argparse
import sys
import time
from collections.abc import Sequence
from pathlib import Path
from zipfile import BadZipFile

from openpyxl.utils.exceptions import InvalidFileException

from recheck.book import Book, RefError
from recheck.checks import SpecError, Status, load_spec, run
from recheck.report import exit_code, money, render, summary


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="recheck",
        description="Незалежний перерахунок підсумків Excel-звіту.",
    )
    parser.add_argument("file", type=Path, help="звіт у форматі .xlsx")
    parser.add_argument("--checks", type=Path, required=True, help="файл правил .json")
    parser.add_argument("--out", type=Path, help="куди записати звіт у Markdown")
    args = parser.parse_args(argv)

    for stream in (sys.stdout, sys.stderr):  # консоль Windows інакше падає на знаку «−»
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")

    started = time.perf_counter()
    try:
        spec = load_spec(args.checks)
        try:
            book = Book(args.file)
        except (InvalidFileException, BadZipFile, KeyError) as error:
            raise SpecError(f"{args.file.name} не читається як .xlsx ({error})") from None
        outcome = run(book, spec)
    except (SpecError, RefError, OSError) as error:
        print(f"Не можу почати перевірку: {error}", file=sys.stderr)
        return 3
    elapsed = time.perf_counter() - started

    print(summary(outcome))
    for result in outcome.results:
        if result.status is Status.DIFF:
            print(
                f"РОЗБІЖНІСТЬ  {result.check.group}. {result.check.name}: "
                f"у файлі {money(result.theirs)}, перерахунок {money(result.mine)}, "
                f"різниця {money(result.diff)} ({result.check.total})"
            )
        elif result.status is Status.SKIPPED:
            print(f"НЕ ПЕРЕВІРЕНО  {result.check.group}. {result.check.name}: {result.reason}")
    for row in outcome.unrecognized:
        print(f"НЕ РОЗПІЗНАНО  {row.sheet}, р. {row.row}: «{row.label}»")
    print(f"Час прогону: {elapsed:.2f} с")

    if args.out:
        text = render(outcome, args.file.name, args.checks.name)
        try:
            args.out.write_text(text, encoding="utf-8", newline="\n")
        except OSError as error:
            # код 1 означав би «є розбіжності», а звіту про них немає
            print(f"Перевірку виконано, але звіт не записано: {error}", file=sys.stderr)
            return 3
        print(f"Звіт: {args.out}")
    return exit_code(outcome)


if __name__ == "__main__":
    raise SystemExit(main())
