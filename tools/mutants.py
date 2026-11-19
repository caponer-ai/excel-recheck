"""Навмисні поломки коду: кожна мусить завалити хоча б один тест.

Якщо поломка «вижила», тести її не бачать, і зелений прогін нічого не доводить.
Запуск із кореня репозиторію, лише на чистому дереві git: python tools/mutants.py
Файл із поломкою відновлюється в finally.
"""

import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
MUTANTS = [
    (
        "M01 текст у сумі мовчки пропускається",
        "recheck/checks.py",
        "    if bad:\n",
        "    if False:\n",
    ),
    (
        "M02 порожній підсумок не помічається",
        "recheck/checks.py",
        "    if total.kind is Kind.EMPTY:\n",
        "    if False:\n",
    ),
    (
        "M03 усі доданки порожні = зійшлось",
        "recheck/checks.py",
        "    if empty == len(parts):\n",
        "    if False:\n",
    ),
    (
        "M04 мінус став плюсом",
        "recheck/checks.py",
        "total_of(plus) - total_of(minus)",
        "total_of(plus) + total_of(minus)",
    ),
    ("M05 множник загублено", "recheck/checks.py", ") * check.scale\n", ")\n"),
    (
        "M06 допуск розширено на копійку",
        "recheck/checks.py",
        "<= check.tol else",
        '<= check.tol + Decimal("0.01") else',
    ),
    (
        "M07 рядки не потрапляють у свою групу",
        "recheck/checks.py",
        "                open_group[2].append(row)\n",
        "                pass\n",
    ),
    (
        "M08 нерозпізнаний рядок ховається",
        "recheck/checks.py",
        "            if has_numbers:\n",
        "            if False:\n",
    ),
    (
        "M09 коди рядків рахуються як числа",
        "recheck/checks.py",
        "        if cell.ref in labels:\n",
        "        if False:\n",
    ),
    (
        "M10 число текстом не вважається проблемою",
        "recheck/checks.py",
        "_BAD_KINDS = (Kind.TEXT, Kind.TEXT_NUMBER, Kind.FORMULA_NO_VALUE)",
        "_BAD_KINDS = (Kind.TEXT, Kind.FORMULA_NO_VALUE)",
    ),
    (
        "M11 формула без значення не вважається проблемою",
        "recheck/checks.py",
        "_BAD_KINDS = (Kind.TEXT, Kind.TEXT_NUMBER, Kind.FORMULA_NO_VALUE)",
        "_BAD_KINDS = (Kind.TEXT, Kind.TEXT_NUMBER)",
    ),
    (
        "M12 float без repr (двійковий шум)",
        "recheck/book.py",
        "Decimal(repr(value))",
        "Decimal(value)",
    ),
    (
        "M13 bool рахується числом",
        "recheck/book.py",
        "    if isinstance(value, bool):",
        "    if False:",
    ),
    (
        "M14 код виходу ігнорує нерозпізнані рядки",
        "recheck/report.py",
        "    if tally[Status.SKIPPED] or run.unrecognized:\n",
        "    if tally[Status.SKIPPED]:\n",
    ),
    (
        "M15 текст у сумі не вважається проблемою",
        "recheck/checks.py",
        "_BAD_KINDS = (Kind.TEXT, Kind.TEXT_NUMBER, Kind.FORMULA_NO_VALUE)",
        "_BAD_KINDS = (Kind.TEXT_NUMBER, Kind.FORMULA_NO_VALUE)",
    ),
]


def git_status() -> str:
    done = subprocess.run(
        ["git", "status", "--porcelain"], cwd=REPO, capture_output=True, text=True, check=True
    )
    return done.stdout.strip()


def main() -> int:
    if git_status():
        print("Дерево не чисте, мутації не запускаю.", file=sys.stderr)
        return 2
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    killed = 0
    for name, rel, old, new in MUTANTS:
        path = REPO / rel
        original = path.read_bytes()
        text = original.decode("utf-8")
        if text.count(old) != 1:
            print(f"{name}: фрагмент знайдено {text.count(old)} разів, очікував 1", file=sys.stderr)
            return 2
        try:
            path.write_bytes(text.replace(old, new).encode("utf-8"))
            done = subprocess.run(
                [
                    sys.executable,
                    "-X",
                    "utf8",
                    "-m",
                    "pytest",
                    "-x",
                    "-q",
                    "-p",
                    "no:cacheprovider",
                ],
                cwd=REPO,
                capture_output=True,
                check=False,
            )
        finally:
            path.write_bytes(original)
        caught = done.returncode != 0
        killed += caught
        print(("СПІЙМАНО  " if caught else "ВИЖИВ     ") + name)
    print()
    print(f"Спіймано {killed} з {len(MUTANTS)}")
    if git_status():
        print("Після прогону дерево не чисте, перевір git status.", file=sys.stderr)
        return 2
    return 0 if killed == len(MUTANTS) else 1


if __name__ == "__main__":
    raise SystemExit(main())
