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
        "        elif has_content(row):\n",
        "        elif False:\n",
    ),
    (
        "M16 рядок з чужим кодом у групі ховається",
        "recheck/checks.py",
        "if tail == label or tail[:1].isalnum():",
        "if False:",
    ),
    (
        "M17 рядки після підсумку ховаються",
        "recheck/checks.py",
        "            if label or has_content(row):\n",
        "            if False:\n",
    ),
    (
        "M18 нескінченний допуск приймається",
        "recheck/checks.py",
        "    if not number.is_finite():\n",
        "    if False:\n",
    ),
    (
        "M19 від'ємний допуск приймається",
        "recheck/checks.py",
        '        if _decimal(item.get("tol", "0"), where) < 0:\n',
        "        if False:\n",
    ),
    (
        "M20 нульовий множник приймається",
        "recheck/checks.py",
        '        if _decimal(item.get("scale", "1"), where) == 0:\n',
        "        if False:\n",
    ),
    (
        "M21 підсумок серед власних доданків приймається",
        "recheck/checks.py",
        "    if total.ref in {cell.ref for cell in parts}:\n",
        "    if False:\n",
    ),
    (
        "M22 зіпсований файл валить програму замість коду 3",
        "recheck/__main__.py",
        "except (InvalidFileException, BadZipFile, KeyError) as error:",
        "except (RuntimeError,) as error:",
    ),
    (
        "M23 двійковий шум у файлі стає розбіжністю",
        "recheck/book.py",
        'format(value, ".15g")',
        'format(value, ".17g")',
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
        "M12 float береться як є (двійковий хвіст)",
        "recheck/book.py",
        'Decimal(format(value, ".15g"))',
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
        "M24 рядок без звичайних чисел мовчки зникає з таблиці",
        "recheck/checks.py",
        "kind is not Kind.EMPTY for col in watched",
        "kind is Kind.NUMBER for col in watched",
    ),
    (
        "M25 належність до групи за самим початком коду",
        "recheck/checks.py",
        "if tail == label or tail[:1].isalnum():",
        "if tail == label:",
    ),
    (
        "M26 регулярний вираз у правилах не перевіряється",
        "recheck/checks.py",
        "        re.compile(str(value))\n",
        "        pass\n",
    ),
    (
        "M27 таблиця без жодної перевірки приймається",
        "recheck/checks.py",
        "        if not sum_cols and not row_rule:\n",
        "        if False:\n",
    ),
    (
        "M28 мінус U+2212 не розпізнається",
        "recheck/book.py",
        r'_TEXT_NUMBER = re.compile(r"[-\u2212]?',
        r'_TEXT_NUMBER = re.compile(r"-?',
    ),
    (
        "M29 від'ємне в дужках не розпізнається",
        "recheck/book.py",
        "        if bracketed:\n",
        "        if False:\n",
    ),
    (
        "M30 правила не в UTF-8 валять програму",
        "recheck/checks.py",
        "    except UnicodeDecodeError:\n",
        "    except UnicodeEncodeError:\n",
    ),
    (
        "M31 нуль перевірок вважається успіхом",
        "recheck/checks.py",
        "    if not results:\n",
        "    if False:\n",
    ),
    (
        "M32 список правил не перевіряється на тип елементів",
        "recheck/checks.py",
        "not all(isinstance(item, dict) for item in value)",
        "False",
    ),
    (
        "M33 цифра замість літери стовпчика валить програму",
        "recheck/checks.py",
        "    if not isinstance(value, str) or not _COLUMN.fullmatch(value):\n",
        "    if False:\n",
    ),
    (
        "M34 звіт, який не записався, дає код «є розбіжності»",
        "recheck/__main__.py",
        "        except OSError as error:\n",
        "        except RuntimeError as error:\n",
    ),
    (
        "M35 допуск, не менший за самі числа, дає «зійшлось»",
        "recheck/checks.py",
        "    if check.tol > 0 and 0 < largest <= check.tol:\n",
        "    if False:\n",
    ),
    (
        "M36 рядок-деталь до першої групи ховається",
        "recheck/checks.py",
        "            elif group is not None:\n",
        "            elif False:\n",
    ),
    (
        "M15 текст у сумі не вважається проблемою",
        "recheck/checks.py",
        "_BAD_KINDS = (Kind.TEXT, Kind.TEXT_NUMBER, Kind.FORMULA_NO_VALUE)",
        "_BAD_KINDS = (Kind.TEXT_NUMBER, Kind.FORMULA_NO_VALUE)",
    ),
    (
        "M37 пропущена позиція мовчки зникає",
        "recheck/acts.py",
        "        if a is None and r is not None:\n",
        "        if a is None and r is not None:\n            continue\n",
    ),
    (
        "M38 зайва позиція мовчки зникає",
        "recheck/acts.py",
        "        if r is None and a is not None:\n",
        "        if r is None and a is not None:\n            continue\n",
    ),
    (
        "M39 дубль номера позиції не помічається",
        "recheck/acts.py",
        "        if key in act_dupes or key in ref_dupes:\n",
        "        if False:\n",
    ),
    (
        "M40 дубль в еталоні не помічається",
        "recheck/acts.py",
        "        if key in act_dupes or key in ref_dupes:\n",
        "        if key in act_dupes:\n",
    ),
    (
        "M41 одиниця виміру не звіряється",
        "recheck/acts.py",
        "status = Status.OK if _unit(ua) == _unit(ur) else Status.DIFF",
        "status = Status.OK",
    ),
    (
        "M42 порожня одиниця не дає «не перевірено»",
        "recheck/acts.py",
        "    if not _unit(ua) or not _unit(ur):\n",
        "    if False:\n",
    ),
    (
        "M43 число текстом вважається числом",
        "recheck/acts.py",
        "    bad = [c for c in (a, r) if c.kind is not Kind.NUMBER]\n",
        "    bad = [c for c in (a, r) if c.kind is Kind.TEXT]\n",
    ),
    (
        "M44 арифметика рядка не перевіряється",
        "recheck/acts.py",
        "status = Status.OK if abs(amount.value - mine) <= tol else Status.DIFF",
        "status = Status.OK",
    ),
    (
        "M45 перерахунок відкидає копійку замість округлення",
        "recheck/acts.py",
        "quantize(_CENT, ROUND_HALF_UP)",
        'quantize(_CENT, "ROUND_DOWN")',
    ),
    (
        "M46 допуск полів розширено на копійку",
        "recheck/acts.py",
        "status = Status.OK if abs(a.value - r.value) <= tol else Status.DIFF",
        'status = Status.OK if abs(a.value - r.value) <= tol + Decimal("0.01") else Status.DIFF',
    ),
    (
        "M47 рядок без номера позиції ховається",
        "recheck/acts.py",
        "        if not key:\n",
        "        if not key:\n            continue\n",
    ),
    (
        "M48 «Разом» акта не звіряється з еталоном",
        "recheck/acts.py",
        "status = Status.OK if abs(total.value - other) <= tol else Status.DIFF",
        "status = Status.OK",
    ),
    (
        "M49 від'ємний допуск приймається",
        "recheck/acts.py",
        "    if not tol.is_finite() or tol < 0:\n",
        "    if not tol.is_finite():\n",
    ),
    (
        "M50 рядки після «Разом» ховаються",
        "recheck/acts.py",
        "            if filled or label:\n",
        "            if False:\n",
    ),
    (
        "M51 два поля на один стовпчик приймаються",
        "recheck/acts.py",
        "    if len(set(used.values())) != len(used):\n",
        "    if False:\n",
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
        if done.returncode not in (0, 1):  # 1 це «тести впали»; решта це зламаний прогін
            print(f"{name}: pytest завершився з кодом {done.returncode}", file=sys.stderr)
            return 2
        caught = done.returncode == 1
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
