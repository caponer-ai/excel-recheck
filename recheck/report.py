"""Звіт у трьох блоках: що зійшлось, що не перевірено і чому, звідки числа."""

from __future__ import annotations

from decimal import Decimal

from recheck.book import Kind
from recheck.checks import Result, Run, Status


def money(value: Decimal | None) -> str:
    """1234567.8 -> «1 234 567,80». Більше двох знаків після коми не обрізаємо."""
    if value is None:
        return "немає"
    exponent = value.as_tuple().exponent
    digits = max(2, -exponent if isinstance(exponent, int) else 2)
    text = f"{value:,.{digits}f}"
    return text.replace(",", " ").replace(".", ",").replace("-", "−")


def counts(run: Run) -> dict[Status, int]:
    return {status: sum(1 for r in run.results if r.status is status) for status in Status}


def exit_code(run: Run) -> int:
    """0 усе зійшлось; 1 є розбіжності; 2 розбіжностей нема, але щось не перевірено."""
    tally = counts(run)
    if tally[Status.DIFF]:
        return 1
    if tally[Status.SKIPPED] or run.unrecognized:
        return 2
    return 0


def summary(run: Run) -> str:
    tally = counts(run)
    uncovered = sum(len(cells) for cells in run.uncovered.values())
    return (
        f"Перевірок: {len(run.results)}. Зійшлось: {tally[Status.OK]}. "
        f"Розбіжностей: {tally[Status.DIFF]}. Не перевірено: {tally[Status.SKIPPED]}.\n"
        f"Чисел у файлі: {run.numeric_total}. Поза перевірками: {uncovered}. "
        f"Нерозпізнаних рядків таблиць: {len(run.unrecognized)}."
    )


def _sources(result: Result) -> str:
    check = result.check
    text = f"`{check.total}` = " + " + ".join(f"`{ref}`" for ref in check.plus)
    text += "".join(f" − `{ref}`" for ref in check.minus)
    if check.scale != 1:
        text += f", множник {check.scale}"
    if check.tol != 0:
        text += f", допуск {money(check.tol)}"
    if result.empty_parts:
        text += f", порожніх доданків: {result.empty_parts}"
    if check.note:
        text += f". {check.note}"
    return text


def render(run: Run, file_name: str, spec_name: str) -> str:
    """Markdown-звіт. Часу прогону тут немає, щоб звіт був відтворюваним."""
    lines = [
        f"# Звіт перевірки: {file_name}",
        "",
        f"Правила: `{spec_name}`",
        "",
        summary(run),
        "",
    ]

    diffs = [r for r in run.results if r.status is Status.DIFF]
    lines += ["## Розбіжності", ""]
    if diffs:
        lines += [
            "| Перевірка | У файлі | Перерахунок | Різниця | Клітинка |",
            "|---|---:|---:|---:|---|",
        ]
        for r in diffs:
            lines.append(
                f"| {r.check.group}. {r.check.name} | {money(r.theirs)} | {money(r.mine)} "
                f"| {money(r.diff)} | `{r.check.total}` |"
            )
    else:
        lines.append("Немає.")
    lines.append("")

    lines += ["## Що не перевірено і чому", ""]
    skipped = [r for r in run.results if r.status is Status.SKIPPED]
    lines += ["### Перевірки, які не вдалося виконати", ""]
    if skipped:
        for r in skipped:
            lines.append(f"- {r.check.group}. {r.check.name}: {r.reason}")
    else:
        lines.append("Немає.")
    lines += ["", "### Рядки таблиць з числами, які не розпізнано", ""]
    if run.unrecognized:
        for row in run.unrecognized:
            lines.append(f"- {row.sheet}, р. {row.row}: «{row.label}»")
    else:
        lines.append("Немає.")
    lines += ["", "### Числа, яких не торкнулась жодна перевірка", ""]
    if run.uncovered:
        for sheet, cells in run.uncovered.items():
            shown = ", ".join(cell.addr for cell in cells[:8])
            more = f" і ще {len(cells) - 8}" if len(cells) > 8 else ""
            as_text = sum(1 for cell in cells if cell.kind is Kind.TEXT_NUMBER)
            note = f" (з них записано текстом: {as_text})" if as_text else ""
            lines.append(f"- {sheet}: {len(cells)}{note}. Адреси: {shown}{more}")
    else:
        lines.append("Немає: кожне число у файлі входить хоча б в одну перевірку.")
    lines.append("")

    lines += ["## Що зійшлось", ""]
    by_group: dict[str, int] = {}
    for r in run.results:
        if r.status is Status.OK:
            by_group[r.check.group] = by_group.get(r.check.group, 0) + 1
    if by_group:
        lines += [f"- {group}: {count}" for group, count in by_group.items()]
    else:
        lines.append("Нічого.")
    lines.append("")

    lines += ["## Звідки числа", ""]
    for r in run.results:
        lines.append(f"- [{r.status.value}] {r.check.group}. {r.check.name}: {_sources(r)}")
    lines.append("")
    return "\n".join(lines)
