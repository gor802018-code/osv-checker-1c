from __future__ import annotations

import argparse
import csv
import io
import json
import re
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import Workbook, load_workbook

TOLERANCE = 0.5

ACTIVE = {"01", "03", "04", "07", "08", "10", "11", "19", "20", "21", "23", "25",
          "26", "28", "29", "41", "43", "44", "45", "50", "51", "52", "55", "57", "58"}
PASSIVE = {"02", "05", "14", "42", "59", "63", "77", "80", "82", "83", "96", "98"}

FIELDS = ("account", "name", "start_debit", "start_credit",
          "turn_debit", "turn_credit", "end_debit", "end_credit")

ACCOUNT_NAMES = {
    "01": "основные средства", "10": "материалы", "19": "НДС по приобретённым ценностям",
    "41": "товары", "43": "готовая продукция", "50": "касса", "51": "расчётные счета",
    "60": "расчёты с поставщиками", "62": "расчёты с покупателями", "66": "кредиты и займы",
    "68": "налоги и сборы", "69": "страховые взносы", "76": "прочие дебиторы и кредиторы",
    "80": "уставный капитал", "84": "прибыль", "90": "продажи", "91": "прочие доходы и расходы",
}


@dataclass
class Issue:
    kind: str
    account: str
    text: str
    amount: float

    def as_dict(self) -> dict:
        return {"kind": self.kind, "account": self.account,
                "text": self.text, "amount": round(self.amount, 2)}


@dataclass
class Report:
    path: Path
    period: str = ""
    rows: list[dict] = field(default_factory=list)
    problems: list[Issue] = field(default_factory=list)
    notes: list[Issue] = field(default_factory=list)
    totals: dict = field(default_factory=dict)


def decode(raw: bytes) -> str:
    for encoding in ("utf-8-sig", "utf-8", "cp1251"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("cp1251", errors="replace")


def read_rows(path: Path) -> list[list[str]]:
    if path.suffix.lower() in (".xlsx", ".xlsm"):
        sheet = load_workbook(path, data_only=True).worksheets[0]
        return [["" if cell is None else str(cell).strip() for cell in row]
                for row in sheet.iter_rows(values_only=True)]
    text = decode(path.read_bytes())
    sample = "\n".join(text.splitlines()[:10])
    delimiter = ";" if sample.count(";") >= sample.count(",") else ","
    return [row for row in csv.reader(io.StringIO(text), delimiter=delimiter)]


def parse_money(value) -> float:
    if value is None:
        return 0.0
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    text = str(value).replace("\xa0", " ").replace("\u202f", " ").strip()
    if not text or text in {"-", "—", "–"}:
        return 0.0
    negative = text.startswith("(") and text.endswith(")")
    text = text.strip("()").replace(" ", "").replace("₽", "").replace("руб.", "").replace("руб", "")
    text = text.replace(",", ".")
    text = re.sub(r"[^0-9.]", "", text)
    if not text:
        return 0.0
    try:
        result = float(text)
    except ValueError:
        return 0.0
    return -result if negative else result


def find_header(rows: list[list[str]]) -> tuple[int, int]:
    for index, row in enumerate(rows):
        cells = [str(cell).strip().lower() for cell in row]
        if any(cell in {"счёт", "счет", "сч."} for cell in cells):
            next_row = rows[index + 1] if index + 1 < len(rows) else []
            sub = {str(cell).strip().lower() for cell in next_row}
            if {"дебет", "кредит"} & sub:
                return index, index + 1
            return index, index
    raise SystemExit("не нашёл строку заголовка со счётом — проверьте структуру выгрузки")


def column_names(rows: list[list[str]], group_row: int, sub_row: int) -> dict[int, str]:
    group = rows[group_row]
    sub = rows[sub_row] if sub_row != group_row else [""] * len(group)
    names: dict[int, str] = {}
    current_group = ""
    for index, cell in enumerate(group):
        text = str(cell).strip()
        if text:
            current_group = text
        parts = [current_group] if sub_row != group_row else []
        if sub_row != group_row:
            parts.append(str(sub[index]).strip() if index < len(sub) else "")
        else:
            parts = [text]
        names[index] = " ".join(part for part in parts if part).lower()
    return names


def canonical(field_name: str) -> str | None:
    text = field_name.replace("ё", "е")
    if text in {"счет", "счёт", "сч."} or text.startswith("счет ") and "сальдо" not in text:
        return "account"
    if "наименован" in text:
        return "name"
    if "сальдо" in text and "начал" in text and "дебет" in text:
        return "start_debit"
    if "сальдо" in text and "начал" in text and "кредит" in text:
        return "start_credit"
    if "оборот" in text and "дебет" in text:
        return "turn_debit"
    if "оборот" in text and "кредит" in text:
        return "turn_credit"
    if "сальдо" in text and "конец" in text and "дебет" in text:
        return "end_debit"
    if "сальдо" in text and "конец" in text and "кредит" in text:
        return "end_credit"
    return None


def read_osv(path: Path) -> Report:
    rows = read_rows(path)
    group_row, sub_row = find_header(rows)
    names = column_names(rows, group_row, sub_row)
    mapping: dict[int, str] = {}
    for index, name in names.items():
        field_name = canonical(name)
        if field_name and field_name not in mapping.values():
            mapping[index] = field_name

    data: list[dict] = []
    for row in rows[sub_row + 1:]:
        first = str(row[0]).strip() if row else ""
        if not first:
            continue
        if first.lower().startswith(("итого", "всего")):
            break
        if not re.match(r"^\d{2}", first):
            continue
        item = {name: 0.0 for name in FIELDS}
        item["account"] = first.split(".")[0][:2]
        item["name"] = ""
        for index, field_name in mapping.items():
            value = row[index] if index < len(row) else ""
            if field_name == "name":
                item["name"] = str(value).strip()
            elif field_name != "account":
                item[field_name] = parse_money(value)
        data.append(item)

    report = Report(path=path, rows=data)
    for row in rows[: group_row + 1]:
        joined = " ".join(str(cell) for cell in row)
        if "ериод" in joined:
            report.period = joined.split(":", 1)[-1].strip() or joined
            break
    return report


def check(report: Report) -> Report:
    total = {name: round(sum(row[name] for row in report.rows), 2) for name in FIELDS[2:]}
    report.totals = total

    pairs = [("сальдо на начало", "start_debit", "start_credit"),
             ("обороты за период", "turn_debit", "turn_credit"),
             ("сальдо на конец", "end_debit", "end_credit")]
    for title, debit, credit in pairs:
        diff = round(total[debit] - total[credit], 2)
        if abs(diff) > TOLERANCE:
            side = "Дт больше Кт" if diff > 0 else "Кт больше Дт"
            report.problems.append(Issue(
                kind="balance", account="—", amount=abs(diff),
                text=(f"итоги по разделу «{title}» не сбалансированы: {side} на "
                      f"{money(abs(diff))} ₽. Обычно значит, что выгрузка неполная "
                      f"(фильтр по счёту или подразделению) либо в базе нарушена проводка.")))

    for row in report.rows:
        expected = round(row["start_debit"] - row["start_credit"]
                         + row["turn_debit"] - row["turn_credit"], 2)
        actual = round(row["end_debit"] - row["end_credit"], 2)
        diff = round(actual - expected, 2)
        if abs(diff) > TOLERANCE:
            report.problems.append(Issue(
                kind="arithmetic", account=row["account"], amount=abs(diff),
                text=(f"по счёту {row['account']} ({row['name']}) конечное сальдо не сходится "
                      f"с расчётом: в файле {money(actual)} ₽, по формуле "
                      f"{money(expected)} ₽, разрыв {money(diff)} ₽. Похоже на ручную правку "
                      f"ячейки или обрезанную выгрузку.")))
        elif abs(diff) > 0:
            report.notes.append(Issue(
                kind="rounding", account=row["account"], amount=abs(diff),
                text=(f"по счёту {row['account']} расхождение {money(abs(diff))} ₽ — "
                      f"в пределах допустимого, это округление.")))

    for row in report.rows:
        balance = round(row["end_debit"] - row["end_credit"], 2)
        if row["account"] in ACTIVE and balance < -TOLERANCE:
            report.problems.append(Issue(
                kind="red_balance", account=row["account"], amount=abs(balance),
                text=(f"по счёту {row['account']} ({row['name']}) кредитовое сальдо "
                      f"{money(abs(balance))} ₽ — «красное сальдо» на активном счёте. "
                      f"На активном счёте не может быть минуса: ищите непроведённый документ, "
                      f"ошибку последовательности проведения или пересорт.")))
        if row["account"] in PASSIVE and balance > TOLERANCE:
            report.problems.append(Issue(
                kind="red_balance", account=row["account"], amount=abs(balance),
                text=(f"по счёту {row['account']} ({row['name']}) дебетовое сальдо "
                      f"{money(abs(balance))} ₽ на пассивном счёте — проверьте проводки "
                      f"и закрытие периода.")))
    return report


def money(value: float) -> str:
    return f"{abs(value):,.2f}".replace(",", " ").replace(".", ",")


def render(report: Report) -> str:
    lines = [f"ОСВ: {report.path.name}"]
    if report.period:
        lines.append(f"Период: {report.period}")
    lines.append(f"Счетов: {len(report.rows)}")
    totals = report.totals
    lines.append("")
    lines.append("Итоги:")
    for title, debit, credit in (("сальдо на начало", "start_debit", "start_credit"),
                                 ("обороты за период", "turn_debit", "turn_credit"),
                                 ("сальдо на конец", "end_debit", "end_credit")):
        diff = totals[debit] - totals[credit]
        mark = "ок" if abs(diff) <= TOLERANCE else f"расхождение {money(diff)} ₽"
        lines.append(f"  {title}: Дт {money(totals[debit])} / Кт {money(totals[credit])} — {mark}")
    lines.append("")
    lines.append(f"Проблем: {len(report.problems)}")
    for number, issue in enumerate(report.problems, start=1):
        lines.append(f"{number}. {issue.text}")
    if report.notes:
        lines.append("")
        lines.append(f"Замечания ({len(report.notes)}):")
        for issue in report.notes:
            lines.append(f"- {issue.text}")
    return "\n".join(lines)


def write_report(path: Path, report: Report) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Проблемы"
    headers = ["Вид", "Счёт", "Сумма, ₽", "Описание"]
    for index, name in enumerate(headers, start=1):
        sheet.cell(row=1, column=index, value=name)
    row_index = 2
    for issue in report.problems + report.notes:
        sheet.cell(row=row_index, column=1, value=issue.kind)
        sheet.cell(row=row_index, column=2, value=issue.account)
        sheet.cell(row=row_index, column=3, value=round(issue.amount, 2))
        sheet.cell(row=row_index, column=4, value=issue.text)
        row_index += 1
    for letter, width in (("A", 14), ("B", 8), ("C", 14), ("D", 110)):
        sheet.column_dimensions[letter].width = width

    data_sheet = workbook.create_sheet("Данные")
    for index, name in enumerate(FIELDS, start=1):
        data_sheet.cell(row=1, column=index, value=name)
    for row_index, item in enumerate(report.rows, start=2):
        for column, name in enumerate(FIELDS, start=1):
            data_sheet.cell(row=row_index, column=column, value=item[name])
    data_sheet.freeze_panes = "A2"
    for letter in "ABCDEFGH":
        data_sheet.column_dimensions[letter].width = 16
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)


def write_csv(path: Path, report: Report) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle, delimiter=";")
        writer.writerow(FIELDS)
        for item in report.rows:
            writer.writerow([item[name] for name in FIELDS])


def main() -> None:
    parser = argparse.ArgumentParser(description="Проверяет оборотно-сальдовую ведомость из 1С")
    parser.add_argument("file", type=Path)
    parser.add_argument("--json", action="store_true", help="вывести результат в JSON")
    parser.add_argument("--csv", type=Path, help="сохранить нормализованные данные")
    parser.add_argument("--report", type=Path, default=Path("outputs/report.xlsx"),
                        help="куда сохранить xlsx с проблемами")
    args = parser.parse_args()

    report = check(read_osv(args.file))
    write_report(args.report, report)
    if args.csv:
        write_csv(args.csv, report)

    if args.json:
        payload = {
            "file": str(args.file),
            "period": report.period,
            "accounts": len(report.rows),
            "totals": {name: round(value, 2) for name, value in report.totals.items()},
            "problems": [issue.as_dict() for issue in report.problems],
            "notes": [issue.as_dict() for issue in report.notes],
        }
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    else:
        print(render(report))
    print(f"\nотчёт: {args.report}")


if __name__ == "__main__":
    main()
