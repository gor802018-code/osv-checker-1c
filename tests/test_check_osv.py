import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from src.check_osv import (  # noqa: E402
    canonical,
    check,
    money,
    parse_money,
    read_osv,
    render,
)

ROOT = Path(__file__).resolve().parents[1]
XLSX = ROOT / "data" / "raw" / "osv_2024_q3.xlsx"
CSV = ROOT / "data" / "raw" / "osv_2024_q3_cp1251.csv"


class ParseMoneyTest(unittest.TestCase):
    def test_spaces_and_comma(self):
        self.assertEqual(parse_money("1 200 000,00"), 1200000.0)

    def test_nbsp(self):
        self.assertEqual(parse_money("96\u00a0000,30"), 96000.3)

    def test_parentheses_are_negative(self):
        self.assertEqual(parse_money("(500,00)"), -500.0)

    def test_empty_and_dash(self):
        self.assertEqual(parse_money(""), 0.0)
        self.assertEqual(parse_money("—"), 0.0)
        self.assertEqual(parse_money(None), 0.0)

    def test_number(self):
        self.assertEqual(parse_money(150000), 150000.0)


class CanonicalTest(unittest.TestCase):
    def test_saldo_start_debit(self):
        self.assertEqual(canonical("сальдо на начало дебет"), "start_debit")

    def test_turnover_credit(self):
        self.assertEqual(canonical("обороты за период кредит"), "turn_credit")

    def test_saldo_end_credit(self):
        self.assertEqual(canonical("сальдо на конец кредит"), "end_credit")

    def test_account(self):
        self.assertEqual(canonical("счёт"), "account")

    def test_unknown(self):
        self.assertIsNone(canonical("примечание"))


class OsvXlsxTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.report = check(read_osv(XLSX))

    def test_accounts_parsed(self):
        self.assertEqual(len(self.report.rows), 18)

    def test_period_detected(self):
        self.assertIn("01.07.2024", self.report.period)

    def test_turnovers_balanced(self):
        totals = self.report.totals
        self.assertAlmostEqual(totals["turn_debit"], totals["turn_credit"], places=2)

    def test_start_saldo_imbalance(self):
        totals = self.report.totals
        self.assertAlmostEqual(totals["start_debit"] - totals["start_credit"], 12480.0, places=2)

    def test_arithmetic_issue_on_62(self):
        found = [i for i in self.report.problems if i.kind == "arithmetic"]
        self.assertEqual(len(found), 1)
        self.assertEqual(found[0].account, "62")
        self.assertAlmostEqual(found[0].amount, 10000.0, places=2)

    def test_red_balance_on_50(self):
        found = [i for i in self.report.problems if i.kind == "red_balance"]
        self.assertEqual(len(found), 1)
        self.assertEqual(found[0].account, "50")
        self.assertAlmostEqual(found[0].amount, 5000.0, places=2)

    def test_rounding_is_note_not_problem(self):
        self.assertEqual(len(self.report.notes), 1)
        self.assertEqual(self.report.notes[0].account, "19")

    def test_render_has_period_and_problems(self):
        text = render(self.report)
        self.assertIn("Проблем: 4", text)
        self.assertIn("01.07.2024", text)


class OsvCsvTest(unittest.TestCase):
    def test_same_result_as_xlsx(self):
        from_csv = check(read_osv(CSV))
        from_xlsx = check(read_osv(XLSX))
        self.assertEqual(len(from_csv.rows), len(from_xlsx.rows))
        self.assertAlmostEqual(from_csv.totals["start_debit"], from_xlsx.totals["start_debit"], places=2)
        self.assertEqual(len(from_csv.problems), len(from_xlsx.problems))


class MoneyFormatTest(unittest.TestCase):
    def test_russian_style(self):
        self.assertEqual(money(12480), "12 480,00")
        self.assertEqual(money(0.3), "0,30")


if __name__ == "__main__":
    unittest.main()
