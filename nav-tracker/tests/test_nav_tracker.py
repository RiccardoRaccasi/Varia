#!/usr/bin/env python3
"""Offline tests for nav_tracker — run with `python tests/test_nav_tracker.py`.

The live fund page cannot be reached from CI, so parsing is exercised against
saved fixtures covering the layouts a fund detail page realistically uses.
"""

import datetime as dt
import os
import shutil
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
FIXTURES = os.path.join(HERE, "fixtures")
sys.path.insert(0, os.path.dirname(HERE))

import nav_tracker as nt  # noqa: E402


def fixture(name):
    with open(os.path.join(FIXTURES, name), encoding="utf-8") as handle:
        return handle.read()


class NumberParsing(unittest.TestCase):
    def test_formats(self):
        cases = {
            "231.845": 231.845, "1'234.56": 1234.56, "1’234.56": 1234.56,
            "1,234.56": 1234.56, "1.234,56": 1234.56, "1 234,56": 1234.56,
            "USD 99.74": 99.74, "231.845 USD": 231.845, "-3.5": -3.5,
            "112'450'000": 112450000.0, "0,5": 0.5,
        }
        for text, expected in cases.items():
            self.assertAlmostEqual(nt.parse_number(text), expected, places=6, msg=text)

    def test_rejects_non_numbers(self):
        for text in ("09.09.2026", "", "n/a", "CH1234846777", None):
            self.assertIsNone(nt.parse_number(text), msg=repr(text))

    def test_glued_markup_is_refused_rather_than_misread(self):
        # "USD231.845" with the space eaten by markup must not read as 845.
        self.assertIsNone(nt.parse_number("USD231.845"))
        self.assertIsNone(nt.parse_number("ISIN CH1234846777"))

    def test_dates(self):
        self.assertEqual(nt.parse_date("09.09.2026"), dt.date(2026, 9, 9))
        self.assertEqual(nt.parse_date("09/09/2026"), dt.date(2026, 9, 9))
        self.assertEqual(nt.parse_date("2026-09-09"), dt.date(2026, 9, 9))
        self.assertEqual(nt.parse_date("as of 16.06.2023"), dt.date(2023, 6, 16))
        self.assertIsNone(nt.parse_date("231.845"))


class Extraction(unittest.TestCase):
    def assert_price(self, html_name, value, when=dt.date(2026, 9, 9)):
        point = nt.fetch_current_price(html=fixture(html_name))
        self.assertAlmostEqual(point.value, value, places=6)
        self.assertEqual(point.date, when)
        self.assertFalse(point.date_is_assumed)
        return point

    def test_column_price_table_takes_newest_row(self):
        point = self.assert_price("price_table.html", 231.845)
        self.assertIn("net asset value", point.label.lower())

    def test_label_value_table(self):
        self.assert_price("label_value.html", 231.845)

    def test_definition_list(self):
        self.assert_price("definition_list.html", 231.845)

    def test_german_swiss_thousands(self):
        self.assert_price("german_swiss_format.html", 1231.845)

    def test_current_price_wording(self):
        point = self.assert_price("current_price_only.html", 231.845)
        self.assertIn("current price", point.label.lower())

    def test_nav_beats_issue_and_redemption_price(self):
        best = nt.extract_candidates(fixture("price_table.html"))[0]
        self.assertEqual(best.field, "nav")
        self.assertAlmostEqual(best.value, 231.845, places=6)

    def test_volume_and_unit_counts_are_never_prices(self):
        for cand in nt.extract_candidates(fixture("label_value.html")):
            self.assertNotIn(cand.value, (112450000.0, 485000.0))

    def test_no_spurious_candidates_from_flattened_text(self):
        cands = nt.extract_candidates(fixture("label_value.html"))
        self.assertEqual(len(cands), 1)
        self.assertAlmostEqual(cands[0].value, 231.845, places=6)

    def test_script_contents_ignored(self):
        self.assertNotIn(999.99, [c.value for c in
                                  nt.extract_candidates(fixture("price_table.html"))])

    def test_missing_date_refuses_by_default(self):
        with self.assertRaises(nt.ParseError):
            nt.fetch_current_price(html=fixture("no_date.html"))

    def test_missing_date_with_assume_today(self):
        point = nt.fetch_current_price(html=fixture("no_date.html"), assume_today=True)
        self.assertTrue(point.date_is_assumed)
        self.assertEqual(point.date, nt.local_today())

    def test_unparseable_page_raises(self):
        with self.assertRaises(nt.ParseError):
            nt.fetch_current_price(html="<html><body><p>Down for maintenance</p></body></html>")

    def test_label_regex_pins_a_field(self):
        point = nt.fetch_current_price(html=fixture("price_table.html"),
                                       label_regex=r"redemption price")
        self.assertAlmostEqual(point.value, 231.85, places=6)

    def test_csv_response_is_understood(self):
        point = nt.fetch_current_price(html=fixture("export.csv"))
        self.assertAlmostEqual(point.value, 231.845, places=6)
        self.assertEqual(point.date, dt.date(2026, 9, 9))


class Workbook(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp()
        self.book = os.path.join(self.dir, "nav.xlsx")
        nt.import_csv(os.path.join(FIXTURES, "export.csv"), self.book)

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def test_import_round_trips(self):
        rows = nt.read_history(self.book)
        self.assertEqual([r.date for r in rows],
                         [dt.date(2026, 9, 8), dt.date(2026, 9, 9)])
        self.assertAlmostEqual(rows[-1].value, 231.845, places=6)
        self.assertEqual(rows[-1].currency, "USD")

    def test_update_appends_a_new_day(self):
        html = fixture("price_table.html").replace("09.09.2026", "10.09.2026") \
                                          .replace("231.845", "232.5")
        result = nt.update(html=html, workbook=self.book)
        self.assertEqual(result.status, "added")
        rows = nt.read_history(self.book)
        self.assertEqual(len(rows), 3)
        self.assertEqual(rows[-1].date, dt.date(2026, 9, 10))
        self.assertAlmostEqual(rows[-1].value, 232.5, places=6)
        self.assertEqual(rows[-1].source, "swissfunddata")

    def test_rerunning_the_same_day_changes_nothing(self):
        first = nt.update(html=fixture("price_table.html"), workbook=self.book)
        self.assertEqual(first.status, "unchanged")
        before = os.path.getmtime(self.book)
        again = nt.update(html=fixture("price_table.html"), workbook=self.book)
        self.assertEqual(again.status, "unchanged")
        self.assertEqual(os.path.getmtime(self.book), before)
        self.assertEqual(len(nt.read_history(self.book)), 2)

    def test_a_restated_value_is_corrected_in_place(self):
        html = fixture("price_table.html").replace("<td>231.845</td>", "<td>231.900</td>")
        result = nt.update(html=html, workbook=self.book)
        self.assertEqual(result.status, "updated")
        rows = nt.read_history(self.book)
        self.assertEqual(len(rows), 2)
        self.assertAlmostEqual(rows[-1].value, 231.9, places=6)

    def test_rows_stay_sorted_when_an_older_day_arrives(self):
        html = fixture("label_value.html").replace("09.09.2026", "05.09.2026") \
                                          .replace("231.845", "216.0")
        nt.update(html=html, workbook=self.book)
        rows = nt.read_history(self.book)
        self.assertEqual([r.date for r in rows], sorted(r.date for r in rows))
        self.assertEqual(rows[0].date, dt.date(2026, 9, 5))

    def test_implausible_move_is_rejected(self):
        html = fixture("price_table.html").replace("09.09.2026", "10.09.2026") \
                                          .replace("<td>231.845</td>", "<td>23184.5</td>")
        with self.assertRaises(nt.SanityError):
            nt.update(html=html, workbook=self.book)
        self.assertEqual(len(nt.read_history(self.book)), 2)   # file untouched

    def test_force_accepts_a_large_move(self):
        html = fixture("price_table.html").replace("09.09.2026", "10.09.2026") \
                                          .replace("<td>231.845</td>", "<td>23184.5</td>")
        self.assertEqual(nt.update(html=html, workbook=self.book, force=True).status, "added")

    def test_future_dated_price_is_rejected(self):
        ahead = (nt.local_today() + dt.timedelta(days=30)).strftime("%d.%m.%Y")
        html = fixture("price_table.html").replace("09.09.2026", ahead)
        with self.assertRaises(nt.SanityError):
            nt.update(html=html, workbook=self.book)

    def test_stale_page_is_not_appended_under_an_assumed_date(self):
        html = fixture("no_date.html")          # value 231.845 == last stored
        result = nt.update(html=html, workbook=self.book, assume_today=True)
        self.assertEqual(result.status, "skipped-stale")
        self.assertEqual(len(nt.read_history(self.book)), 2)

    def test_dry_run_writes_nothing(self):
        html = fixture("price_table.html").replace("09.09.2026", "10.09.2026") \
                                          .replace("231.845", "232.5")
        before = os.path.getmtime(self.book)
        result = nt.update(html=html, workbook=self.book, dry_run=True)
        self.assertEqual(result.status, "added")
        self.assertEqual(os.path.getmtime(self.book), before)
        self.assertEqual(len(nt.read_history(self.book)), 2)

    def test_missing_workbook_is_a_clear_error(self):
        with self.assertRaises(nt.NavTrackerError):
            nt.read_history(os.path.join(self.dir, "absent.xlsx"))


class CommandLine(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp()
        self.book = os.path.join(self.dir, "nav.xlsx")

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def test_import_probe_update_show(self):
        self.assertEqual(nt.main(["import-csv", os.path.join(FIXTURES, "export.csv"),
                                  "-w", self.book]), 0)
        self.assertEqual(nt.main(["probe", "--html",
                                  os.path.join(FIXTURES, "price_table.html")]), 0)
        self.assertEqual(nt.main(["update", "-w", self.book, "--html",
                                  os.path.join(FIXTURES, "price_table.html"),
                                  "--json"]), 0)
        self.assertEqual(nt.main(["show", "-w", self.book, "-n", "2"]), 0)

    def test_parse_failure_exits_2(self):
        empty = os.path.join(self.dir, "empty.html")
        with open(empty, "w") as handle:
            handle.write("<html><body>nothing here</body></html>")
        self.assertEqual(nt.main(["fetch", "--html", empty]), 2)

    def test_sanity_failure_exits_3(self):
        nt.main(["import-csv", os.path.join(FIXTURES, "export.csv"), "-w", self.book])
        bad = os.path.join(self.dir, "bad.html")
        with open(bad, "w", encoding="utf-8") as handle:
            handle.write(fixture("price_table.html")
                         .replace("09.09.2026", "10.09.2026")
                         .replace("<td>231.845</td>", "<td>23184.5</td>"))
        self.assertEqual(nt.main(["update", "-w", self.book, "--html", bad]), 3)


if __name__ == "__main__":
    unittest.main(verbosity=2)
