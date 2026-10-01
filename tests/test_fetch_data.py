import unittest
from datetime import date

from fetch_data import parse_list_page, select_active_row


class PeriodSelectionTests(unittest.TestCase):
    def test_future_first_row_does_not_replace_active_period(self):
        html = '''
        <tbody class="ecl-table__body">
          <tr class="ecl-table__row"><td>099836</td><td>Korea</td>
            <td>01-10-2026</td><td>31-12-2026</td><td>100 Kilogram</td>
            <td><a href="quota_tariff_details.jsp?Lang=en&StartDate=2026-10-01&Code=099836">More info</a></td></tr>
          <tr class="ecl-table__row"><td>099836</td><td>Korea</td>
            <td>01-07-2026</td><td>30-09-2026</td><td>20 Kilogram</td>
            <td><a href="quota_tariff_details.jsp?Lang=en&StartDate=2026-07-01&Code=099836">More info</a></td></tr>
        </tbody>'''
        rows = parse_list_page("099836", html)
        self.assertEqual(len(rows), 2)
        self.assertEqual(select_active_row(rows, date(2026, 9, 30))["detail_start_date"], "2026-07-01")
        self.assertEqual(select_active_row(rows, date(2026, 10, 1))["detail_start_date"], "2026-10-01")


if __name__ == "__main__":
    unittest.main()
