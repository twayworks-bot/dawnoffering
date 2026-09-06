import unittest
from datetime import datetime, date
import pandas as pd
from app import get_target_week_dates, convert_google_drive_url, normalize_date, load_env

class TestDawnOffering(unittest.TestCase):
    
    def test_get_target_week_dates_weekday(self):
        # 2026-09-09 (Wednesday) -> should return Mon 9/7 to Fri 9/11
        ref = date(2026, 9, 9)
        dates = get_target_week_dates(ref)
        self.assertEqual(len(dates), 5)
        self.assertEqual(dates[0], date(2026, 9, 7))
        self.assertEqual(dates[4], date(2026, 9, 11))
        
    def test_get_target_week_dates_weekend_sunday(self):
        # 2026-09-06 (Sunday) -> should return next week Mon 9/7 to Fri 9/11
        ref = date(2026, 9, 6)
        dates = get_target_week_dates(ref)
        self.assertEqual(len(dates), 5)
        self.assertEqual(dates[0], date(2026, 9, 7))
        self.assertEqual(dates[4], date(2026, 9, 11))
        
    def test_get_target_week_dates_weekend_saturday(self):
        # 2026-09-05 (Saturday) -> should return next week Mon 9/7 to Fri 9/11
        ref = date(2026, 9, 5)
        dates = get_target_week_dates(ref)
        self.assertEqual(len(dates), 5)
        self.assertEqual(dates[0], date(2026, 9, 7))
        self.assertEqual(dates[4], date(2026, 9, 11))

    def test_convert_google_drive_url_sheets(self):
        url = "https://docs.google.com/spreadsheets/d/1RWKAewSqTB6wVCw_-R86NjoGszd0Y1Sy/edit?usp=sharing"
        converted, utype = convert_google_drive_url(url)
        self.assertEqual(utype, "excel")
        self.assertIn("export?format=xlsx", converted)

    def test_convert_google_drive_url_file(self):
        url = "https://drive.google.com/file/d/1RWKAewSqTB6wVCw_-R86NjoGszd0Y1Sy/view?usp=sharing"
        converted, utype = convert_google_drive_url(url)
        self.assertEqual(utype, "excel")
        self.assertIn("export=download", converted)

    def test_normalize_date_formats(self):
        # Format 1: String date like '2026. 09. 07. 월요일'
        d1 = normalize_date("2026. 09. 07. 월요일", 2026)
        self.assertEqual(d1, date(2026, 9, 7))

        # Format 2: '9/7' (assuming 2026)
        d2 = normalize_date("9/7", 2026)
        self.assertEqual(d2, date(2026, 9, 7))

        # Format 3: Timestamp/Datetime object
        d3 = normalize_date(pd.Timestamp("2026-09-07"), 2026)
        self.assertEqual(d3, date(2026, 9, 7))

if __name__ == '__main__':
    unittest.main()
