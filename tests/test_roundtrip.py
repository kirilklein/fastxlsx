import datetime
import os
import tempfile
import unittest

import fastxlsx
from python_calamine import CalamineWorkbook


class RoundTripTest(unittest.TestCase):
    def _read_back(self, path, sheet_index=0):
        return CalamineWorkbook.from_path(path).get_sheet_by_index(sheet_index).to_python()

    def test_mixed_types_roundtrip(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "out.xlsx")
            workbook = fastxlsx.Workbook()
            sheet = workbook.add_sheet("Patients")
            sheet.append(["Patient_ID", "Score", "Eligible", "Note"])
            sheet.append([123, 4.5, True, "NSQIP"])
            sheet.append([124, -1.0, False, None])
            workbook.save(path)

            rows = self._read_back(path)
            self.assertEqual(rows[0], ["Patient_ID", "Score", "Eligible", "Note"])
            self.assertEqual(rows[1], [123.0, 4.5, True, "NSQIP"])
            # None -> blank cell; calamine returns "" for the trailing empty cell.
            self.assertEqual(rows[2][:3], [124.0, -1.0, False])

    def test_dates_roundtrip(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "out.xlsx")
            workbook = fastxlsx.Workbook()
            sheet = workbook.add_sheet("Dates")
            sheet.append([datetime.date(2024, 1, 15),
                          datetime.datetime(2024, 1, 15, 13, 30, 45)])
            workbook.save(path)

            rows = self._read_back(path)
            self.assertEqual(rows[0][0], datetime.date(2024, 1, 15))
            self.assertEqual(rows[0][1], datetime.datetime(2024, 1, 15, 13, 30, 45))

    def test_timezone_aware_datetime_keeps_wallclock(self):
        # tzinfo is dropped, but the wall-clock time (incl. the time of day) is kept.
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "out.xlsx")
            workbook = fastxlsx.Workbook()
            sheet = workbook.add_sheet("TZ")
            aware = datetime.datetime(2024, 1, 15, 13, 30, 45, tzinfo=datetime.timezone.utc)
            sheet.append([aware])
            workbook.save(path)

            self.assertEqual(self._read_back(path)[0][0],
                             datetime.datetime(2024, 1, 15, 13, 30, 45))

    def test_multiple_sheets(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "out.xlsx")
            workbook = fastxlsx.Workbook()
            first = workbook.add_sheet("First")
            second = workbook.add_sheet("Second")
            first.append(["a", 1])
            second.append(["b", 2])
            workbook.save(path)

            self.assertEqual(self._read_back(path, 0)[0], ["a", 1.0])
            self.assertEqual(self._read_back(path, 1)[0], ["b", 2.0])


if __name__ == "__main__":
    unittest.main()
