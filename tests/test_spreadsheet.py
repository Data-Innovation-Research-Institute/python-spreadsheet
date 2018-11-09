import unittest

from spreadsheet import Spreadsheet


class SpreadsheetTests(unittest.TestCase):

    def test_cells(self):
        spreadsheet = Spreadsheet()
        self.assertEqual(spreadsheet.current_cell(), 'A1')
        spreadsheet.set_value('1')
        self.assertEqual(spreadsheet.current_cell(), 'B1')


if __name__ == '__main__':
    unittest.main()
