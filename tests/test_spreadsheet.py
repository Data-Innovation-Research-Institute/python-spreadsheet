import unittest

from spreadsheet import Spreadsheet


class SpreadsheetTests(unittest.TestCase):

    def test_cells(self):
        spreadsheet = Spreadsheet()
        self.assertEqual(spreadsheet.current_cell(), 'A1')
        spreadsheet.set_value('1')
        self.assertEqual(spreadsheet.current_cell(), 'B1')
        spreadsheet.set_value('1', advance_row=True)
        self.assertEqual(spreadsheet.current_cell(), 'A2')
        spreadsheet.set_value('1', advance_by_rows=3)
        self.assertEqual(spreadsheet.current_cell(), 'A5')
        spreadsheet.set_values(['1'])
        self.assertEqual(spreadsheet.current_cell(), 'A6')


if __name__ == '__main__':
    unittest.main()
