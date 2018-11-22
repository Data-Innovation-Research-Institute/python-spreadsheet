import unittest

from spreadsheet import Spreadsheet


class SpreadsheetTests(unittest.TestCase):

    def test_letter(self):
        spreadsheet = Spreadsheet()
        self.assertEqual(spreadsheet.letter(0), 'A')
        self.assertEqual(spreadsheet.letter(25), 'Z')
        self.assertEqual(spreadsheet.letter(26), 'AA')
        self.assertEqual(spreadsheet.letter(51), 'ZZ')
        self.assertEqual(spreadsheet.letter(52), 'AAA')
        self.assertEqual(spreadsheet.letter(77), 'ZZZ')

    def test_get_sheet(self):
        spreadsheet = Spreadsheet()
        self.assertFalse('Stats' in spreadsheet.workbook.sheetnames)
        sheet = spreadsheet.get_sheet('Stats')
        self.assertEqual(sheet.title, 'Stats')

    def test_select_sheet(self):
        spreadsheet = Spreadsheet()
        self.assertNotEqual(spreadsheet.current_sheet().title, 'Stats')
        spreadsheet.select_sheet('Stats')
        self.assertEqual(spreadsheet.current_sheet().title, 'Stats')

    def test_advance_column(self):
        spreadsheet = Spreadsheet()
        self.assertEqual(spreadsheet.current_cell(), 'A1')
        spreadsheet.advance_column()
        self.assertEqual(spreadsheet.current_cell(), 'B1')
        spreadsheet.advance_column()
        self.assertEqual(spreadsheet.current_cell(), 'C1')

    def test_advance_row(self):
        spreadsheet = Spreadsheet()
        spreadsheet.advance_column()
        spreadsheet.advance_column()
        spreadsheet.advance_column()
        self.assertEqual(spreadsheet.current_cell(), 'D1')
        spreadsheet.advance_row()
        self.assertEqual(spreadsheet.current_cell(), 'A2')

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
