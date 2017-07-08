import win32com.client as win32
from win32com.client import constants as c
__version__ = 1.0
# http://msdn.microsoft.com/en-us/library/office/ff194068(v=office.14).aspx


class Excel:
    def __init__(self):
        self.excel = win32.gencache.EnsureDispatch('Excel.Application')

    def open(self, file, visible=False, read_only=False):
        # takes in a filename, workSheet name and returns an excel & worksheet object
        self.excel.Visible = visible
        self.wb = self.excel.Workbooks.Open(Filename=file, ReadOnly=read_only)
        # self.wb = self.excel.ActiveWorkbook
        return

    def append(self, worksheet_name, items_to_add, add_to_column="A", check_entire_row_is_blank=True,
               goto_bottom_of_column=''):
        # Appends array to the end of an excel file. It checks the row is entirely blank before writing.
        # WARNING: when sending in a date, change to mm/dd/yy due to "feature" in excel com
        # excel.Visible = False
        ws = self.excel.Worksheets(worksheet_name)
        if goto_bottom_of_column:  # allows you to check the end of one column, but write to another
            rowToWriteIn = self.last_empty_row_in_column(goto_bottom_of_column, check_entire_row_is_blank)
        else:
            rowToWriteIn = self.last_empty_row_in_column(add_to_column, check_entire_row_is_blank)
        range_to_write_to = add_to_column + str(rowToWriteIn) + ":" \
                            + chr(ord(add_to_column) + len(items_to_add)-1) + str(rowToWriteIn)  # chr(ord("A")+1)="B"
        ws.Range(range_to_write_to).Value2 = [i for i in items_to_add]

    def make_work_sheet_active(self, worksheet_name):
        # nice. Makes worksheet active. Returns a worksheet or false
        return self.excel.Worksheets(worksheet_name).Activate

    def close(self, save=True):
        self.wb.Close(SaveChanges=save)

    def last_empty_row_in_column(self, column, check_entire_row_is_blank=False):
        row = self.excel.Range(column + str(self.excel.ActiveSheet.Rows.Count)).End(c.xlUp).Row + 1 #find last row in column with data, go down 1
        if check_entire_row_is_blank:
            while self.excel.WorksheetFunction.CountA(self.excel.Range(str(row) + ":" + str(row))) != 0:  #row has an entry
                row += 1  # go down a row
        return row

    def last_row_used_range(self):
        return self.excel.Cells.SpecialCells(c.xlCellTypeLastCell).Row

    def append_to_open_excel(self, ws, itemsToAdd, addToColumn="A", check_entire_row_is_blank=True, goto_bottom_of_column=''):
        # Appends array to the end of an excel file. It checks the row is entirely blank before writing.
        # WARNING: when sending in a date, change to mm/dd/yy due to "feature" in excel com
        try:
            if goto_bottom_of_column:
                # Allows you to check the end of one column, but write to another
                rowToWriteIn = self.last_empty_row_in_column(goto_bottom_of_column, check_entire_row_is_blank)
            else:
                rowToWriteIn = self.last_empty_row_in_column(addToColumn, check_entire_row_is_blank)
            range_to_write_to = addToColumn + str(rowToWriteIn) + ":" \
                                + chr(ord(addToColumn) + len(itemsToAdd)-1) + str(rowToWriteIn)  # chr(ord("A")+1)="B"
            ws.Range(range_to_write_to).Value2 = [i for i in itemsToAdd]
            return True
        except:
            return False

    def auto_fill(self, ws, from_range, to_range):
        # e.g. fromRange="A1:A3"
        # http://msdn.microsoft.com/en-us/library/office/ff195345(v=office.14).aspx
        source_range = ws.Range(from_range)
        fill_range = ws.Range(to_range)
        source_range.auto_fill(fill_range)

    def auto_fill_down_from_end(self, ws, column, amount):
        row_end_of_data = self.last_empty_row_in_column(column) - 1
        range_from = column + str(row_end_of_data)
        range_to = range_from + ':' + column + str(row_end_of_data+amount)
        self.auto_fill(ws, range_from, range_to)

    def last_cell_value_in_column(self, ws, column):
        # find last cell value
        return ws.Range(column + str(ws.Rows.Count)).End(c.xlUp).Value2

    def last_row_in_column(self, ws, column):
        # find last row number
        return ws.Range(column + str(ws.Rows.Count)).End(c.xlUp).Row

    def save_as(self, filename, save_as_default=False):
        wb = self.excel.ActiveWorkbook
        if save_as_default:
            self.excel.DisplayAlerts = False
            wb.SaveAs(filename, FileFormat=c.xlWorkbookDefault, ConflictResolution=c.xlLocalSessionChanges)
            self.excel.DisplayAlerts = True
        else:
            self.excel.DisplayAlerts = False
            wb.SaveAs(filename, ConflictResolution=c.xlLocalSessionChanges)
            self.excel.DisplayAlerts = True
