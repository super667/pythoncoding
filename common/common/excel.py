import os

class CExcel_Win32:
    def __init__(self, filepath):
        self.filepath = os.path.abspath(filepath)

    def open(self):
        import win32com.client
        self.xlApp = win32com.client.Dispatch('Excel.Application')
        self.xlBook = self.xlApp.Workbooks.Open(self.filepath)

    def save(self):
        self.xlBook.Save()

    def close(self):
        self.xlBook.Close(False)

    def getSheetCount(self):
        return self.xlBook.Sheets.Count

    def getSheetNameList(self):
        sheet_name_list = []
        for index in range(0, self.xlBook.Sheets.Count):
            sheet = self.xlBook.Worksheets[index]
            sheet_name_list.append(sheet.Name)
        return sheet_name_list

    def getRowCount(self, sheet_name):
        sheet = self.xlBook.Worksheets(sheet_name)
        return sheet.UsedRange.Rows.Count

    def getColumnCount(self, sheet_name):
        sheet = self.xlBook.Worksheets(sheet_name)
        return sheet.UsedRange.Columns.Count

    def getDimensions(self, sheet_name):
        row = self.getRowCount()
        col = self.getColumnCount()
        return (row, col, None, None, None)

    def getCellValue(self, sheet_name, row, col):
        sheet = self.xlBook.Worksheets(sheet_name)
        return sheet.Cells(row, col).Value

    def setCellValue(self, sheet_name, row, col, value):
        sheet = self.xlBook.Worksheets(sheet_name)
        sheet.Cells(row, col).Value = value

    def setCellRed(self, sheet_name, row, col):
        sheet = self.xlBook.Worksheets(sheet_name)
        sheet.Cells(row, col).Interior.ColorIndex = 3
        
    def deleteRowValue(self, sheet_name, row):
        sheet = self.xlBook.Worksheets(sheet_name)
        sheet.Rows(row).Delete()
        
