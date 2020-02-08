import win32com.client
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = 1
k = self.excel.Workbooks.Add()
v = k.Worksheets(1)
v.Cells(1,1).Value = "Hello, World"