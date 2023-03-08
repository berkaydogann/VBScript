Set objXLApp = CreateObject("Excel.Application")
Set objXLWb = objXLApp.Workbooks.Open("C:\Users\berka\Desktop\VbsTestler\Excel.xls")
Set oSheet = objXLWb.Worksheets(1)
objXLWb.SaveAs "C:\Users\berka\Desktop\VbsTestler\Excel.xlsx", 51
objXLApp.Quit