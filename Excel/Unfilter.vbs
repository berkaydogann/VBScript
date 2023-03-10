Set objXLApp = CreateObject("Excel.Application")
objXLApp.Visible = False
objXLApp.DisplayAlerts = False
Set objXLWb = objXLApp.Workbooks.Open("C:\Users\berka\Desktop\VbsTestler\Test.xlsx")

objXLWb.Worksheets(1).AutoFilterMode  = False

objXLWb.Save
objXLWb.Close
objXLApp.Quit




