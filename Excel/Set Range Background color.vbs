Set objXLApp = CreateObject("Excel.Application")
objXLApp.Visible = false
Set objWorkbook = objXLApp.Workbooks.Open("C:\Users\berka\Desktop\VbsTestler\Excel.xlsx")
Set oSheet = objWorkbook.Worksheets(1) 
 
oSheet.Range("A1:H7").Interior.Color = RGB(255, 0, 0) 'Red
 
objWorkbook.save
objWorkbook.close false
objXLApp.Quit

Set objWorkbook = nothing
Set objXLApp = nothing