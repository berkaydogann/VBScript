Set objXLApp = CreateObject("Excel.Application")
objXLApp.Visible = false
Set objWorkbook = objXLApp.Workbooks.Open("C:\RobustaJiraProjects\IP-705\Test 1.xlsx")
Set oSheet = objWorkbook.Worksheets(1) 
 
oSheet.Range("A1:H7").Interior.Color = RGB(255, 0, 0)
oSheet.Range("E2:E14").NumberFormat = "#,##0.00" 


objWorkbook.save
objWorkbook.close false
objXLApp.Quit

Set objWorkbook = nothing
Set objXLApp = nothing