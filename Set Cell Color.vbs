Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = false
objExcel.displayalerts=false

Set objWorkbook = objExcel.Workbooks.open("C:\RobustaJiraProjects\IP-598 Excel Get Cell Color VB\VBA Get Cell Color Test.xlsx")
Set objWorksheet1 = objWorkbook.Worksheets("Sheet1")

Set objRange = objWorksheet1.UsedRange


For intRowCounter = 1 to objWorksheet1.usedRange.Rows.Count 

objWorksheet1.Range("B" & intRowCounter) = objWorksheet1.Cells(intRowCounter,1).Interior.ColorIndex 

Next

objWorkbook.save
objWorkbook.close
objExcel.Quit()
Set objWorksheet1 = Nothing
Set objWorkbook = Nothing
Set ObjExcel = Nothing

WScript.Echo "Finished."
WScript.Quit