Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = false
objExcel.displayalerts=false

Set objWorkbook = objExcel.Workbooks.open("C:\Users\berka\Desktop\VbsTestler\Excel.xlsx")
Set objWorksheet1 = objWorkbook.Worksheets("Sheet1")

objWorksheet1.Name = "SheetNewOne"

objWorkbook.save
objWorkbook.close
objExcel.Quit()
Set objWorksheet1 = Nothing
Set objWorkbook = Nothing
Set ObjExcel = Nothing


WScript.Quit