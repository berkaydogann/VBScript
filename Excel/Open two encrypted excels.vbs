Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = false
objExcel.displayalerts=false

Set objWorkbook = objExcel.Workbooks.Open ("C:\Users\berka\Desktop\VbsTestler\CopyExcel.xlsx",,True,,"123") 

Set objWorksheet1 = objWorkbook.Worksheets("Sheet1")

objWorksheet1.SaveAs "C:\Users\berka\Desktop\VbsTestler\CopyExcel.xlsx",,""
ReadOnly=False
IgnoreReadOnlyRecommended=true


objWorkbook.save
objWorkbook.close
objExcel.Quit()
Set objWorksheet1 = Nothing
Set objWorkbook = Nothing
Set ObjExcel = Nothing






