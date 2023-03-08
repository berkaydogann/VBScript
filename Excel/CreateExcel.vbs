strFileName = "C:\Users\berka\Desktop\VbsTestler\Excel.xlsx"

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

Set objWorkbook = objExcel.Workbooks.Add()
objWorkbook.SaveAs(strFileName)

objExcel.Quit