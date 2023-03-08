Set objXLApp = CreateObject("Excel.Application")
objXLApp.Visible = False
objXLApp.DisplayAlerts = False
Set objXLWb = objXLApp.Workbooks.Open("C:\Users\berka\Desktop\VbsTestler\Excel.xlsx")
objXLWb.Worksheets(1).Range("L2").AutoFilter 10,"<>Accessorize"
objXLWb.Worksheets(1).Range("L2").AutoFilter 29,"<>Pasif"

objXLWb.SaveAs "C:\Users\berka\Desktop\VbsTestler\ExcelFilter.xlsx", 51 
objXLWb.Close
objXLApp.Quit

