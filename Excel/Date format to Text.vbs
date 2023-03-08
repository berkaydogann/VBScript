Dim CellValue, NumRows
Set objXLApp = CreateObject("Excel.Application")
objXLApp.Visible = false

Set objXLWb = objXLApp.Workbooks.Open("C:\Users\berka\Desktop\VbsTestler\LoginLogout.xlsx")
Set objWorksheet1 = objXLWb.Worksheets("LoginLogout")
Set objRange = objWorksheet1.UsedRange



objRange.Cells(1, 1).Numberformat = "@"
objRange.Cells(1, 1).Value = DateConvertedToText
MsgBox objRange.Cells(1, 1).Value

objXLWb.saveAs "C:\Users\berka\Desktop\VbsTestler\LoginLogoutNew.xlsx",51
objXLWb.close false
objXLApp.Quit

Set objXLWb = nothing
Set objXLApp = nothing




