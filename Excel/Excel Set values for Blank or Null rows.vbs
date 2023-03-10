Dim CellValue, NumRows
Set objXLApp = CreateObject(“Excel.Application”)
objXLApp.Visible = false
Set objXLWb = objXLApp.Workbooks.Open("C:\Users\berka\Desktop\VbsTestlerfilter.xlsx")
Set objWorksheet1 = objXLWb.Worksheets(“Sheet1”)
Set objRange = objWorksheet1.UsedRange
For intRowCounter = 2 to objWorksheet1.usedRange.Rows.Count
	CellValue = objWorksheet1.Range(“G”&intRowCounter).Value
	If IsEmpty(objWorksheet1.Range(“G”&intRowCounter).Value) = True or 						    			IsNull(objWorksheet1.Range(“G”&intRowCounter).Value) = True Then
		objWorksheet1.Range(“G”&intRowCounter).Value = “NotNull”
	End If
Next
objXLWb.save
objXLWb.close false
objXLApp.Quit
Set objXLWb = nothing
Set objXLApp = nothing

