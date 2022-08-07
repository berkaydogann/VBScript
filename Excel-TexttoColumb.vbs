Set objXLApp = CreateObject("Excel.Application")
objXLApp.Visible = false

Set objXLWb = objXLApp.Workbooks.Open("C:\Users\user\Downloads\satis_taahhut_rapor.xlsx")
Set oSheet = objXLWb.Worksheets("report") 


set date2 = oSheet.Range("N15")
date2 = cDate("date2")

objXLWb.save
objXLWb.close false
objXLApp.Quit

Set objXLWb = nothing
Set objXLApp = nothing
