Set objXLApp = CreateObject("Excel.Application")
objXLApp.Visible = false
Set objXLWb = objXLApp.Workbooks.Open("C:\Users\Desktop\Test.xlsx")
Set oSheet = objXLWb.Worksheets(1)

oSheet.Range("A1:L45").WrapText = True
oSheet.Columns("A:L").VerticalAlignment   = -4108            
oSheet.Columns("A:L").HorizontalAlignment = -4108                                       

objXLWb.save
objXLWb.close false
objXLApp.Quit

Set objXLWb = nothing
Set objXLApp = nothing

