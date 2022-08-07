Set objXLApp = CreateObject("Excel.Application")
objXLApp.Visible = False


Set objXLWb = objXLApp.Workbooks.Open("C:\Users\user\Desktop\Test.xlsx")
Set oSheet = objXLWb.Worksheets(1) 
 
oSheet.Columns("B:B").Hidden = True

objXLWb.save
objXLWb.close false
objXLApp.Quit

Set objXLWb = nothing
Set objXLApp = nothing