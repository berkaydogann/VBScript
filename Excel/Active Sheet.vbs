Set objXLApp = CreateObject("Excel.Application")
objXLApp.Visible = False
Set objXLWb = objXLApp.Workbooks.Open("C:\Users\berka\Desktop\VbsTestler\Test.xlsx")


Set pSheet = objXLWb.Worksheets(3)
pSheet.Activate


objXLWb.save
objXLWb.close false
objXLApp.Quit


Set objXLWb = nothing
Set objXLApp = nothing

