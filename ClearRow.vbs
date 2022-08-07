Set objXLApp = CreateObject("Excel.Application")
objXLApp.Visible = false
objXLApp.displayalerts=false
Set objXLWb = objXLApp.Workbooks.Open("C:\Users\user\Desktop\Test\FTE_Taslak.xlsx")

objXLWb.Worksheets("Bireysel T-1").Range("A1:E4").ClearContents
objXLWb.save
objXLWb.close false
objXLApp.Quit

Set objXLWb = nothing
Set objXLApp = nothing
