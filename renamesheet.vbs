Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = false
objExcel.displayalerts=false

Set objWorkbook = objExcel.Workbooks.open("C:\Users\user\Desktop\Test\Asistan_HGO_Taslak.xlsx")
Set objWorksheet1 = objWorkbook.Worksheets("Hedefler")

objWorksheet1.Name = "Hedefler234"

objWorkbook.save
objWorkbook.close
objExcel.Quit()
Set objWorksheet1 = Nothing
Set objWorkbook = Nothing
Set ObjExcel = Nothing


WScript.Quit