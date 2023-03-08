Set objXLApp = CreateObject("Excel.Application") 

objXLApp.Visible = false 

Set objXLWb = objXLApp.Workbooks.Open("C:\Users\user\Desktop\VBS Kodlari\hyperlink\Oviexcel.xlsx") 

Set oSheet = objXLWb.Worksheets(1) 

MsgBox oSheet.Range("C2").Hyperlinks(1).Address
objXLWb.save 

objXLWb.close false 

objXLApp.Quit 

Set objXLWb = nothing 

Set objXLApp = nothing

msgbox bitt

