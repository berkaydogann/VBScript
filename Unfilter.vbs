Set objXLApp = CreateObject("Excel.Application")
objXLApp.Visible = False
objXLApp.DisplayAlerts = False
Set objXLWb = objXLApp.Workbooks.Open("C:\Users\user\Desktop\Testler\BerkayDeneme.xlsx")
objXLWb.Worksheets(1).Rows.Hidden = False
objXLWb.Worksheets(1).AutoFilterMode  = False

objXLWb.SaveAs "C:\Users\user\Desktop\Testler\Perakende_072021_Son.xlsx", 51 
objXLWb.Close
objXLApp.Quit
