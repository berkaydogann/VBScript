Set objXLApp = CreateObject("Excel.Application")
objXLApp.Visible = False
objXLApp.DisplayAlerts = False
Set objXLWb = objXLApp.Workbooks.Open("C:\Users\user\Desktop\Testler\BerkayDeneme - Copy.xlsx")
objXLWb.Worksheets(1).Range("L2").AutoFilter 10,"<>Accessorize"
objXLWb.Worksheets(1).Range("L2").AutoFilter 29,"<>Pasif"

objXLWb.SaveAs "C:\Users\user\Desktop\Testler\Perakende_072021_Son.xlsx", 51 
objXLWb.Close
objXLApp.Quit

