Set xlapp = CreateObject("Excel.Application")
xlapp.Visible = False
xlapp.DisplayAlerts = False
set x = xlapp.Workbooks.Open("C:\Users\berka\Desktop\VbsTestler\Excel.xlsx")
x.Worksheets("Sayfa3").Delete
x.Save
x.Close
