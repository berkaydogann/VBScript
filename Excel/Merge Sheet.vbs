Set xlapp = CreateObject("Excel.Application")
xlapp.Visible = False
Set x = xlapp.Workbooks.Open("C:\Users\berka\Desktop\VbsTestler\Excel.xlsx")
Set y = xlapp.Workbooks.Open("C:\Users\berka\Desktop\VbsTestler\CopyExcel.xlsx") 
Set WSx = x.Worksheets(1) 
Set WSy = y.Worksheets(1)
WSx.Copy WSy 
x.Save
y.Save
x.Close
y.Close