Set xlapp = CreateObject("Excel.Application")
xlapp.Visible = False
xlapp.DisplayAlerts = False
set obj1 = xlapp.Workbooks.Open("C:\Users\user\Desktop\Test.xlsx")
Set obj2=obj1.sheets.Add  
obj2.name="Sheet5"
obj2.name="Sheet6"
obj1.Save
obj1.Close
Set obj1=Nothing                                 
Set obj2 = Nothing                               
Set xlapp=Nothing
