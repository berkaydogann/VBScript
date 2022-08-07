Dim xlapp ' as excel object 
Dim WSx, WSy ' as excel worksheet 
Dim x, y ' as workbook 
Set xlapp = CreateObject("Excel.Application")
Set x = xlapp.Workbooks.Open("C:\Users\user\Desktop\Test\ucms_genesys_campaign_success_detail.xls")
Set y = xlapp.Workbooks.Open("C:\Users\user\Desktop\Test\Asistan_HGO_Taslak.xlsx") 
Set WSx = x.Worksheets("ucms_genesys_campaign_success_d") 
Set WSy = y.Worksheets("Ham Data")
WSx.Copy WSy ' copy worksheet to other workbook
Set objWorksheet = y.Worksheets("ucms_genesys_campaign_success_d")
objWorksheet.Name = "Ham Data"
'WSy.Name = "Sheet1 adını"
Set WSx = nothing 
Set WSy = nothing 
Set objWorksheet = nothing
y.Save 
y.close 
x.Close

