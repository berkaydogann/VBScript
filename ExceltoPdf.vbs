Dim Excel
Dim ExcelDoc
    
Set Excel = CreateObject("Excel.Application")

'Open the Document
Set ExcelDoc = Excel.Workbooks.open("C:\Users\user\Downloads\Kalite_Taslak.xlsx")
'ExcelDoc.Worksheets("Özet").Range("A3:N23").Copy
'ExcelDoc.Worksheets("Sayfa1").Range("E15:K30").PasteSpecial -4163
ExcelDoc.Worksheets("Özet").Range("C3:T19").ExportAsFixedFormat xlTypePDF,"C:\Users\user\Downloads\22.03Deneme.pdf" ,xlQualityStandard, True, True,1,5,True

ExcelDoc.save
Excel.ActiveWorkbook.Close
Excel.Application.Quit
