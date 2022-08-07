Set objXLApp = CreateObject("Excel.Application")
objXLApp.Visible = False
Set objXLWb = objXLApp.Workbooks.Open("C:\Users\Emir\Desktop\Perakende_2021_07_Tapdk.xlsx")
Set oSheet = objXLWb.Worksheets(2)  
Set oSheet2 = objXLWb.Worksheets(1)  

Dim rCount
For x=1 To oSheet.Range("A1").CurrentRegion.Rows.Count
If oSheet.Cells(x,1).EntireRow.Hidden = False Then
rCount= rCount +1 
End If
Next

For i=2 to rCount
oSheet.Cells(i,12).Formula ="=VLOOKUP(I"&i&",'Noktalar (2)'!I:I,1,0)"
oSheet.Cells(i,2).Formula ="=VLOOKUP(A"&i&",'Noktalar (2)'!C:C,1,0)"
Next 

Dim rCount2
For x=1 To oSheet2.Range("A1").CurrentRegion.Rows.Count
If oSheet2.Cells(x,1).EntireRow.Hidden = False Then
rCount2= rCount2 +1 
End If
Next

For i=2 to rCount2
oSheet2.Cells(i,2).Formula ="=VLOOKUP(C"&i&",Noktalar!A:A,1,0)"
Next 

objXLWb.SaveAs "C:\Users\Emir\Desktop\Perakende_072021_kullanılıcakSon.xlsx", 51 
MsgBox rCount
objXLApp.Quit
