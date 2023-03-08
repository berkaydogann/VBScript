sFolder = "C:\Users\user\Desktop\VBS Kodlari\AllExcelMerge"
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set objXLApp = CreateObject("Excel.Application")
objXLApp.Visible = false

Set mergeExcel = objXLApp.Workbooks.Open("C:\Users\user\Desktop\VBS Kodlari\AllExcelMerge\out\Merged.xlsx")
Set oMasterSheet = mergeExcel.Worksheets("Sheet1")

For Each oFile In oFSO.GetFolder(sFolder).Files
  
	  If (oFSO.GetExtensionName(oFile.Name)) = "xlsx"  Then
		
		
		Set objXLWb = objXLApp.Workbooks.Open("C:\Users\user\Desktop\VBS Kodlari\AllExcelMerge\" & oFile.Name)
			
			Set WSx = objXLWb.Worksheets(1)
					
			
					lastrow = mergeExcel.Worksheets("Sheet1").UsedRange.Rows.Count + 1
					
					WSx.Range("A2:G100000").Copy
					mergeExcel.Worksheets("Sheet1").Range("A"&lastrow&":"&"G"&lastrow).PasteSpecial Paste =xlValues

					
		objXLWb.close false
	 
  End if  
Next

mergeExcel.save
mergeExcel.close
objXLApp.Quit

Set objXLWb = nothing
Set objXLApp = nothing
MsgBox "bitti"