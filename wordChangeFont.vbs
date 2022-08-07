Set objWord = CreateObject("Word.Application")
objWord.Visible = False
Set objDoc = objWord.Documents.Open("C:\MergeAfterWrite.docx")

Set objRange = objDoc.Range()
'objDoc.Tables.Add objRange,1,3

Set objTable = objDoc.Tables(2)
objTable.Cell(1, 1).Range.Font.Bold = True

For j = 1 To objTable.Columns.Count
      objTable.Cell(2, j).Range.Font.Bold = True

Next

objDoc.Save
objWord.Quit
MsgBox "Bitti"