Sub KopiereSoftwareZuordnung()
    Dim sheet1 As Worksheet
    Dim sheet2 As Worksheet
    Dim lastRowSheet1 As Long
    Dim lastRowSheet2 As Long
    Dim rngSheet1 As Range
    Dim rngSheet2 As Range
    Dim cell As Range
    Dim dict As Object
    Dim outputRow As Long
    
    ' Set the sheet you are working with
    Set sheet1 = ThisWorkbook.Sheets("X")
    Set sheet2 = ThisWorkbook.Sheets("Y")
    
    ' Define the last row of both
    lastRowSheet1 = sheet1.Cells(sheet1.Rows.Count, 1).End(xlUp).Row
    lastRowSheet2 = sheet2.Cells(sheet2.Rows.Count, 1).End(xlUp).Row
    
    ' Range of both sheets
    Set rngSheet1 = sheet1.Range("A2:A" & lastRowSheet1)
    Set rngSheet2 = sheet2.Range("A2:A" & lastRowSheet2)
    
    ' Dictionary for sorting out duplicates
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Save all Data that already exists in Sheet2
    For Each cell In rngSheet2
        If Not dict.exists(cell.Value) Then
            dict.Add cell.Value, cell.Value
        End If
    Next cell
    
    ' Look for all values, with a certain condition
    outputRow = lastRowSheet2
    For Each cell In sheet1
        If IsError(wsInstallierte.Cells(cell.Row, 4).Value) Then ' The condition
            If dict.exists(cell.Value) Then ' if the value isn't in sheet 2
                wsInstallierte.Cells(outputRow, 1).Value = cell.Value ' then paste it in sheet 2
                outputRow = outputRow + 1 ' increment the counter
            End If
        End If
    Next cell
    
    ' Clean the dictionary
    Set dict = Nothing
    
    MsgBox "Success", vbInformation
End Sub
