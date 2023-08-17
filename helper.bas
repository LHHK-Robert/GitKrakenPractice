Function CopyRange(sourceSheetName As String, destinationSheetName As String, startRow As Long, startColumn As Long, endRow As Long, endColumn As Long, pasteStartRow As Long, pasteStartColumn As Long)
    Dim sourceSheet As Worksheet
    Dim destinationSheet As Worksheet
    Dim sourceRange As Range
    Dim destinationRange As Range
    
    ' Set the source and destination sheets
    Set sourceSheet = ThisWorkbook.Sheets(sourceSheetName)
    Set destinationSheet = ThisWorkbook.Sheets(destinationSheetName)
    
    ' Set the source and destination ranges
    Set sourceRange = sourceSheet.Range(sourceSheet.Cells(startRow, startColumn), sourceSheet.Cells(endRow, endColumn))
    Set destinationRange = destinationSheet.Range(destinationSheet.Cells(pasteStartRow, pasteStartColumn), destinationSheet.Cells(pasteStartRow + endRow - startRow, pasteStartColumn + endColumn - startColumn))
    
    ' Copy the range from the source sheet to the destination sheet
    sourceRange.Copy Destination:=destinationRange
End Function

Sub FindValue()
    Dim searchRange As Range
    Dim foundCell As Range
    Dim searchValue As String
    
    ' Set the search range and value
    Set searchRange = ThisWorkbook.Sheets("Sheet1").Range("A:A")
    searchValue = "apple"
    
    ' Search for the value in the range
    Set foundCell = searchRange.Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole)
    
    ' Check if the value was found
    If Not foundCell Is Nothing Then
        ' Value was found
        MsgBox "Value found in cell: " & foundCell.Address
    Else
        ' Value was not found
        MsgBox "Value not found"
    End If
End Sub

Sub CreateOrClearSheet(sheetName As String)
    Dim newSheet As Worksheet
    On Error Resume Next
    Set newSheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If newSheet Is Nothing Then
        ' Create a new sheet with the specified name
        Set newSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        newSheet.Name = sheetName
    Else
        ' Clear the existing sheet
        newSheet.Cells.Clear
    End If
End Sub

Sub FindLastColumn()
    Dim lastColumn As Long
    With ThisWorkbook.Sheets("Sheet1")
        lastColumn = .Cells(1, .Columns.Count).End(xlToLeft).Column
    End With
    MsgBox "The last used cell in row 1 is: " & lastColumn
End Sub

Sub FindLastCell()
    Dim lastRow As Long
    With ThisWorkbook.Sheets("Sheet1")
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
    MsgBox "The last used cell in column A is: " & lastRow
End Sub

Sub CopyRangeWithLoop()
    Dim sourceSheet As Worksheet
    Dim destinationSheet As Worksheet
    Dim sourceRange As Range
    Dim destinationRange As Range
    Dim i As Integer
    
    ' Set the source and destination sheets
    Set sourceSheet = ThisWorkbook.Sheets("Sheet1")
    Set destinationSheet = ThisWorkbook.Sheets("Sheet2")
    
    ' Loop through the rows in the source sheet
    For i = 1 To 10
        ' Set the source and destination ranges
        Set sourceRange = sourceSheet.Range("A" & i & ":D" & i)
        Set destinationRange = destinationSheet.Range("E" & i & ":H" & i)
        
        ' Copy the range from the source sheet to the destination sheet
        sourceRange.Copy Destination:=destinationRange
    Next i
End Sub