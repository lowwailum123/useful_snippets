This version of the macro first searches for the header cell containing the text "arrived_at" in the first row of the worksheet. If the header cell is found, it determines the column name and then proceeds with the same logic as before to check and highlight the old dates. If the header cell is not found, it displays a message box indicating that the "arrived_at" column was not found in the worksheet.
```VBA
Sub HighlightOldDates()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentDate As Date
    Dim cell As Range
    Dim arrivedAtColumn As Range
    Dim headerCell As Range
    Dim columnName As String
    
    ' Set the target worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1") ' Change "Sheet1" to your sheet's name
    
    ' Find the header cell with the name "arrived_at"
    Set headerCell = ws.Rows(1).Find(What:="arrived_at", LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not headerCell Is Nothing Then
        ' Get the column name
        columnName = headerCell.Address(0, 0)
        
        ' Find the last row with data in the found column
        lastRow = ws.Cells(ws.Rows.Count, columnName).End(xlUp).Row
        
        ' Set the column containing "arrived_at" dates
        Set arrivedAtColumn = ws.Range(columnName & "2:" & columnName & lastRow)
        
        ' Get the current date
        currentDate = Date
        
        ' Loop through each cell in the "arrived_at" column
        For Each cell In arrivedAtColumn
            If IsDate(cell.Value) Then
                ' Calculate the difference in months
                Dim monthsPassed As Long
                monthsPassed = DateDiff("m", cell.Value, currentDate)
                
                ' Check if 2 or more months have passed
                If monthsPassed >= 2 Then
                    cell.Interior.Color = RGB(255, 0, 0) ' Highlight the cell in red
                Else
                    cell.Interior.ColorIndex = xlNone ' Clear any previous highlighting
                End If
            End If
        Next cell
    Else
        MsgBox "Column 'arrived_at' not found in the worksheet."
    End If
End Sub
```
