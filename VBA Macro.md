In this version of the macro, I've added the ability to copy the entire rows that meet the condition to another worksheet ("Sheet2" in this case). The rows are copied below the existing data in the target worksheet. Each time a row is copied, the lastRowTarget variable is incremented to ensure that the next row is pasted in the correct position. The EntireRow.Copy method is used to copy the entire row.

As before, remember to adjust the worksheet names and any other relevant details to match your specific Excel setup.
```VBA
Sub CopyAndHighlightOldDates()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRowSource As Long
    Dim lastRowTarget As Long
    Dim currentDate As Date
    Dim cell As Range
    Dim arrivedAtColumn As Range
    Dim headerCell As Range
    Dim columnName As String
    
    ' Set the source and target worksheets
    Set wsSource = ThisWorkbook.Worksheets("Sheet1") ' Change "Sheet1" to your source sheet's name
    Set wsTarget = ThisWorkbook.Worksheets("Sheet2") ' Change "Sheet2" to your target sheet's name
    
    ' Find the header cell with the name "arrived_at" in the source worksheet
    Set headerCell = wsSource.Rows(1).Find(What:="arrived_at", LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not headerCell Is Nothing Then
        ' Get the column name
        columnName = headerCell.Address(0, 0)
        
        ' Find the last row with data in the found column of the source worksheet
        lastRowSource = wsSource.Cells(wsSource.Rows.Count, columnName).End(xlUp).Row
        
        ' Set the column containing "arrived_at" dates in the source worksheet
        Set arrivedAtColumn = wsSource.Range(columnName & "2:" & columnName & lastRowSource)
        
        ' Get the current date
        currentDate = Date
        
        ' Find the last row with data in the target worksheet
        lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, columnName).End(xlUp).Row + 1
        
        ' Loop through each cell in the "arrived_at" column of the source worksheet
        For Each cell In arrivedAtColumn
            If IsDate(cell.Value) Then
                ' Calculate the difference in months
                Dim monthsPassed As Long
                monthsPassed = DateDiff("m", cell.Value, currentDate)
                
                ' Check if 2 or more months have passed
                If monthsPassed >= 2 Then
                    cell.EntireRow.Copy wsTarget.Rows(lastRowTarget)
                    lastRowTarget = lastRowTarget + 1
                    cell.Interior.Color = RGB(255, 0, 0) ' Highlight the cell in red
                Else
                    cell.Interior.ColorIndex = xlNone ' Clear any previous highlighting
                End If
            End If
        Next cell
    Else
        MsgBox "Column 'arrived_at' not found in the source worksheet."
    End If
End Sub

```
