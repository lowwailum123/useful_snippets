```VBA
Sub CopyAndHighlightOldDates()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim currentDate As Date
    Dim cell As Range
    Dim arrivedAtColumn As Range
    Dim emailColumn As Range
    Dim headerCell As Range
    Dim columnLetter As String
    Dim targetRow As Long
    
    ' Set the source and target worksheets
    Set wsSource = ThisWorkbook.Worksheets("Sheet1") ' Change "Sheet1" to your source sheet's name
    
    ' Create or get the "Processed" worksheet
    Set wsTarget = CreateOrGetProcessedWorksheet()
    

    
    ' Find the header cell with the name "arrived_at" in the source worksheet
    Set headerCell = wsSource.Rows(1).Find(What:="arrived_at", LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not headerCell Is Nothing Then
        columnLetter = Split(Cells(1, headerCell.Column).Address, "$")(1) ' Get the column letter
        
        ' Set the column containing "arrived_at" dates and the email column in the source worksheet
        Set arrivedAtColumn = wsSource.Range(columnLetter & "2:" & columnLetter & wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row)
        
        ' Find the header cell with the name "email" in the source worksheet
        Set headerCell = wsSource.Rows(1).Find(What:="email", LookIn:=xlValues, LookAt:=xlWhole)
        If Not headerCell Is Nothing Then
            Set emailColumn = wsSource.Range(columnLetter & "2:" & columnLetter & wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row)
        End If
        
        ' Get the current date
        currentDate = Date
        
        ' Clear the target worksheet (except the header)
        wsTarget.Rows.Clear
        wsSource.Rows(1).Copy wsTarget.Rows(1)
        
        targetRow = 2 ' Start from the second row in the target worksheet
        
        ' Loop through each cell in the "arrived_at" column of the source worksheet
        For Each cell In arrivedAtColumn
            If IsDate(cell.Value) Then
                ' Calculate the difference in months
                Dim monthsPassed As Long
                monthsPassed = DateDiff("m", cell.Value, currentDate)
                
                ' Check if 2 or more months have passed
                If monthsPassed <= -2 Then
                    ' Copy the entire row to the target worksheet
                    wsSource.Rows(cell.Row).Copy wsTarget.Rows(targetRow)
                    
                    ' Convert email column to hyperlinks
                    If Not emailColumn Is Nothing Then
                        ' Check if the current column is the "email" column
                        If wsSource.Cells(1, cell.Column).Value = "email" Then
                            emailColumn.Cells(cell.Row - 1, 1).Hyperlinks.Add _
                                Anchor:=emailColumn.Cells(cell.Row - 1, 1), _
                                Address:="mailto:" & emailColumn.Cells(cell.Row - 1, 1).Value
                        End If
                    End If
                    
                    targetRow = targetRow + 1
                    
                    ' Highlight the cell in red
                    cell.Interior.Color = RGB(255, 0, 0)
                    cell.Font.Color = RGB(255, 255, 255)
                Else
                    cell.Interior.ColorIndex = xlNone ' Clear any previous highlighting
                End If
            End If
        Next cell
    Else
        MsgBox "Column 'arrived_at' not found in the source worksheet."
    End If
    ConvertEmailColumnToHyperlinks wsTarget
    
End Sub

Function ConvertEmailColumnToHyperlinks(ByVal ws As Worksheet)
    Dim headerCell As Range
    Dim emailColumn As Range
    
    ' Find the header cell with the name "email" in the provided worksheet
    Set headerCell = ws.Rows(1).Find(What:="email", LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not headerCell Is Nothing Then
        ' Set the email column based on the header cell's column
        Set emailColumn = ws.Range(ws.Cells(2, headerCell.Column), ws.Cells(ws.Rows.Count, headerCell.Column).End(xlUp))
        
        ' Convert email column to hyperlinks
        For Each cell In emailColumn
            If cell.Value <> "" Then
                Dim subject As String
                Dim mailtoLink As String
                subject = "Subject"
                mailtoLink = "mailto:" & cell.Value & "?subject=" & subject
                cell.Hyperlinks.Add _
                    Anchor:=cell, _
                    Address:=mailtoLink
            End If
        Next cell
    End If
End Function




Function CreateOrGetProcessedWorksheet() As Worksheet
    Dim wsProcessed As Worksheet
    On Error Resume Next ' Turn off error handling temporarily
    
    ' Try to set the reference to the existing "Processed" worksheet
    Set wsProcessed = ThisWorkbook.Worksheets("Processed")
    
    On Error GoTo 0 ' Reset error handling
    
    If wsProcessed Is Nothing Then
        ' If "Processed" worksheet doesn't exist, add it
        Set wsProcessed = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsProcessed.Name = "Processed"
        
        ' Call the function to convert the email column to hyperlinks in the new "Processed" worksheet
        ConvertEmailColumnToHyperlinks wsProcessed

    End If
    
    ' Return the reference to the "Processed" worksheet
    Set CreateOrGetProcessedWorksheet = wsProcessed
End Function

```
