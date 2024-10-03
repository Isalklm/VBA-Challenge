Attribute VB_Name = "Module2"
Sub ApplyColorConditionalFormatting()

    ' Declare variables
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim cell As Range
    Dim RelevantSheets As Variant
    Dim ColsToFormat As Variant
    Dim i As Integer

    ' Specify which sheets to apply this formatting to
    
    RelevantSheets = Array("Q1", "Q2", "Q3", "Q4")   ' Example sheets
    ColsToFormat = Array(11, 12)    ' Column K and L

    ' Loop through each relevant worksheet
    For Each ws In ThisWorkbook.Sheets
        If Not IsError(Application.Match(ws.Name, RelevantSheets, 0)) Then
        
            ' Find the last row in the "Quarterly Change" column (Column K)
            LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row

            ' Loop through each specified column (K for Quarterly Change, L for Percent Change)
            For i = LBound(ColsToFormat) To UBound(ColsToFormat)
            
                ' Loop through each cell in the column from row 2 to the last row
                For Each cell In ws.Range(ws.Cells(2, ColsToFormat(i)), ws.Cells(LastRow, ColsToFormat(i)))
                    If cell.Value > 0 Then
                    
                        ' If the value is positive, set the background color to green
                        cell.Interior.ColorIndex = 4       ' Green color
                    ElseIf cell.Value < 0 Then
                    
                        ' If the value is negative, set the background color to red
                        cell.Interior.ColorIndex = 3     ' Red color
                    Else
                        ' If the value is zero, no color change
                        cell.Interior.ColorIndex = 0 ' No fill
                    End If
                Next cell
            Next i
        End If
    Next ws
End Sub

