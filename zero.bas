Dim currentRow As Variant
Dim currentCol As Variant
Dim CurrentCell As Range

Sub zeroEmpty()

Application.ScreenUpdating = False

lastColumn = 93

currentRow = Selection.row

For i = 6 To lastColumn
    Set CurrentCell = Cells(currentRow, i)
    If (TypeName(CurrentCell.Value) = "Empty") And _
       (columns(i).Hidden = False) Then
        CurrentCell.Value = 0
    End If
Next i

Application.ScreenUpdating = True

End Sub



Sub zeroEmptyCategory()

Application.ScreenUpdating = False

currentRow = Selection.row
currentCol = Selection.column

While Cells(currentRow, currentCol).Borders(xlEdgeLeft).Weight <> xlThick
    currentCol = currentCol - 1
Wend

Do
    Set CurrentCell = Cells(currentRow, currentCol)
    If (TypeName(CurrentCell.Value) = "Empty") And _
       (columns(currentCol).Hidden = False) Then
        CurrentCell.Value = 0
    End If
    currentCol = currentCol + 1
Loop While Cells(currentRow, currentCol).Borders(xlEdgeLeft).Weight <> xlThick

Cells(currentRow, currentCol).Select

Application.ScreenUpdating = True

End Sub





