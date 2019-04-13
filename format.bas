Sub weekends()

Application.ScreenUpdating = False

lastColumn = 93
lastRow = 2007

For x = 6 To lastRow Step 7
    Range(Cells(x - 1, 2), Cells(x, lastColumn)).Interior.Color = RGB(217, 217, 217)
Next x

Application.ScreenUpdating = True

End Sub

Sub months()

Application.ScreenUpdating = False

lastColumn = 89
lastRow = 2007

For x = 3 To lastRow
    Set CurrentCell = Cells(x, 3)
    If CurrentCell.Value = 1 Then
        With Range(Cells(x, 1), Cells(x, lastColumn)).Borders(xlEdgeTop)
         .LineStyle = xlContinuous
         .Weight = xlThick
         .ColorIndex = xlColorIndexAutomatic
        End With
    End If
Next x

Application.ScreenUpdating = True

End Sub

