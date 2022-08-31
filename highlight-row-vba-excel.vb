'highlights active row
Sub Worksheet_SelectionChange(ByVal Target As Excel.Range)

Static xRow

If xRow <> "" Then
    With Rows(xRow).Interior
        .ColorIndex = xlNone
    End With
End If

Active_Row = Selection.Row
xRow = Active_Row
With Rows(Active_Row).Interior
    .ColorIndex = 27
    .Pattern = xlSolid
End With

End Sub
