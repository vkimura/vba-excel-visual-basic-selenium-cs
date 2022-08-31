Sub ForEachCell()
    Dim Cell As Range
    Dim Result() As String
    Dim URL() As String
    Dim Count As Integer
    Dim ConcatenatedUrl As String
    Dim lastRow, i As Long
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For Each Cell In Sheets("404 CDICollege English FinalSql").Range("A2:A5")
        Result() = Split(Cell.Value, "//")
        UrlArr = Split(Result(1), "/")
        Count = UBound(UrlArr)
        ConcatenatedUrl = ""
        
        'create url
        For i = 1 To Count
            ConcatenatedUrl = ConcatenatedUrl & "/" & UrlArr(i)
        Next i
        
        'only append if Count > 1 for edge case when Cell.Value = https://www.cdicollege.ca/
        If Count > 1 Then
            ConcatenatedUrl = ConcatenatedUrl & "/" 'append last forward slash
        End If
        
        'If Count = 1 Then
        '    ConcatenatedUrl = "/" + UrlArr(0) + "/" + UrlArr(1) + "/"
        'ElseIf Count = 2 Then
        '    ConcatenatedUrl = "/" + UrlArr(0) + "/" + UrlArr(1) + "/" + UrlArr(2) + "/"
        'ElseIf Count = 3 Then
        '    ConcatenatedUrl = "/" + UrlArr(0) + "/" + UrlArr(1) + "/" + UrlArr(2) + "/" + UrlArr(3) + "/"
        'End If
        'If Count = 1 Then
        '    ConcatenatedUrl = "https://" + UrlArr(0) + "/formation-en-presentiel/quebec/" + UrlArr(1)
        'ElseIf Count = 2 Then
        '    ConcatenatedUrl = "https://" + UrlArr(0) + "/formation-en-presentiel/quebec/" + UrlArr(1) + "/" + UrlArr(2)
        'ElseIf Count = 3 Then
        '    ConcatenatedUrl = "https://" + UrlArr(0) + "/formation-en-presentiel/quebec/" + UrlArr(1) + "/" + UrlArr(2) + "/" + UrlArr(3)
        'End If
        'For i = 0 To UBound(UrlArr)
        '    ConcatenatedUrl = UrlArr(i)
        'Next i
        
        'Cell.Offset(0, 5).Value = Split(Cell.Value, "/")
        Cell.Offset(0, 6).Value = ConcatenatedUrl
        'Cell.Offset(0, 6).Value = GetPageTitle(ConcatenatedUrl)
    Next Cell
    
End Sub
