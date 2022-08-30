Sub ForEachCell404En()
    Dim Cell As Range
    Dim Result() As String
    Dim URL() As String
    Dim Count As Integer
    Dim ConcatenatedUrl As String
    Dim lastRow As Long
    Dim CountLoop As Integer
    
    CountLoop = 3
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For Each Cell In Sheets("404 English").Range("A3:A3")
        Result() = Split(Cell.Value, "//")
        UrlArr = Split(Result(1), "/")
        Count = UBound(UrlArr)
        
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
        'Cell.Offset(0, 4).Value = ConcatenatedUrl
        Cell.Offset(0, 2).Value = GetPageTitle(Cell.Value)
        CountLoop = CountLoop + 1
        Debug.Print CountLoop
    Next Cell
    
End Sub

Function GetPageTitle(ByVal URL As String) As String
    Dim bot As New WebDriver
    Dim Response As String
    Dim Error404 As String
    On Error GoTo ClearError
    
    bot.Start "chrome"
    bot.Get URL
    bot.Wait 3000
    'GetPageTitle = bot.Window.Title
    Error404 = bot.FindElementById("error-404").Text
    
    TimeFuture = Now() + TimeValue("00:00:02")
    
    'Get 404 error text if it exists
    If (IsEmpty(bot.FindElementById("error-404").Text) = False) Then
        GetPageTitle = bot.FindElementById("error-404").Text
    ElseIf ((IsEmpty(bot.Window.Title)) = False) Then
        GetPageTitle = bot.Window.Title
    Else
        GetPageTitle = "No Title"
    End If

ProcExit:
    Exit Function
ClearError:
    Debug.Print "'GetPageTitle' Run-time error '" _
        & Err.Number & "':" & vbLf & "    " & Err.Description
    ''GetPageTitle' Run-time error '13': UnknownError
    If (Err.Number = 13) Then
        GetPageTitle = "Site cannot be reached"
    'Run-time error '7': NoSuchElementError Element not found for Id=error-404
    ElseIf (Err.Number = 7) Then
        GetPageTitle = bot.Window.Title
    End If
    Resume ProcExit
End Function

Private Sub TestTime()
    TimeFuture = Now() + TimeValue("00:00:02")
    'The following will run after 5 seconds. Note MsgModal() is in a separate Module ModuleMsgModal
    Application.OnTime TimeFuture, "'MsgModal""Now()""'"
    'The following will display message box if current time < the set future time
    If (Time() < TimeFuture) Then
        MsgBox (Time() & " : " & TimeFuture)
    End If
End Sub



