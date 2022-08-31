Sub ConnectSQLRedirects()
    Dim connection As ADODB.connection
    Set connection = New ADODB.connection
    Let server_name = "D117330\MSSQLSERVER02"
    Let database_name = "CDICollege"
    
    Dim sqlQuery As String
    sqlQuery = "SELECT * FROM [CDICollege].[dbo].[Redirects] WHERE [To] LIKE '%/study-on-campus/%/admissions/%'"
    'sqlQuery = "SELECT *  FROM [CDICollege].[dbo].[Redirects] Where RedirectID = 2;"
    
    Dim rsSql As New ADODB.Recordset

    
    With connection
        .ConnectionString = "Provider=SQLNCLI11;Server=" & server_name & _
            ";database=" & database_name & ";Integrated Security=SSPI;"
        .ConnectionTimeout = 10
        .Open
    End With
    
    If connection.State = 1 Then
        Debug.Print "Connected!"
    End If
    
    rsSql.CursorLocation = adUseClient
    rsSql.Open sqlQuery, connection, adOpenStatic
    
    ThisWorkbook.Sheets("SQLRedirectTable").Range("A2").CopyFromRecordset rsSql
End Sub
