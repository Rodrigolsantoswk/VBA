'A sub to open a SQL Server Connection and search query results
Sub Connection()
  'Declaring the variables using ADODB
  Dim conn As ADODB.Connection
  Dim rs As ADODB.Recordset
  Dim sConnString As String
  SERVER = "ServerName"
  'The connection String
  sConnString = "Provider = sqloledb; " & _
                        "Data Source=" & SERVER & "; " & _
                        "Initial Catalog= DatabaseName;" & _
                        "User ID =UserName;" & _
                        "Password =Password;"
  'Initializing Connection and ResultSet
  Set conn = New ADODB.Connection
  Set rs = New ADODB.Recordset
  'Opening Connection using Connection String
  conn.Open sConnString
  'Searching results for a query
  Query= "Query Here"
  Set rs= conn.ExecteQuery(Query)
  'Inserting data in the Sheet
  If Not rs.EOF Then
    Sheets("SheetName").Cells(2, 1).CopyFromRecordset rs
    rs.Close
  End If
  'Closing the connection
  If CBool(conn.State And adStateOpen) Then conn.Close
  Set conn = Nothing
  Set rs = Nothing
End Sub
