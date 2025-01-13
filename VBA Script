Sub InsertDynamicData()
    Dim conn As Object
    Dim cmd As Object
    Dim connString As String
    Dim sql As String
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim rowData As String
    Dim i As Integer
    Dim lastRow As Long
    Dim lastCol As Long
    Dim userName As String
    Dim timeStamp As String

   ' Get the username of the person who opened the workbook
    userName = Environ("USERNAME")
    
    ' Get the current timestamp in HHMMSS format
    timeStamp = Format(Now, "HHMMSS")
    
    ' Set your worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" to your sheet name

    ' Find the last row and column with data
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Set the range dynamically
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    ' Connection string to SAP BW HANA
    ' SERVER:PORT  - use your hana server name and port 
    ' DOMAIN - use your Kerberous domain. Kerberous External ID example: user@corp.company.com
    connString = "Driver={HDBODBC};ServerNode=SERVER:PORT;SNC=1;SNC_PARTNERNAME=p:CN=" & userName & "@DOMAIN"

    ' Create a connection object
    Set conn = CreateObject("ADODB.Connection")

    On Error GoTo ConnectionError

    ' Open the connection
    conn.Open connString

    ' Loop through each row in the range
    For Each cell In rng.Rows
        rowData = ""
        For i = 1 To lastCol ' Loop through all columns dynamically
            rowData = rowData & "'" & cell.Cells(1, i).Value & "',"
        Next i
        ' Add timestamp and username to the row data
        rowData = rowData & "'" & timeStamp & userName & "'"
        
       ' rowData = Left(rowData, Len(rowData) - 1) ' Remove the last comma

        ' SQL query to insert data
        ' YOUR_SCHEMA - is a schema where DB table you insert data is placed
        sql = "INSERT INTO YOUR_SCHEMA.T3 (" & GetColumnNames(lastCol) & " ) VALUES (" & rowData & ")"
       
        ' Create a command object
        Set cmd = CreateObject("ADODB.Command")
        cmd.ActiveConnection = conn
        cmd.CommandText = sql

        ' Execute the query
        cmd.Execute
    Next cell

    ' If insertion is successful
    MsgBox "Data inserted successfully!", vbInformation

    ' Close the connection
    conn.Close
    Set conn = Nothing
    Set cmd = Nothing
    Exit Sub

ConnectionError:
    MsgBox "Failed to connect to SAP BW: " & Err.Description, vbCritical
    Set conn = Nothing
End Sub

Function GetColumnNames(lastCol As Long) As String
    Dim colNames As String
    Dim i As Integer

    ' Generate column names dynamically
    For i = 1 To lastCol
        colNames = colNames & "COLUMN" & i & ","
       
    Next i
    colNames = Left(colNames, Len(colNames) - 1) ' Remove the last comma
    colNames = colNames & "," & "COLUMN3"
    
    GetColumnNames = colNames
End Function
