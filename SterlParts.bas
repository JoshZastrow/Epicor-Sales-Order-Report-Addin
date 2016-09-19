Attribute VB_Name = "SterlParts"
Sub ConnectSqlServer()

    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sConnString As String
    Dim ws As Worksheet
    
    'Delete Sales backlog sheet and re add it
    For Each sht In ActiveWorkbook.Worksheets
        If sht.Name = "Sales Backlog" Then
            Application.DisplayAlerts = False
            sht.Delete
            Application.DisplayAlerts = True
        End If
    Next sht
    ActiveWorkbook.Sheets.Add after:=ActiveWorkbook.Sheets(Sheets.Count)
    ActiveWorkbook.Sheets(Sheets.Count).Name = "Sales Backlog"
    Set ws = ActiveWorkbook.Sheets("Sales Backlog")
    
    ' Create the connection string.
    sConnString = "Provider=SQLOLEDB.1;" & _
                  "Integrated Security=SSPI;" & _
                  "Persist Security Info=False;" & _
                  "Initial Catalog=ERP10PROD;" & _
                  "Data Source=HEMSQL1"
    
    ' Create the Connection and Recordset objects.
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    ' Open the connection and execute.
    conn.Open sConnString
    Query = strFileContent()
    Set rs = conn.Execute(Query)
    
    ' Check we have data.
    If Not rs.EOF Then
        ' Transfer result
        ws.Range("A2").CopyFromRecordset rs
    ' Close the recordset
        rs.Close
    Else
        MsgBox "Error: No records returned.", vbCritical
    End If

    ' Clean up
    If CBool(conn.State And adStateOpen) Then conn.Close
    Set conn = Nothing
    Set rs = Nothing
    ws.Cells(1, 1).Value = "Order"
    ws.Cells(1, 2).Value = "Part"
    ws.Cells(1, 3).Value = "ProdCode"
    ws.Cells(1, 4).Value = "Due"
    ws.Cells(1, 5).Value = "Owed"
    ws.Cells(1, 6).Value = "Stocked"
    ws.Cells(1, 7).Value = "$/Per"
    ws.Cells(1, 8).Value = "Ext. Price"
    ws.Cells(1, 9).Value = "Router"
    ws.Range(Cells(1, 1), Cells(1, 8)).Font.Bold = True
    ws.UsedRange.Columns.AutoFit
    ws.UsedRange.HorizontalAlignment = xlLeft
    ws.Columns(7).NumberFormat = "$#,#00.00"
    ws.Columns(8).NumberFormat = "$#,#00.00"
    
    
End Sub

Function strFileContent() As String

Dim strFilename As String: strFilename = _
"S:\Engineering\Josh Zastrow\Epicor\Epicor-Sales-Order-Report-Addin\SalesOrderInfo.sql"

Dim iFile As Integer: iFile = FreeFile

Open strFilename For Input As #iFile
strFileContent = Input(LOF(iFile), iFile)
Close #iFile

End Function



