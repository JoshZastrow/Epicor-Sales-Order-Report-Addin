Sub Download_Standard_BOM()
'Initializes variables
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim ConnectionString As String
Dim StrQuery As String
Dim SODetail() As String
'Setup the connection string for accessing MS SQL database
   'Make sure to change:
       '1: PASSWORD
       '2: USERNAME
       '3: REMOTE_IP_ADDRESS
       '4: DATABASE
    ConnectionString = "Provider=SQLOLEDB.1;" & _
                       "Password=10 mil 4 Freedom!;" & _
                       "Persist Security Info=True;" & _
                       "User ID=HEM/Joshua Zastrow;" & _
                       "Data Source=70.90.240.163;" & _
                       "Use Procedure for Prepare=1;" & _
                       "Auto Translate=True;" & _
                       "Packet Size=4096;" & _
                       "Use Encryption for Data=False;" & _
                       "Tag with column collation when possible=False;" & _
                       "Initial Catalog=HEMSQL1"

    'Opens connection to the database
    cnn.Open ConnectionString
    'Timeout error in seconds for executing the entire query; this will run for 15 minutes before VBA timesout, but your database might timeout before this value
    cnn.CommandTimeout = 900

    'This is your actual MS SQL query that you need to run; you should check this query first using a more robust SQL editor (such as HeidiSQL) to ensure your query is valid
    StrQuery = _
    "SELECT Rel.OurReqQty AS [QTY Owed], Rel.OurStockQty AS [Stock]," & _
       "SO.DocUnitPrice AS [$/Per], ROUND(Rel.OurReqQty * SO.DocUnitPrice, 2) AS [Ext. Price]" & _
    "FROM Erp.OrderRel Rel" & _
    "INNER JOIN Erp.OrderDtl SO ON" & _
      "Rel.Company = SO.Company AND" & _
      "Rel.OrderNum = SO.OrderNum AND" & _
      "Rel.OrderLine = SO.OrderLine" & _
    "WHERE Rel.OrderNum = '75039' AND" & _
      "Rel.OrderLine = '1'    AND" & _
      "Rel.OrderRelNum = '1'"

    'Performs the actual query
    rst.Open StrQuery, cnn
    'Dumps all the results from the StrQuery into cell A2 of the first sheet in the active workbook
    Sheets(1).Range("A2").CopyFromRecordset rst
End Sub