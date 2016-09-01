Attribute VB_Name = "ReformatStockStatusReport"
Sub ReformatStockStatus()

Dim wb As Workbook
Dim ws As Worksheet
Dim data As Worksheet

Set wb = ActiveWorkbook

'Make sure it's the stock status report
If wb.Sheets(1).Range("Q4") <> "Stock Status Report" Then
    MsgBox ("This add-in can only run with the Stock Status Report")
    Exit Sub
End If

Set ws = wb.Sheets("StockStatus")

'Make sure the workbook is correct

'Check for Tabulated Data sheet, else create one
For Each sht In wb.Worksheets
    If sht.Name = "Tabulated Data" Then
        Application.DisplayAlerts = False
        sht.Delete
        Application.DisplayAlerts = True
        Exit For
    End If
Next sht
wb.Worksheets.Add After:=wb.Sheets(Sheets.Count)
wb.Sheets(Sheets.Count).Name = "Tabulated Data"
Set data = wb.Sheets("Tabulated Data")

'Create Header Row
ws.Activate
Set Headers = ws.Range("A12", "AV13")
data.Cells(1, 1) = "PartNum"
data.Cells(1, 2) = ws.Range("A12").Value ' Warehouse
data.Cells(1, 3) = ws.Range("G12").Value ' Part Class
data.Cells(1, 4) = ws.Range("k12").Value ' Type
data.Cells(1, 5) = ws.Range("P13").Value ' On Hand Qty
data.Cells(1, 6) = ws.Range("W13").Value ' Base On Hand
data.Cells(1, 7) = ws.Range("AD13").Value ' Unit Cost
data.Cells(1, 8) = ws.Range("AJ13").Value ' Mat'l Burden
data.Cells(1, 9) = ws.Range("AR12").Value ' Mth
data.Cells(1, 10) = ws.Range("AV12").Value ' Extended Cost
data.UsedRange.Font.Bold = True

'Set Data Range
LstRow = ws.Rows(ws.Rows.Count).End(xlUp).Row
Set DataRange = ws.Range("A23", "A" & LstRow)

'Pull Data
For Each rw In DataRange
    If InStr(rw.Value, "Part") Then
    
        'Newest row on tabulated data
        NextRow = data.Rows(data.Rows.Count).End(xlUp).Offset(1, 0).Row
        Debug.Print Mid(rw.Value, 11, InStr(12, rw.Value, " ") - 10)
        data.Cells(NextRow, 1) = Mid(rw.Value, 11, InStr(12, rw.Value, " ") - 10)
        data.Cells(NextRow, 2) = rw.Cells(3, 1)
        data.Cells(NextRow, 3) = rw.Cells(4, 6)
        data.Cells(NextRow, 4) = rw.Cells(4, 10)
        data.Cells(NextRow, 5) = rw.Cells(6, 13) ' On Hand Qty
        data.Cells(NextRow, 6) = rw.Cells(6, 22) ' BAse On Hand
        data.Cells(NextRow, 7) = rw.Cells(6, 28) ' Unit Cost
        data.Cells(NextRow, 8) = rw.Cells(4, 34)
        data.Cells(NextRow, 9) = rw.Cells(4, 42)
        data.Cells(NextRow, 10) = rw.Cells(4, 46) ' Extended Cost

        
    End If

Next rw
End Sub

