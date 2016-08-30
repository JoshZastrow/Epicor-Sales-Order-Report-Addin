Attribute VB_Name = "SalesForecast"
Option Explicit


'Figure out last row
'from row 21 to last row:
'if column A value has the word "Customer":
'get the customer ID
'incorporate condition to check to see if sheets are already made
'create a new sheet
'rename the sheet with the custID
'copy each header field and paste into new sheet
'if column A has data in it:
'copy each cell, paste into the new sheet
'Delete Column A for each of the new customer sheets

Sub Reformat()

Dim LastCell As Range
Dim wb As Workbook
Dim ws As Worksheet
Dim FirstCell As Range
Dim cell As Variant
Dim SlashLocation As Integer
Dim CustID As String
Dim FirstCol As Range
Dim LastCol As Range
Dim field As Range
Dim entry As Range
Dim i As Integer
Dim j As Integer
Dim DataRow As Range
Dim sht As Worksheet
Dim DataSht As String
Dim SODetail() As String
Dim PartNum As String

'Set the New Sheet Name
DataSht = "Tabulated Data"

'Declare the main workbook object and the Data worksheet object
Set wb = ActiveWorkbook
Set ws = wb.Sheets(1)

'Start and end of where the data is
Set LastCell = ws.Range("A10000").End(xlUp)
Set FirstCell = ws.Range("A21")

'initialize row counter
i = 2

'Set first dummy customer Sheet
CustID = "Skip1"

'Check to see if sheets area already made
For Each sht In wb.Worksheets
    If sht.Name = DataSht Then
        Application.DisplayAlerts = False
        sht.Delete
        Application.DisplayAlerts = True
    End If
Next sht
        
'Create new Worksheet
Worksheets.Add After:=Worksheets(Worksheets.Count)
Worksheets(Worksheets.Count).Name = DataSht
        
'Create the "Customer" Field
wb.Worksheets(DataSht).Range("A1").Value = "Customer"
wb.Worksheets(DataSht).Range("A1").Font.Bold = True

'Copy headers and paste into new sheet
For Each field In ws.Range("A21:CA21")
    If field.Value <> "" And field.Value <> "CGrp" Then
        Worksheets(DataSht).Range("Z1").End(xlToLeft).Offset(0, 1).Value = field.Value
    End If
Next field

'looping through each row
For Each cell In ws.Range(FirstCell, LastCell)
    
    If InStr(cell.Value, "Customer") Then

        'Get Customer ID
        SlashLocation = InStr(cell.Value, "/")
        CustID = Mid(cell.Value, 11, SlashLocation - 11)

    ElseIf cell.Value <> "" Then
        
        'fill in the customer field
        wb.Worksheets(DataSht).Range("A10000").End(xlUp).Offset(1, 0) = CustID
        
        Set DataRow = ws.Range("A" & cell.Row, "CA" & cell.Row)
        
        'Get Sales Order Details
        SODetail = split(DataRow(1), "/")
        PartNum = DataRow(7).Value
        debug.print(SODetail(0) & " " & SODetail(1) & " " & SODetail(2))

        'Copy Data to new sheet
        For Each entry In DataRow
        
            'Check for missing product codes
            If entry.Value = "" And entry.Column = 13 Then
                entry.Value = "MISSING PRODCODE"
            End If
            
            'Check for missing dates
            If entry.Value = "" And entry.Column = 28 Then
                entry.Value = "MISSING DATE"
            End If
            
            If entry.Value <> "" Then
                Worksheets(DataSht).Cells(i, 50).End(xlToLeft).Offset(0, 1) = entry.Value
            End If
        Next entry
    
        i = i + 1
        
    End If
Next cell

wb.Worksheets(DataSht).UsedRange.Columns.AutoFit
wb.Worksheets(DataSht).Range("A1:CA1").Font.Bold = True
End Sub


