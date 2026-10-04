Attribute VB_Name = "Module1"
Sub CleanRawData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long
    Dim cell As Range
    
    Set ws = ThisWorkbook.Sheets("RawData")
    
    Application.ScreenUpdating = False
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    For i = lastRow To 2 Step -1
        If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(i, 1), ws.Cells(i, lastCol))) = 0 Then
            ws.Rows(i).Delete
        End If
    Next i
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For Each cell In ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol))
        If cell.HasFormula = False And VarType(cell.Value) = vbString Then
            cell.Value = Trim(cell.Value)
        End If
    Next cell
    
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).RemoveDuplicates Columns:=1, Header:=xlYes
    
    Application.ScreenUpdating = True
    
    MsgBox "Data cleaned successfully!", vbInformation
    
End Sub
Sub GenerateSummary()
    Dim wsData As Worksheet
    Dim wsDash As Worksheet
    Dim lastRow As Long
    Dim regions() As String
    Dim categories() As String
    Dim i As Integer
    Dim r As Long
    
    Set wsData = ThisWorkbook.Sheets("RawData")
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    
    wsDash.Range("A4:Z100").Clear
    
    wsDash.Range("A4") = "Total Sales"
    wsDash.Range("B4") = "=SUM(RawData!L2:L" & lastRow & ")"
    wsDash.Range("A5") = "Total Profit"
    wsDash.Range("B5") = "=SUM(RawData!O2:O" & lastRow & ")"
    wsDash.Range("A6") = "Total Orders"
    wsDash.Range("B6") = "=COUNTA(RawData!A2:A" & lastRow & ")"
    
    wsDash.Range("A4:A6").Font.Bold = True
    wsDash.Range("B4:B6").NumberFormat = "#,##0.00"
    
    regions = Split("East,West,Central,South", ",")
    wsDash.Range("A9") = "Region"
    wsDash.Range("B9") = "Total Sales"
    wsDash.Range("A9:B9").Font.Bold = True
    
    r = 10
    For i = 0 To UBound(regions)
        wsDash.Cells(r, 1) = regions(i)
        wsDash.Cells(r, 2) = "=SUMIF(RawData!G2:G" & lastRow & ",A" & r & ",RawData!L2:L" & lastRow & ")"
        wsDash.Cells(r, 2).NumberFormat = "#,##0.00"
        r = r + 1
    Next i
    
    categories = Split("Furniture,Office Supplies,Technology", ",")
    r = r + 2
    wsDash.Cells(r, 1) = "Category"
    wsDash.Cells(r, 2) = "Total Sales"
    wsDash.Range(wsDash.Cells(r, 1), wsDash.Cells(r, 2)).Font.Bold = True
    r = r + 1
    
    For i = 0 To UBound(categories)
        wsDash.Cells(r, 1) = categories(i)
        wsDash.Cells(r, 2) = "=SUMIF(RawData!I2:I" & lastRow & ",A" & r & ",RawData!L2:L" & lastRow & ")"
        wsDash.Cells(r, 2).NumberFormat = "#,##0.00"
        r = r + 1
    Next i
    
    wsDash.Columns("A:B").AutoFit
    
    MsgBox "Summary generated successfully!", vbInformation
    
End Sub
Sub CreatePivotTable()
    Dim wsData As Worksheet
    Dim wsDash As Worksheet
    Dim pvtCache As PivotCache
    Dim pvtTable As PivotTable
    Dim lastRow As Long
    Dim lastCol As Long
    Dim pvtRange As Range
    Dim pt As PivotTable
    
    Set wsData = ThisWorkbook.Sheets("RawData")
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    
    On Error Resume Next
    For Each pt In wsDash.PivotTables
        pt.TableRange2.Clear
    Next pt
    On Error GoTo 0
    
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    lastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column
    
    Set pvtRange = wsData.Range(wsData.Cells(1, 1), wsData.Cells(lastRow, lastCol))
    
    Set pvtCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=pvtRange)
    
    Set pvtTable = pvtCache.CreatePivotTable( _
        TableDestination:=wsDash.Range("A30"), _
        TableName:="SalesPivot")
    
    With pvtTable
        .PivotFields("Category").Orientation = xlRowField
        .PivotFields("Sub-Category").Orientation = xlRowField
        .PivotFields("Sales").Orientation = xlDataField
        .PivotFields("Sum of Sales").Function = xlSum
        .PivotFields("Profit").Orientation = xlDataField
        .PivotFields("Sum of Profit").Function = xlSum
        .PivotFields("Region").Orientation = xlColumnField
    End With
    
    MsgBox "Pivot table created successfully!", vbInformation
    
End Sub
Sub CreateChart()
    Dim wsDash As Worksheet
    Dim pt As PivotTable
    Dim cht As ChartObject
    Dim dataRange As Range
    
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    
    On Error Resume Next
    Set pt = wsDash.PivotTables("SalesPivot")
    On Error GoTo 0
    
    If pt Is Nothing Then
        MsgBox "Pivot table not found! Run CreatePivotTable first.", vbExclamation
        Exit Sub
    End If
    
    For Each cht In wsDash.ChartObjects
        cht.Delete
    Next cht
    
    Set dataRange = pt.TableRange1
    
    Set cht = wsDash.ChartObjects.Add(Left:=300, Top:=20, Width:=450, Height:=300)
    
    With cht.Chart
        .SetSourceData Source:=dataRange
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Sales & Profit by Category"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Category"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Amount"
    End With
    
    MsgBox "Chart created successfully!", vbInformation
    
End Sub
Sub FormatDashboard()
    Dim wsDash As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    
    wsDash.Range("A1").Font.Size = 16
    wsDash.Range("A1").Font.Bold = True
    wsDash.Range("A1").Font.Name = "Arial"
    
    wsDash.Range("A4:A6").Font.Size = 11
    wsDash.Range("B4:B6").NumberFormat = "$#,##0.00"
    
    wsDash.Range("B10:B13").NumberFormat = "$#,##0.00"
    
    lastRow = 13
    For Each cell In wsDash.Range("B10:B13")
        If cell.Value < 0 Then
            cell.Font.Color = RGB(255, 0, 0)
        Else
            cell.Font.Color = RGB(0, 128, 0)
        End If
    Next cell
    
    With wsDash.Range("A9:B9")
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    
    Dim catHeaderRow As Long
    catHeaderRow = 16
    With wsDash.Range("A" & catHeaderRow & ":B" & catHeaderRow)
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    
    wsDash.Range("B17:B19").NumberFormat = "$#,##0.00"
    
    wsDash.Columns("A:B").AutoFit
    wsDash.Cells.Borders.LineStyle = xlNone
    
    With wsDash.Range("A9:B13").Borders
        .LineStyle = xlContinuous
        .Color = RGB(200, 200, 200)
    End With
    
    With wsDash.Range("A16:B19").Borders
        .LineStyle = xlContinuous
        .Color = RGB(200, 200, 200)
    End With
    
    MsgBox "Dashboard formatted successfully!", vbInformation
    
End Sub
Sub GenerateFullReport()
    Application.ScreenUpdating = False
    
    Call CleanRawData
    Call GenerateSummary
    Call CreatePivotTable
    Call CreateChart
    Call FormatDashboard
    
    Application.ScreenUpdating = True
    
    ThisWorkbook.Sheets("Dashboard").Activate
    
    MsgBox "Full report generated successfully!", vbInformation, "Report Ready"
    
End Sub
