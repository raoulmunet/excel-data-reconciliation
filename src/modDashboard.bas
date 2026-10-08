Attribute VB_Name = "modDashboard"
Option Explicit

' Excel Data Reconciliation - optional dashboard UI (v0.2 preview).
' Import this module AND modReconciliation.bas into the same .xlsm file.
' Designed for Excel 2019 Desktop on Windows; runtime testing pending.

Private Const DASH_SHEET As String = "Control Panel"
Private gLastReport As Workbook

Public Sub BuildDashboard()
    Dim ws As Worksheet
    Dim i As Long
    Dim oldA As String, oldB As String, oldKey As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DASH_SHEET)
    On Error GoTo Failed

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = DASH_SHEET
    Else
        oldA = CStr(ws.Range("B6").Value2)
        oldB = CStr(ws.Range("B8").Value2)
        oldKey = CStr(ws.Range("B10").Value2)
    End If

    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    ws.Cells.Clear

    With ws
        .Cells.Font.Name = "Aptos"
        .Cells.Font.Size = 11
        .Columns("A").ColumnWidth = 24
        .Columns("B").ColumnWidth = 57
        .Columns("C").ColumnWidth = 15
        .Columns("D").ColumnWidth = 21
        .Columns("E").ColumnWidth = 4

        .Range("A1:D2").Merge
        .Range("A1").Value2 = "EXCEL DATA RECONCILIATION"
        .Range("A1:D2").Interior.Color = RGB(29, 53, 87)
        .Range("A1:D2").Font.Color = vbWhite
        .Range("A1:D2").Font.Bold = True
        .Range("A1:D2").Font.Size = 19
        .Range("A1:D2").VerticalAlignment = xlCenter

        .Range("A4").Value2 = "Configure two sources and run a non-destructive comparison."
        .Range("A6").Value2 = "Source A (CSV)"
        .Range("A8").Value2 = "Source B (CSV)"
        .Range("A10").Value2 = "Key column"
        .Range("B6").Value2 = oldA
        .Range("B8").Value2 = oldB
        If Len(oldKey) = 0 Then oldKey = "CustomerID"
        .Range("B10").Value2 = oldKey
        .Range("B6:B10").Interior.Color = RGB(242, 247, 252)
        .Range("B6:B10").Borders.LineStyle = xlContinuous
        .Range("B6:B10").Borders.Color = RGB(218, 228, 238)
        .Rows("6:10").RowHeight = 27
        .Range("A14").Value2 = "Status"
        .Range("B14").Value2 = "Ready - choose both CSV files"
        .Range("A14").Font.Bold = True
        .Range("B14").Font.Color = RGB(45, 94, 143)
        .Range("A17:D18").Merge
        .Range("A17").Value2 = "The report opens in a separate workbook. Original files are never saved."
        .Range("A17").WrapText = True
        .Range("A17").Font.Color = RGB(95, 100, 110)
        .Range("A20").Value2 = "Version 0.2 UI preview | Windows Excel Desktop | VBA"
        .Range("A20").Font.Color = RGB(95, 100, 110)
    End With

    AddUiButton ws, "Browse A", "DashboardBrowseA", ws.Range("D6"), RGB(39, 100, 159)
    AddUiButton ws, "Browse B", "DashboardBrowseB", ws.Range("D8"), RGB(39, 100, 159)
    AddUiButton ws, "Run reconciliation", "DashboardRun", ws.Range("B12"), RGB(31, 125, 85)
    AddUiButton ws, "Save report", "DashboardSaveReport", ws.Range("D12"), RGB(39, 100, 159)
    ws.Activate
    ws.Range("B6").Select
    Exit Sub

Failed:
    MsgBox "Cannot build dashboard: " & Err.Description, vbExclamation
End Sub

Private Sub AddUiButton(ByVal ws As Worksheet, ByVal caption As String, _
                        ByVal procedureName As String, ByVal anchor As Range, _
                        ByVal fillColor As Long)
    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, anchor.Left, _
                                anchor.Top, anchor.Width, 28)
    With btn
        .Name = "recon_" & procedureName
        .Fill.ForeColor.RGB = fillColor
        .Line.Visible = msoFalse
        .TextFrame.Characters.Text = caption
        .TextFrame.Characters.Font.Color = vbWhite
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Size = 10
        .OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & procedureName
        .Placement = xlMoveAndSize
    End With
End Sub

Public Sub DashboardBrowseA()
    SelectCsvInto "B6"
End Sub

Public Sub DashboardBrowseB()
    SelectCsvInto "B8"
End Sub

Private Sub SelectCsvInto(ByVal destination As String)
    Dim selected As Variant
    selected = Application.GetOpenFilename("CSV files (*.csv),*.csv", , "Select CSV file")
    If VarType(selected) = vbBoolean Then Exit Sub
    ThisWorkbook.Worksheets(DASH_SHEET).Range(destination).Value2 = CStr(selected)
    SetDashboardStatus "Ready to reconcile"
End Sub

Public Sub DashboardRun()
    Dim ws As Worksheet, sourceA As String, sourceB As String, keyColumn As String
    Dim candidate As Workbook

    On Error GoTo Failure
    Set ws = ThisWorkbook.Worksheets(DASH_SHEET)
    sourceA = Trim$(CStr(ws.Range("B6").Value2))
    sourceB = Trim$(CStr(ws.Range("B8").Value2))
    keyColumn = Trim$(CStr(ws.Range("B10").Value2))
    If Len(sourceA) = 0 Or Len(sourceB) = 0 Or Len(keyColumn) = 0 Then
        MsgBox "Provide both CSV paths and a key column.", vbInformation
        Exit Sub
    End If
    If Len(Dir$(sourceA)) = 0 Or Len(Dir$(sourceB)) = 0 Then
        MsgBox "At least one source file does not exist.", vbExclamation
        Exit Sub
    End If

    Set gLastReport = Nothing
    SetDashboardStatus "Running..."
    ReconcileCsvFiles sourceA, sourceB, keyColumn

    ' The existing v0.1 engine creates a new report workbook and activates it
    ' on success. It catches its own errors; confirm the expected worksheets.
    Set candidate = ActiveWorkbook
    If Not candidate Is Nothing Then
        If Not candidate Is ThisWorkbook Then
            If IsReconciliationReport(candidate) Then Set gLastReport = candidate
        End If
    End If

    If gLastReport Is Nothing Then
        SetDashboardStatus "No report generated - check error message"
    Else
        SetDashboardStatus "Completed - report workbook is open"
        gLastReport.Activate
    End If
    Exit Sub

Failure:
    SetDashboardStatus "Error: " & Err.Description
    MsgBox "Dashboard error: " & Err.Description, vbExclamation
End Sub

Private Function IsReconciliationReport(ByVal book As Workbook) As Boolean
    Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet
    On Error Resume Next
    Set ws1 = book.Worksheets("Summary")
    Set ws2 = book.Worksheets("Differences")
    Set ws3 = book.Worksheets("Data Issues")
    On Error GoTo 0
    IsReconciliationReport = Not ws1 Is Nothing And _
                             Not ws2 Is Nothing And Not ws3 Is Nothing
End Function

Public Sub DashboardSaveReport()
    Dim filename As Variant
    Dim displayName As String
    On Error GoTo Failure

    If gLastReport Is Nothing Then
        MsgBox "Run reconciliation first. There is no report to save.", vbInformation
        Exit Sub
    End If
    If Not IsWorkbookOpen(gLastReport) Then
        Set gLastReport = Nothing
        MsgBox "The report was closed. Run reconciliation again.", vbInformation
        Exit Sub
    End If
    displayName = "Reconciliation_" & Format$(Now, "yyyymmdd_hhnnss") & ".xlsx"
    filename = Application.GetSaveAsFilename(displayName, _
                   "Excel Workbook (*.xlsx),*.xlsx", , "Save reconciliation report")
    If VarType(filename) = vbBoolean Then Exit Sub

    Application.DisplayAlerts = False
    gLastReport.SaveAs Filename:=CStr(filename), FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    SetDashboardStatus "Report saved: " & CStr(filename)
    MsgBox "Report saved successfully.", vbInformation
    Exit Sub

Failure:
    Application.DisplayAlerts = True
    MsgBox "Could not save the report: " & Err.Description, vbExclamation
End Sub

Private Function IsWorkbookOpen(ByVal book As Workbook) As Boolean
    Dim candidate As Workbook
    On Error GoTo NotOpen
    For Each candidate In Application.Workbooks
        If candidate Is book Then
            IsWorkbookOpen = True
            Exit Function
        End If
    Next candidate
NotOpen:
End Function

Private Sub SetDashboardStatus(ByVal message As String)
    On Error Resume Next
    ThisWorkbook.Worksheets(DASH_SHEET).Range("B14").Value2 = message
    On Error GoTo 0
End Sub
