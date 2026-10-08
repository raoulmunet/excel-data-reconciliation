Attribute VB_Name = "modUnifiedDashboard"
Option Explicit

' v0.5 reporting integration (KPI/formatting/run history). Windows Excel 2019 target.
' Requires modReportPresentation.bas in addition to the existing modules.
' Keep modReconciliation.bas, modDashboard.bas and modAdvancedReconciliation.bas
' imported. This module is separate to preserve the tested v0.2 UI.
Private Const UI_SHEET As String = "Reconciliation v0.4"
Private mReport As Workbook

Public Sub BuildUnifiedDashboard()
    Dim ws As Worksheet, n As Long
    Dim a As String, b As String, keys As String, mode As String
    On Error GoTo Failed
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(UI_SHEET)
    On Error GoTo Failed
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = UI_SHEET
    Else
        a = CStr(ws.Range("B6").Value2)
        b = CStr(ws.Range("B8").Value2)
        keys = CStr(ws.Range("B10").Value2)
        mode = CStr(ws.Range("B12").Value2)
    End If
    For n = ws.Shapes.Count To 1 Step -1
        ws.Shapes(n).Delete
    Next n
    ws.Cells.Clear
    With ws
        .Cells.Font.Name = "Calibri"
        .Cells.Font.Size = 11
        .Columns("A").ColumnWidth = 25
        .Columns("B").ColumnWidth = 55
        .Columns("C").ColumnWidth = 4
        .Columns("D").ColumnWidth = 23
        .Columns("E").ColumnWidth = 4
        .Range("A1:D2").Merge
        .Range("A1").Value2 = "DATA RECONCILIATION | v0.4"
        .Range("A1:D2").Interior.Color = RGB(26, 52, 82)
        .Range("A1:D2").Font.Color = vbWhite
        .Range("A1:D2").Font.Bold = True
        .Range("A1:D2").Font.Size = 18
        .Range("A1:D2").VerticalAlignment = xlCenter
        .Range("A4").Value2 = "Compare UTF-8 CSV sources using one or more key fields."
        .Range("A6").Value2 = "Source A (CSV)"
        .Range("A8").Value2 = "Source B (CSV)"
        .Range("A10").Value2 = "Key columns"
        .Range("A12").Value2 = "Comparison mode"
        .Range("B6").Value2 = a
        .Range("B8").Value2 = b
        If Len(Trim$(keys)) = 0 Then keys = "CompanyID,InvoiceNo"
        .Range("B10").Value2 = keys
        If Not ValidMode(mode) Then mode = "EXACT"
        .Range("B12").Value2 = mode
        .Range("B6:B12").Interior.Color = RGB(239, 246, 252)
        .Range("B6:B12").Borders.LineStyle = xlContinuous
        .Range("B6:B12").Borders.Color = RGB(205, 220, 235)
        .Rows("6:12").RowHeight = 28
        .Range("B12").Validation.Delete
        .Range("B12").Validation.Add Type:=xlValidateList, _
              AlertStyle:=xlValidAlertStop, Formula1:="EXACT,TRIM,IGNORE_CASE"
        .Range("B12").Validation.InCellDropdown = True
        .Range("B12").Validation.ErrorTitle = "Unsupported comparison mode"
        .Range("B12").Validation.ErrorMessage = "Choose EXACT, TRIM or IGNORE_CASE."
        .Range("B12").Validation.ShowError = True
        .Range("A16").Value2 = "Status"
        .Range("A16").Font.Bold = True
        .Range("B16").Value2 = "Ready - select source files"
        .Range("A19:D20").Merge
        .Range("A19").Value2 = "EXACT: strict | TRIM: ignore surrounding spaces | IGNORE_CASE: trim and ignore letter case"
        .Range("A19").WrapText = True
        .Range("A19").Font.Color = RGB(80, 96, 112)
        .Range("A22:D23").Merge
        .Range("A22").Value2 = "Export XLSX preserves all three report sheets. PDF exports them together."
        .Range("A22").WrapText = True
        .Range("A22").Font.Color = RGB(80, 96, 112)
        .Range("A25").Value2 = "v0.4 preview | Windows Excel 2019 | manual verification pending"
        .Range("A25").Font.Color = RGB(100, 100, 100)
    End With
    UiButton ws, "Browse A", "UnifiedBrowseA", ws.Range("D6"), RGB(35, 103, 161)
    UiButton ws, "Browse B", "UnifiedBrowseB", ws.Range("D8"), RGB(35, 103, 161)
    UiButton ws, "RUN COMPARISON", "UnifiedRun", ws.Range("B14"), RGB(27, 125, 87)
    UiButton ws, "Export XLSX", "UnifiedSaveXlsx", ws.Range("B18"), RGB(35, 103, 161)
    UiButton ws, "Export PDF", "UnifiedSavePdf", ws.Range("D18"), RGB(35, 103, 161)
    ws.Activate
    ws.Range("B6").Select
    Exit Sub
Failed:
    MsgBox "Cannot create dashboard: " & Err.Description, vbExclamation
End Sub

Private Sub UiButton(ByVal ws As Worksheet, ByVal caption As String, _
                     ByVal macroName As String, ByVal anchor As Range, ByVal color As Long)
    Dim shape As Shape
    Set shape = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                      anchor.Left, anchor.Top, anchor.Width, 27)
    With shape
        .Name = "unified_" & macroName
        .Fill.ForeColor.RGB = color
        .Line.Visible = msoFalse
        .TextFrame.Characters.Text = caption
        .TextFrame.Characters.Font.Color = vbWhite
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Size = 10
        .OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macroName
        .Placement = xlMoveAndSize
    End With
End Sub

Public Sub UnifiedBrowseA()
    UiBrowse "B6"
End Sub

Public Sub UnifiedBrowseB()
    UiBrowse "B8"
End Sub

Private Sub UiBrowse(ByVal address As String)
    Dim selected As Variant
    selected = Application.GetOpenFilename("CSV files (*.csv),*.csv", , _
                                          "Select a CSV data source")
    If VarType(selected) = vbBoolean Then Exit Sub
    ThisWorkbook.Worksheets(UI_SHEET).Range(address).Value2 = CStr(selected)
    UiStatus "Source selected"
End Sub

Public Sub UnifiedRun()
    Dim ws As Worksheet
    Dim a As String, b As String, keys As String, mode As String
    Dim oldBook As Workbook, candidate As Workbook
    Dim oldCount As Long
    Dim startTick As Double, elapsed As Double, presentationWarning As String
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets(UI_SHEET)
    a = Trim$(CStr(ws.Range("B6").Value2))
    b = Trim$(CStr(ws.Range("B8").Value2))
    keys = Trim$(CStr(ws.Range("B10").Value2))
    mode = UCase$(Trim$(CStr(ws.Range("B12").Value2)))
    If Len(a) = 0 Or Len(b) = 0 Or Len(keys) = 0 Then
        MsgBox "Select both CSV files and enter key column(s).", vbInformation
        Exit Sub
    End If
    If Not ValidMode(mode) Then
        MsgBox "Choose EXACT, TRIM or IGNORE_CASE.", vbInformation
        Exit Sub
    End If
    If StrComp(a, b, vbTextCompare) = 0 Then
        MsgBox "Select two different source files.", vbExclamation
        Exit Sub
    End If
    If Len(Dir$(a)) = 0 Or Len(Dir$(b)) = 0 Then
        MsgBox "At least one CSV file was not found.", vbExclamation
        Exit Sub
    End If

    ' Keep the previous valid report if the next run fails.
    Set oldBook = Nothing
    If Not mReport Is Nothing Then
        If WorkbookIsOpen(mReport) Then Set oldBook = mReport
    End If
    Set mReport = Nothing
    oldCount = Application.Workbooks.Count
    startTick = Timer
    UiStatus "Processing..."
    ' The existing engine handles exceptions internally and shows its own dialog.
    AdvancedReconcileCsvFiles a, b, keys, mode
    If Application.Workbooks.Count > oldCount Then
        Set candidate = Application.ActiveWorkbook
        If Not candidate Is Nothing Then
            If Not candidate Is ThisWorkbook Then
                If ValidReport(candidate, a, b, keys) Then Set mReport = candidate
            End If
        End If
    End If

    elapsed = Timer - startTick
    If elapsed < 0 Then elapsed = elapsed + 86400#
    If mReport Is Nothing Then
        Set mReport = oldBook
        LogReconciliationRun a, b, keys, mode, "FAILED / NO REPORT", elapsed, _
                             "Reconciliation produced no validated new workbook."
        UiStatus "Run did not produce a new report; inspect the error dialog"
    Else
        On Error Resume Next
        EnhanceReport mReport
        If Err.Number <> 0 Then
            presentationWarning = Err.Description
            Err.Clear
        End If
        On Error GoTo Failed
        If Len(presentationWarning) > 0 Then
            LogReconciliationRun a, b, keys, mode, "SUCCESS / STYLE WARNING", _
                                 elapsed, presentationWarning
            UiStatus "Reconciled; report styling warning - see Run History"
        Else
            LogReconciliationRun a, b, keys, mode, "SUCCESS", elapsed
            UiStatus "Completed - KPI report ready"
        End If
        mReport.Activate
    End If
    Exit Sub
Failed:
    Set mReport = oldBook
    UiStatus "Error: " & Err.Description
    MsgBox "Dashboard error: " & Err.Description, vbExclamation
End Sub

Private Function ValidReport(ByVal book As Workbook, ByVal a As String, _
                             ByVal b As String, ByVal keys As String) As Boolean
    Dim summary As Worksheet, diffs As Worksheet, issues As Worksheet
    On Error GoTo Invalid
    Set summary = book.Worksheets("Summary")
    Set diffs = book.Worksheets("Differences")
    Set issues = book.Worksheets("Data Issues")
    ValidReport = (CStr(summary.Range("B2").Value2) = a And _
                   CStr(summary.Range("B3").Value2) = b And _
                   CStr(summary.Range("B4").Value2) = keys)
    Exit Function
Invalid:
    ValidReport = False
End Function

Private Function ValidMode(ByVal value As String) As Boolean
    Select Case UCase$(Trim$(value))
        Case "EXACT", "TRIM", "IGNORE_CASE": ValidMode = True
    End Select
End Function

Private Function WorkbookIsOpen(ByVal book As Workbook) As Boolean
    Dim item As Workbook
    On Error GoTo Closed
    For Each item In Application.Workbooks
        If item Is book Then
            WorkbookIsOpen = True
            Exit Function
        End If
    Next item
Closed:
End Function

Private Function ExportWorkbook() As Workbook
    If mReport Is Nothing Then
        MsgBox "Run a reconciliation first.", vbInformation
    ElseIf Not WorkbookIsOpen(mReport) Then
        Set mReport = Nothing
        MsgBox "The last report has been closed. Run reconciliation again.", vbInformation
    Else
        Set ExportWorkbook = mReport
    End If
End Function

Public Sub UnifiedSaveXlsx()
    Dim report As Workbook, target As Variant
    Dim previousAlerts As Boolean
    On Error GoTo Failed
    Set report = ExportWorkbook()
    If report Is Nothing Then Exit Sub
    target = Application.GetSaveAsFilename( _
           "Reconciliation_" & Format$(Now, "yyyymmdd_hhnnss") & ".xlsx", _
           "Excel Workbook (*.xlsx),*.xlsx", , "Export Excel report")
    If VarType(target) = vbBoolean Then Exit Sub
    previousAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    report.SaveAs Filename:=CStr(target), FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = previousAlerts
    UiStatus "Excel report exported"
    MsgBox "Excel report saved.", vbInformation
    Exit Sub
Failed:
    Application.DisplayAlerts = previousAlerts
    MsgBox "Excel export failed: " & Err.Description, vbExclamation
End Sub

Public Sub UnifiedSavePdf()
    Dim report As Workbook, target As Variant
    On Error GoTo Failed
    Set report = ExportWorkbook()
    If report Is Nothing Then Exit Sub
    target = Application.GetSaveAsFilename( _
           "Reconciliation_" & Format$(Now, "yyyymmdd_hhnnss") & ".pdf", _
           "PDF files (*.pdf),*.pdf", , "Export PDF report")
    If VarType(target) = vbBoolean Then Exit Sub
    report.ExportAsFixedFormat Type:=xlTypePDF, Filename:=CStr(target), _
                  Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                  IgnorePrintAreas:=False, OpenAfterPublish:=False
    UiStatus "PDF report exported"
    MsgBox "PDF report saved.", vbInformation
    Exit Sub
Failed:
    MsgBox "PDF export failed: " & Err.Description, vbExclamation
End Sub

Private Sub UiStatus(ByVal message As String)
    On Error Resume Next
    ThisWorkbook.Worksheets(UI_SHEET).Range("B16").Value2 = message
    On Error GoTo 0
End Sub
