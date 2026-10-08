Attribute VB_Name = "modReportPresentation"
Option Explicit

' v0.5 reporting add-on: KPI dashboard, report formatting, run history.
' Reference environment: Windows 11 + Excel 2019 Desktop. Runtime untested.
' Import alongside modUnifiedDashboard and modAdvancedReconciliation.
Private Const HISTORY_SHEET As String = "Run History"

Public Sub EnhanceReport(ByVal report As Workbook)
    Dim summary As Worksheet, differences As Worksheet, issues As Worksheet
    Dim kpi As Worksheet, onlyA As Double, onlyB As Double, changed As Double
    Dim unchanged As Double, issuesCount As Double, matches As Double
    Dim compared As Double, matchRate As Double
    On Error GoTo Failed

    Set summary = report.Worksheets("Summary")
    Set differences = report.Worksheets("Differences")
    Set issues = report.Worksheets("Data Issues")
    onlyA = CDbl(summary.Range("B6").Value2)
    onlyB = CDbl(summary.Range("B7").Value2)
    changed = CDbl(summary.Range("B8").Value2)
    unchanged = CDbl(summary.Range("B9").Value2)
    issuesCount = CDbl(summary.Range("B10").Value2)

    ' Count unique matched business keys with differences, not field differences.
    matches = CountChangedKeys(differences)
    compared = unchanged + matches
    If compared > 0 Then matchRate = unchanged / compared Else matchRate = 0

    On Error Resume Next
    Set kpi = report.Worksheets("KPI Dashboard")
    On Error GoTo Failed
    If kpi Is Nothing Then
        Set kpi = report.Worksheets.Add(Before:=report.Worksheets(1))
        kpi.Name = "KPI Dashboard"
    Else
        kpi.Cells.Clear
    End If

    With kpi
        .Cells.Font.Name = "Calibri"
        .Columns("A").ColumnWidth = 4
        .Columns("B").ColumnWidth = 27
        .Columns("C").ColumnWidth = 18
        .Columns("D").ColumnWidth = 27
        .Columns("E").ColumnWidth = 18
        .Columns("F").ColumnWidth = 4
        .Range("B2:E3").Merge
        .Range("B2").Value2 = "RECONCILIATION | RESULTS"
        .Range("B2:E3").Interior.Color = RGB(28, 53, 83)
        .Range("B2:E3").Font.Color = vbWhite
        .Range("B2:E3").Font.Bold = True
        .Range("B2:E3").Font.Size = 18
        .Range("B2:E3").VerticalAlignment = xlCenter

        .Range("B5").Value2 = "Matched keys"
        .Range("C5").Value2 = compared
        .Range("D5").Value2 = "Match rate"
        .Range("E5").Value2 = matchRate
        .Range("E5").NumberFormat = "0.0%"
        .Range("B7").Value2 = "Unchanged keys"
        .Range("C7").Value2 = unchanged
        .Range("D7").Value2 = "Changed keys"
        .Range("E7").Value2 = matches
        .Range("B9").Value2 = "Field differences"
        .Range("C9").Value2 = changed
        .Range("D9").Value2 = "Data issues"
        .Range("E9").Value2 = issuesCount
        .Range("B11").Value2 = "Only in A"
        .Range("C11").Value2 = onlyA
        .Range("D11").Value2 = "Only in B"
        .Range("E11").Value2 = onlyB
        .Range("B5:E11").Borders.Color = RGB(216, 226, 236)
        .Range("B5:E11").Borders.LineStyle = xlContinuous
        .Range("B5:E11").RowHeight = 28
        .Range("B5:B11").Font.Bold = True
        .Range("D5:D11").Font.Bold = True
        .Range("C5:C11").Font.Size = 13
        .Range("E5:E11").Font.Size = 13
        .Range("B14").Value2 = "Sources"
        .Range("B15").Value2 = "Source A"
        .Range("C15").Value2 = CStr(summary.Range("B2").Value2)
        .Range("B16").Value2 = "Source B"
        .Range("C16").Value2 = CStr(summary.Range("B3").Value2)
        .Range("B17").Value2 = "Key columns"
        .Range("C17").Value2 = CStr(summary.Range("B4").Value2)
        .Range("B18").Value2 = "Comparison"
        .Range("C18").Value2 = CStr(summary.Range("B5").Value2)
        .Range("B21:E22").Merge
        .Range("B21").Value2 = "Match rate = unchanged matched keys / all matched keys. " & _
           "Unmatched and duplicate keys are excluded from this rate."
        .Range("B21").WrapText = True
        .Range("B21").Font.Color = RGB(91, 100, 115)
        .Range("B21:E22").VerticalAlignment = xlCenter
        .Range("B14").Font.Bold = True
        .Range("B14").Font.Size = 13
        .Range("C15:E18").WrapText = True
        .Rows("15:18").RowHeight = 26
    End With

    StyleDataSheet summary, 2
    StyleDataSheet differences, 5
    StyleDataSheet issues, 3
    SetupPrintPage kpi, True
    SetupPrintPage summary, True
    SetupPrintPage differences, False
    SetupPrintPage issues, False
    kpi.Activate
    Exit Sub
Failed:
    Err.Raise vbObjectError + 701, "EnhanceReport", _
              "Could not prepare report presentation: " & Err.Description
End Sub

Private Function CountChangedKeys(ByVal ws As Worksheet) As Long
    Dim dict As Object, last As Long, r As Long, value As String
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbBinaryCompare
    last = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    For r = 2 To last
        If StrComp(CStr(ws.Cells(r, 2).Value2), "Changed", vbBinaryCompare) = 0 Then
            value = CStr(ws.Cells(r, 1).Value2)
            If Not dict.Exists(value) Then dict.Add value, True
        End If
    Next r
    CountChangedKeys = dict.Count
End Function

Private Sub StyleDataSheet(ByVal ws As Worksheet, ByVal headerColumns As Long)
    Dim row As Long, last As Long, kind As String
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, headerColumns))
        .Interior.Color = RGB(28, 53, 83)
        .Font.Color = vbWhite
        .Font.Bold = True
        .RowHeight = 27
    End With
    ws.Columns.AutoFit
    ws.Columns("A").ColumnWidth = WorksheetFunction.Min(38, _
         WorksheetFunction.Max(20, ws.Columns("A").ColumnWidth))
    If ws.Name = "Differences" Or ws.Name = "Data Issues" Then
        last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        For row = 2 To last
            If row Mod 2 = 0 Then
                ws.Range(ws.Cells(row, 1), ws.Cells(row, headerColumns)).Interior.Color = _
                        RGB(246, 249, 252)
            End If
            If ws.Name = "Differences" Then
                kind = CStr(ws.Cells(row, 2).Value2)
                Select Case kind
                    Case "Changed"
                        ws.Cells(row, 2).Interior.Color = RGB(255, 229, 191)
                    Case "Only in A"
                        ws.Cells(row, 2).Interior.Color = RGB(255, 214, 214)
                    Case "Only in B"
                        ws.Cells(row, 2).Interior.Color = RGB(214, 238, 222)
                End Select
            Else
                ws.Cells(row, 2).Interior.Color = RGB(255, 229, 191)
            End If
        Next row
    End If
End Sub

Private Sub SetupPrintPage(ByVal ws As Worksheet, ByVal portrait As Boolean)
    With ws.PageSetup
        .Orientation = IIf(portrait, xlPortrait, xlLandscape)
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = IIf(ws.Name = "KPI Dashboard", "", "$1:$1")
        .CenterHeader = "Excel Data Reconciliation"
        .RightFooter = "Page &P / &N"
        .LeftFooter = "&D"
    End With
    If ws.Name = "KPI Dashboard" Then
        ws.PageSetup.PrintArea = "$B$2:$E$22"
    End If
End Sub

Public Sub LogReconciliationRun(ByVal sourceA As String, ByVal sourceB As String, _
                                 ByVal keyColumns As String, ByVal mode As String, _
                                 ByVal result As String, ByVal elapsedSeconds As Double, _
                                 Optional ByVal notes As String = "")
    Dim ws As Worksheet, r As Long
    On Error GoTo Failed
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(HISTORY_SHEET)
    On Error GoTo Failed
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:= _
                 ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = HISTORY_SHEET
        ws.Range("A1:H1").Value = Array("Timestamp", "Status", "Source A", _
              "Source B", "Key columns", "Mode", "Seconds", "Notes")
        With ws.Rows(1)
            .Font.Bold = True
            .Interior.Color = RGB(222, 235, 247)
        End With
        ws.Columns("A").ColumnWidth = 21
        ws.Columns("B").ColumnWidth = 23
        ws.Columns("C:D").ColumnWidth = 41
        ws.Columns("E:F").ColumnWidth = 22
        ws.Columns("G").ColumnWidth = 12
        ws.Columns("H").ColumnWidth = 56
    End If
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value2 = Now
    ws.Cells(r, 1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    ws.Cells(r, 2).Value2 = result
    ws.Cells(r, 3).Value2 = sourceA
    ws.Cells(r, 4).Value2 = sourceB
    ws.Cells(r, 5).Value2 = keyColumns
    ws.Cells(r, 6).Value2 = mode
    ws.Cells(r, 7).Value2 = Round(elapsedSeconds, 2)
    ws.Cells(r, 8).Value2 = notes
    Exit Sub
Failed:
    ' Logging is best-effort; it must not invalidate a successful reconciliation.
    Debug.Print "Could not record run history: " & Err.Description
End Sub
