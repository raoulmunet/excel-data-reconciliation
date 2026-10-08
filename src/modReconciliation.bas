Attribute VB_Name = "modReconciliation"
Option Explicit

' Excel Data Reconciliation v0.1
' Reference runtime: Windows 11 + Excel 2019 Desktop (testing pending).
' Late binding; no external library reference required.

Public Sub RunReconciliation()
    Dim firstPath As Variant, secondPath As Variant
    Dim keyName As String
    firstPath = Application.GetOpenFilename("CSV files (*.csv),*.csv", , "Select Source A CSV")
    If VarType(firstPath) = vbBoolean Then Exit Sub
    secondPath = Application.GetOpenFilename("CSV files (*.csv),*.csv", , "Select Source B CSV")
    If VarType(secondPath) = vbBoolean Then Exit Sub
    keyName = Trim$(InputBox("Unique key column name (e.g. CustomerID):", _
                              "Reconciliation", "CustomerID"))
    If Len(keyName) = 0 Then Exit Sub
    ReconcileCsvFiles CStr(firstPath), CStr(secondPath), keyName
End Sub

Public Sub ReconcileCsvFiles(ByVal pathA As String, ByVal pathB As String, _
                             ByVal keyColumn As String)
    Dim bookA As Workbook, bookB As Workbook, report As Workbook
    Dim a As Variant, b As Variant
    Dim headersA As Object, headersB As Object
    Dim indexesA As Object, indexesB As Object
    Dim duplicatesA As Object, duplicatesB As Object
    Dim outDiff As Worksheet, outIssues As Worksheet, outSummary As Worksheet
    Dim keyA As Long, keyB As Long, r As Long, c As Long
    Dim rr As Long, ir As Long, aOnly As Long, bOnly As Long
    Dim changed As Long, equalCount As Long, issueCount As Long
    Dim k As Variant, h As Variant, rowA As Long, rowB As Long
    Dim originalValue As String, updatedValue As String
    Dim oldScreen As Boolean, oldEvents As Boolean
    Dim previousAlerts As Boolean
    Dim errNumber As Long, errText As String
    On Error GoTo Failure

    If Len(Trim$(pathA)) = 0 Or Len(Trim$(pathB)) = 0 Then _
       Err.Raise vbObjectError + 110, , "Both CSV paths are required."
    If StrComp(pathA, pathB, vbTextCompare) = 0 Then _
       Err.Raise vbObjectError + 111, , "Choose two different source files."

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    previousAlerts = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Set bookA = OpenCsvWorkbook(pathA)
    Set bookB = OpenCsvWorkbook(pathB)
    a = RangeAsMatrix(bookA.Worksheets(1))
    b = RangeAsMatrix(bookB.Worksheets(1))
    Set headersA = BuildHeaders(a)
    Set headersB = BuildHeaders(b)
    If headersA.Count <> headersB.Count Then _
        Err.Raise vbObjectError + 112, , "Source headers do not match."
    For Each h In headersA.Keys
        If Not headersB.Exists(CStr(h)) Then _
            Err.Raise vbObjectError + 113, , "Missing column in B: " & CStr(h)
    Next h
    keyColumn = Trim$(keyColumn)
    If Not headersA.Exists(keyColumn) Then _
        Err.Raise vbObjectError + 114, , "Key column does not exist: " & keyColumn
    keyA = CLng(headersA(keyColumn))
    keyB = CLng(headersB(keyColumn))

    Set indexesA = CreateObject("Scripting.Dictionary")
    Set indexesB = CreateObject("Scripting.Dictionary")
    Set duplicatesA = CreateObject("Scripting.Dictionary")
    Set duplicatesB = CreateObject("Scripting.Dictionary")
    indexesA.CompareMode = vbTextCompare
    indexesB.CompareMode = vbTextCompare
    duplicatesA.CompareMode = vbTextCompare
    duplicatesB.CompareMode = vbTextCompare
    BuildIndex a, keyA, indexesA, duplicatesA
    BuildIndex b, keyB, indexesB, duplicatesB

    Set report = Workbooks.Add(xlWBATWorksheet)
    Set outSummary = report.Worksheets(1)
    outSummary.Name = "Summary"
    Set outDiff = report.Worksheets.Add(After:=outSummary)
    outDiff.Name = "Differences"
    Set outIssues = report.Worksheets.Add(After:=outDiff)
    outIssues.Name = "Data Issues"

    outDiff.Range("A1:E1").Value = Array("Key", "Status", "Column", "Source A", "Source B")
    outIssues.Range("A1:C1").Value = Array("Source", "Issue", "Key")
    rr = 2
    ir = 2

    For Each k In duplicatesA.Keys
        WriteIssue outIssues, ir, "A", "Duplicate or blank key", CStr(k)
    Next k
    For Each k In duplicatesB.Keys
        WriteIssue outIssues, ir, "B", "Duplicate or blank key", CStr(k)
    Next k
    issueCount = ir - 2

    For Each k In indexesA.Keys
        If Not duplicatesA.Exists(k) And Not duplicatesB.Exists(k) Then
            If Not indexesB.Exists(k) Then
                WriteDifference outDiff, rr, CStr(k), "Only in A", "", "", ""
                aOnly = aOnly + 1
            Else
                rowA = CLng(indexesA(k))
                rowB = CLng(indexesB(k))
                Dim rowChanged As Boolean
                rowChanged = False
                For Each h In headersA.Keys
                    If StrComp(CStr(h), keyColumn, vbTextCompare) <> 0 Then
                        originalValue = SafeText(a(rowA, CLng(headersA(h))))
                        updatedValue = SafeText(b(rowB, CLng(headersB(h))))
                        If StrComp(originalValue, updatedValue, vbBinaryCompare) <> 0 Then
                            WriteDifference outDiff, rr, CStr(k), "Changed", _
                                            CStr(h), originalValue, updatedValue
                            changed = changed + 1
                            rowChanged = True
                        End If
                    End If
                Next h
                If Not rowChanged Then equalCount = equalCount + 1
            End If
        End If
    Next k

    For Each k In indexesB.Keys
        If Not duplicatesA.Exists(k) And Not duplicatesB.Exists(k) Then
            If Not indexesA.Exists(k) Then
                WriteDifference outDiff, rr, CStr(k), "Only in B", "", "", ""
                bOnly = bOnly + 1
            End If
        End If
    Next k

    With outSummary
        .Range("A1:B1").Value = Array("Metric", "Value")
        .Cells(2, 1).Value = "Source A": .Cells(2, 2).Value = pathA
        .Cells(3, 1).Value = "Source B": .Cells(3, 2).Value = pathB
        .Cells(4, 1).Value = "Key column": .Cells(4, 2).Value = keyColumn
        .Cells(5, 1).Value = "Only in A": .Cells(5, 2).Value = aOnly
        .Cells(6, 1).Value = "Only in B": .Cells(6, 2).Value = bOnly
        .Cells(7, 1).Value = "Field differences": .Cells(7, 2).Value = changed
        .Cells(8, 1).Value = "Matching records unchanged": .Cells(8, 2).Value = equalCount
        .Cells(9, 1).Value = "Data issues": .Cells(9, 2).Value = issueCount
        .Cells(10, 1).Value = "Generated at": .Cells(10, 2).Value = Now
        .Cells(10, 2).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    End With
    StyleReport outSummary
    StyleReport outDiff
    StyleReport outIssues

CleanExit:
    On Error Resume Next
    If Not bookA Is Nothing Then bookA.Close SaveChanges:=False
    If Not bookB Is Nothing Then bookB.Close SaveChanges:=False
    Application.DisplayAlerts = previousAlerts
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    On Error GoTo 0
    Exit Sub

Failure:
    errNumber = Err.Number
    errText = Err.Description
    If Not report Is Nothing Then report.Close SaveChanges:=False
    MsgBox "Reconciliation failed (" & errNumber & "): " & errText, _
           vbExclamation, "Excel Data Reconciliation"
    Resume CleanExit
End Sub

Private Function OpenCsvWorkbook(ByVal csvPath As String) As Workbook
    ' Comma-delimited native Excel CSV import; may coerce numbers/dates.
    Workbooks.OpenText Filename:=csvPath, Origin:=xlWindows, _
                       DataType:=xlDelimited, Comma:=True, _
                       Semicolon:=False, Tab:=False, Local:=False
    Set OpenCsvWorkbook = ActiveWorkbook
End Function

Private Function RangeAsMatrix(ByVal ws As Worksheet) As Variant
    Dim data As Variant
    Dim cells As Range
    Set cells = ws.UsedRange
    If cells.Row <> 1 Or cells.Column <> 1 Then _
        Err.Raise vbObjectError + 115, , "CSV must start at cell A1."
    data = cells.Value2
    If Not IsArray(data) Then _
        Err.Raise vbObjectError + 116, , "CSV must contain a header and data rows."
    If UBound(data, 1) < 2 Then _
        Err.Raise vbObjectError + 117, , "CSV contains no data rows."
    RangeAsMatrix = data
End Function

Private Function BuildHeaders(ByRef data As Variant) As Object
    Dim result As Object, c As Long, name As String
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    For c = 1 To UBound(data, 2)
        name = Trim$(SafeText(data(1, c)))
        If Len(name) = 0 Then Err.Raise vbObjectError + 118, , "Empty header."
        If result.Exists(name) Then _
           Err.Raise vbObjectError + 119, , "Duplicate header: " & name
        result.Add name, c
    Next c
    Set BuildHeaders = result
End Function

Private Sub BuildIndex(ByRef data As Variant, ByVal keyCol As Long, _
                       ByVal idx As Object, ByVal dups As Object)
    Dim r As Long, c As Long, key As String, isEmpty As Boolean
    For r = 2 To UBound(data, 1)
        isEmpty = True
        For c = 1 To UBound(data, 2)
            If Len(Trim$(SafeText(data(r, c)))) > 0 Then
                isEmpty = False
                Exit For
            End If
        Next c
        If Not isEmpty Then
            key = Trim$(SafeText(data(r, keyCol)))
            If Len(key) = 0 Then
                If Not dups.Exists("<EMPTY KEY>") Then dups.Add "<EMPTY KEY>", True
            ElseIf idx.Exists(key) Then
                If Not dups.Exists(key) Then dups.Add key, True
            Else
                idx.Add key, r
            End If
        End If
    Next r
End Sub

Private Function SafeText(ByVal value As Variant) As String
    If IsError(value) Then
        SafeText = "#CELL_ERROR"
    ElseIf IsNull(value) Or IsEmpty(value) Then
        SafeText = ""
    Else
        SafeText = CStr(value)
    End If
End Function

Private Sub WriteDifference(ByVal ws As Worksheet, ByRef nextRow As Long, _
                            ByVal key As String, ByVal status As String, _
                            ByVal field As String, ByVal oldValue As String, _
                            ByVal newValue As String)
    ws.Cells(nextRow, 1).Resize(1, 5).NumberFormat = "@"
    ws.Cells(nextRow, 1).Value2 = key
    ws.Cells(nextRow, 2).Value2 = status
    ws.Cells(nextRow, 3).Value2 = field
    ws.Cells(nextRow, 4).Value2 = oldValue
    ws.Cells(nextRow, 5).Value2 = newValue
    nextRow = nextRow + 1
End Sub

Private Sub WriteIssue(ByVal ws As Worksheet, ByRef nextRow As Long, _
                       ByVal source As String, ByVal issue As String, ByVal key As String)
    ws.Cells(nextRow, 1).Value2 = source
    ws.Cells(nextRow, 2).Value2 = issue
    ws.Cells(nextRow, 3).NumberFormat = "@"
    ws.Cells(nextRow, 3).Value2 = key
    nextRow = nextRow + 1
End Sub

Private Sub StyleReport(ByVal ws As Worksheet)
    With ws.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(222, 235, 247)
    End With
    ws.Columns.AutoFit
    ws.Rows(1).AutoFilter
End Sub
