Attribute VB_Name = "modAdvancedReconciliation"
Option Explicit

' v0.3 preview: advanced text-preserving CSV reconciliation.
' Windows Excel 2019 / VBA7 target. Not yet runtime tested.
' Import alongside modReconciliation.bas and modDashboard.bas.
' Uses late-bound ADODB.Stream (UTF-8) and Scripting.Dictionary.
' CSV values stay strings (including leading zeroes and long IDs).

Public Sub RunAdvancedReconciliation()
    Dim pa As Variant, pb As Variant, keys As String, mode As String
    pa = Application.GetOpenFilename("CSV files (*.csv),*.csv", , "Advanced: choose source A")
    If VarType(pa) = vbBoolean Then Exit Sub
    pb = Application.GetOpenFilename("CSV files (*.csv),*.csv", , "Advanced: choose source B")
    If VarType(pb) = vbBoolean Then Exit Sub
    keys = Trim$(InputBox("Comma-separated key columns (e.g. CompanyID,InvoiceNo)", _
                          "Composite keys", "CustomerID"))
    If Len(keys) = 0 Then Exit Sub
    mode = UCase$(Trim$(InputBox( _
           "Comparison: EXACT, TRIM or IGNORE_CASE", "Comparison mode", "EXACT")))
    If mode = "" Then Exit Sub
    AdvancedReconcileCsvFiles CStr(pa), CStr(pb), keys, mode
End Sub

Public Sub AdvancedReconcileCsvFiles(ByVal pathA As String, ByVal pathB As String, _
                                     ByVal keyColumns As String, _
                                     Optional ByVal comparisonMode As String = "EXACT")
    Dim a As Collection, b As Collection, ha As Object, hb As Object
    Dim ia As Object, ib As Object, da As Object, db As Object
    Dim headerA As Variant, headerB As Variant, names As Variant
    Dim keyColsA() As Long, keyColsB() As Long
    Dim report As Workbook, summary As Worksheet, diffs As Worksheet, issues As Worksheet
    Dim oldScreen As Boolean, oldEvents As Boolean
    Dim p As Long, r As Long, c As Long, rr As Long, ir As Long
    Dim onlyA As Long, onlyB As Long, changed As Long, unchanged As Long
    Dim rowA As Variant, rowB As Variant, k As Variant, field As Variant
    Dim key As String, va As String, vb As String, rowChanged As Boolean
    Dim errorCode As Long, errorMessage As String, stateSaved As Boolean

    On Error GoTo Failed
    comparisonMode = UCase$(Trim$(comparisonMode))
    If comparisonMode <> "EXACT" And comparisonMode <> "TRIM" And _
       comparisonMode <> "IGNORE_CASE" Then _
       Err.Raise vbObjectError + 501, , "Mode must be EXACT, TRIM or IGNORE_CASE."
    If Len(Trim$(pathA)) = 0 Or Len(Trim$(pathB)) = 0 Then _
       Err.Raise vbObjectError + 502, , "Both source paths are required."
    If StrComp(pathA, pathB, vbTextCompare) = 0 Then _
       Err.Raise vbObjectError + 503, , "Use two different CSV files."

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    stateSaved = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set a = ReadUtf8Csv(pathA)
    Set b = ReadUtf8Csv(pathB)
    headerA = a(1)
    headerB = b(1)
    Set ha = CreateHeaderMap(headerA)
    Set hb = CreateHeaderMap(headerB)
    If ha.Count <> hb.Count Then _
       Err.Raise vbObjectError + 504, , "Source column counts differ."
    For Each field In ha.Keys
        If Not hb.Exists(CStr(field)) Then _
            Err.Raise vbObjectError + 505, , "Column missing in source B: " & CStr(field)
    Next field

    names = Split(keyColumns, ",")
    ReDim keyColsA(LBound(names) To UBound(names))
    ReDim keyColsB(LBound(names) To UBound(names))
    For p = LBound(names) To UBound(names)
        key = Trim$(CStr(names(p)))
        If Len(key) = 0 Then Err.Raise vbObjectError + 506, , "Empty key column name."
        If Not ha.Exists(key) Then _
           Err.Raise vbObjectError + 507, , "Unknown key column: " & key
        keyColsA(p) = CLng(ha(key))
        keyColsB(p) = CLng(hb(key))
    Next p

    Set ia = CreateObject("Scripting.Dictionary")
    Set ib = CreateObject("Scripting.Dictionary")
    Set da = CreateObject("Scripting.Dictionary")
    Set db = CreateObject("Scripting.Dictionary")
    ia.CompareMode = vbBinaryCompare
    ib.CompareMode = vbBinaryCompare
    da.CompareMode = vbBinaryCompare
    db.CompareMode = vbBinaryCompare

    IndexRows a, keyColsA, ia, da
    IndexRows b, keyColsB, ib, db

    Set report = Workbooks.Add(xlWBATWorksheet)
    Set summary = report.Worksheets(1)
    summary.Name = "Summary"
    Set diffs = report.Worksheets.Add(After:=summary)
    diffs.Name = "Differences"
    Set issues = report.Worksheets.Add(After:=diffs)
    issues.Name = "Data Issues"
    diffs.Range("A1:E1").Value = Array("Key", "Status", "Column", "Source A", "Source B")
    issues.Range("A1:C1").Value = Array("Source", "Issue", "Key")
    rr = 2
    ir = 2

    For Each k In da.Keys
        AddIssue issues, ir, "A", "Duplicate or empty key", CStr(k)
    Next k
    For Each k In db.Keys
        AddIssue issues, ir, "B", "Duplicate or empty key", CStr(k)
    Next k

    For Each k In ia.Keys
        If Not da.Exists(k) And Not db.Exists(k) Then
            If Not ib.Exists(k) Then
                AddDiff diffs, rr, CStr(k), "Only in A", "", "", ""
                onlyA = onlyA + 1
            Else
                rowA = a(CLng(ia(k)))
                rowB = b(CLng(ib(k)))
                rowChanged = False
                For Each field In ha.Keys
                    If Not IsKeyField(CStr(field), names) Then
                        va = CStr(rowA(CLng(ha(field))))
                        vb = CStr(rowB(CLng(hb(field))))
                        If NormalizeValue(va, comparisonMode) <> _
                           NormalizeValue(vb, comparisonMode) Then
                            AddDiff diffs, rr, CStr(k), "Changed", CStr(field), va, vb
                            changed = changed + 1
                            rowChanged = True
                        End If
                    End If
                Next field
                If Not rowChanged Then unchanged = unchanged + 1
            End If
        End If
    Next k
    For Each k In ib.Keys
        If Not da.Exists(k) And Not db.Exists(k) Then
            If Not ia.Exists(k) Then
                AddDiff diffs, rr, CStr(k), "Only in B", "", "", ""
                onlyB = onlyB + 1
            End If
        End If
    Next k

    With summary
        .Range("A1:B1").Value = Array("Metric", "Value")
        .Cells(2, 1).Value = "Source A": .Cells(2, 2).Value = pathA
        .Cells(3, 1).Value = "Source B": .Cells(3, 2).Value = pathB
        .Cells(4, 1).Value = "Key columns": .Cells(4, 2).Value = keyColumns
        .Cells(5, 1).Value = "Comparison mode": .Cells(5, 2).Value = comparisonMode
        .Cells(6, 1).Value = "Only in A": .Cells(6, 2).Value = onlyA
        .Cells(7, 1).Value = "Only in B": .Cells(7, 2).Value = onlyB
        .Cells(8, 1).Value = "Changed fields": .Cells(8, 2).Value = changed
        .Cells(9, 1).Value = "Unchanged matched records": .Cells(9, 2).Value = unchanged
        .Cells(10, 1).Value = "Data issues": .Cells(10, 2).Value = ir - 2
        .Cells(11, 1).Value = "Generated at": .Cells(11, 2).Value = Now
        .Cells(11, 2).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    End With
    FormatSheet summary
    FormatSheet diffs
    FormatSheet issues
CleanExit:
    If stateSaved Then
        Application.ScreenUpdating = oldScreen
        Application.EnableEvents = oldEvents
    End If
    Exit Sub
Failed:
    errorCode = Err.Number
    errorMessage = Err.Description
    On Error Resume Next
    If Not report Is Nothing Then report.Close SaveChanges:=False
    On Error GoTo 0
    MsgBox "Advanced reconciliation failed (" & errorCode & "): " & _
            errorMessage, vbExclamation
    Resume CleanExit
End Sub

Private Function ReadUtf8Csv(ByVal filePath As String) As Collection
    Dim stream As Object, data As String
    Dim output As New Collection, lineValues As Collection
    Dim fields() As String, idx As Long, pos As Long, ch As String
    Dim cellText As String, inQuotes As Boolean, afterQuote As Boolean
    Dim endedRow As Boolean, fieldCount As Long
    Dim item As Variant

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    On Error GoTo ReadFailed
    stream.LoadFromFile filePath
    data = stream.ReadText(-1)
    stream.Close
    If Len(data) > 0 Then
        If AscW(Left$(data, 1)) = &HFEFF Or AscW(Left$(data, 1)) = -257 Then _
            data = Mid$(data, 2)
    End If

    Set lineValues = New Collection
    cellText = ""
    For pos = 1 To Len(data)
        ch = Mid$(data, pos, 1)
        endedRow = False
        If inQuotes Then
            If ch = """" Then
                If pos < Len(data) And Mid$(data, pos + 1, 1) = """" Then
                    cellText = cellText & """"
                    pos = pos + 1
                Else
                    inQuotes = False
                    afterQuote = True
                End If
            Else
                cellText = cellText & ch
            End If
        Else
            Select Case ch
                Case """"
                    If Len(cellText) = 0 And Not afterQuote Then
                        inQuotes = True
                    Else
                        Err.Raise vbObjectError + 510, , "Unexpected quote in CSV."
                    End If
                Case ","
                    lineValues.Add cellText
                    cellText = ""
                    afterQuote = False
                Case vbCr, vbLf
                    If ch = vbCr And pos < Len(data) Then
                        If Mid$(data, pos + 1, 1) = vbLf Then pos = pos + 1
                    End If
                    endedRow = True
                Case Else
                    If afterQuote Then _
                        Err.Raise vbObjectError + 511, , "Characters after closing quote."
                    cellText = cellText & ch
            End Select
        End If
        If endedRow Then
            lineValues.Add cellText
            AppendCsvRecord output, lineValues
            Set lineValues = New Collection
            cellText = ""
            afterQuote = False
        End If
    Next pos

    If inQuotes Then Err.Raise vbObjectError + 512, , "Unclosed quoted field."
    If Len(cellText) > 0 Or lineValues.Count > 0 Or afterQuote Then
        lineValues.Add cellText
        AppendCsvRecord output, lineValues
    End If
    If output.Count < 2 Then _
       Err.Raise vbObjectError + 513, , "CSV needs headers and at least one data row."
    Set ReadUtf8Csv = output
    Exit Function
ReadFailed:
    On Error Resume Next
    stream.Close
    On Error GoTo 0
    Err.Raise vbObjectError + 514, , "CSV read/parse error in " & filePath & _
              ". Check UTF-8 encoding and CSV quoting."
End Function

Private Sub AppendCsvRecord(ByVal result As Collection, ByVal cells As Collection)
    Dim values() As String, i As Long, allEmpty As Boolean
    If cells.Count = 1 And Len(CStr(cells(1))) = 0 Then Exit Sub
    ReDim values(1 To cells.Count)
    allEmpty = True
    For i = 1 To cells.Count
        values(i) = CStr(cells(i))
        If Len(values(i)) > 0 Then allEmpty = False
    Next i
    If allEmpty Then Exit Sub
    If result.Count > 0 Then
        Dim header As Variant
        header = result(1)
        If UBound(values) <> UBound(header) Then _
           Err.Raise vbObjectError + 515, , "CSV row has wrong field count."
    End If
    result.Add values
End Sub

Private Function CreateHeaderMap(ByVal header As Variant) As Object
    Dim m As Object, c As Long, n As String
    Set m = CreateObject("Scripting.Dictionary")
    m.CompareMode = vbTextCompare
    For c = LBound(header) To UBound(header)
        n = Trim$(CStr(header(c)))
        If Len(n) = 0 Then Err.Raise vbObjectError + 516, , "Empty header name."
        If m.Exists(n) Then Err.Raise vbObjectError + 517, , "Duplicate header: " & n
        m.Add n, c
    Next c
    Set CreateHeaderMap = m
End Function

Private Sub IndexRows(ByVal rows As Collection, ByRef cols() As Long, _
                      ByVal idx As Object, ByVal duplicates As Object)
    Dim r As Long, i As Long, row As Variant, key As String
    Dim val As String, isEmpty As Boolean
    For r = 2 To rows.Count
        row = rows(r)
        key = ""
        isEmpty = False
        For i = LBound(cols) To UBound(cols)
            val = Trim$(CStr(row(cols(i))))
            If Len(val) = 0 Then isEmpty = True
            key = key & Len(val) & ":" & val
        Next i
        If isEmpty Then
            If Not duplicates.Exists("<EMPTY KEY>") Then duplicates.Add "<EMPTY KEY>", True
        ElseIf idx.Exists(key) Then
            If Not duplicates.Exists(key) Then duplicates.Add key, True
        Else
            idx.Add key, r
        End If
    Next r
End Sub

Private Function IsKeyField(ByVal fieldName As String, ByVal keys As Variant) As Boolean
    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        If StrComp(fieldName, Trim$(CStr(keys(i))), vbTextCompare) = 0 Then
            IsKeyField = True
            Exit Function
        End If
    Next i
End Function

Private Function NormalizeValue(ByVal value As String, ByVal mode As String) As String
    Select Case mode
        Case "TRIM": NormalizeValue = Trim$(value)
        Case "IGNORE_CASE": NormalizeValue = UCase$(Trim$(value))
        Case Else: NormalizeValue = value
    End Select
End Function

Private Sub AddDiff(ByVal ws As Worksheet, ByRef nextRow As Long, _
                    ByVal k As String, ByVal status As String, _
                    ByVal columnName As String, ByVal oldV As String, ByVal newV As String)
    ws.Cells(nextRow, 1).Resize(1, 5).NumberFormat = "@"
    ws.Cells(nextRow, 1).Value2 = k
    ws.Cells(nextRow, 2).Value2 = status
    ws.Cells(nextRow, 3).Value2 = columnName
    ws.Cells(nextRow, 4).Value2 = oldV
    ws.Cells(nextRow, 5).Value2 = newV
    nextRow = nextRow + 1
End Sub

Private Sub AddIssue(ByVal ws As Worksheet, ByRef nextRow As Long, _
                     ByVal sourceName As String, ByVal issue As String, ByVal k As String)
    ws.Cells(nextRow, 1).Resize(1, 3).NumberFormat = "@"
    ws.Cells(nextRow, 1).Value2 = sourceName
    ws.Cells(nextRow, 2).Value2 = issue
    ws.Cells(nextRow, 3).Value2 = k
    nextRow = nextRow + 1
End Sub

Private Sub FormatSheet(ByVal ws As Worksheet)
    ws.Rows(1).Font.Bold = True
    ws.Rows(1).Interior.Color = RGB(222, 235, 247)
    ws.Columns.AutoFit
    ws.Rows(1).AutoFilter
End Sub
