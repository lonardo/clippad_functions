' 函数名: HostWordFormatTablesThreeLine
' 描述: 将当前 Word 文档中的表格整理为常见三线表风格，并复制表格摘要；验证 Word 表格能力和 Host 写入计划诊断
' 适用应用: Word
' 搜索范围: 全文
' 搜索对象: 无

Option Explicit

Const wdLineStyleNone = 0
Const wdLineStyleSingle = 1
Const wdBorderTop = -1
Const wdBorderLeft = -2
Const wdBorderBottom = -3
Const wdBorderRight = -4
Const wdBorderHorizontal = -5
Const wdBorderVertical = -6
Const wdLineWidth050pt = 4
Const wdLineWidth150pt = 12
Const wdAutoFitWindow = 2

Function Main()
    On Error Resume Next

    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = "{""ok"":false,""code"":""E_NO_WORD_APP"",""message"":""未取得 Word 应用，请在 Word 中运行该预设""}"
        Exit Function
    End If

    Main = HostWordFormatTablesThreeLine(appObj)
End Function

Function HostWordFormatTablesThreeLine(appObj)
    On Error Resume Next

    Dim doc
    Set doc = appObj.ActiveDocument
    If Err.Number <> 0 Or TypeName(doc) = "Empty" Or TypeName(doc) = "Nothing" Then
        Err.Clear
        HostWordFormatTablesThreeLine = "{""ok"":false,""code"":""E_NO_DOCUMENT"",""message"":""当前没有活动 Word 文档""}"
        Exit Function
    End If

    If doc.Tables.Count = 0 Then
        HostWordFormatTablesThreeLine = "{""ok"":false,""code"":""E_NO_TABLE"",""message"":""当前文档没有表格""}"
        Exit Function
    End If

    Dim planId, previewJson
    planId = SafeBeginWritePlan()
    SafeRecordWrite "word_tables_three_line", "{""tableCount"":" & CStr(doc.Tables.Count) & "}", "office.word.table.format"
    previewJson = SafePreviewWritePlan()

    Dim tbl, count
    count = 0
    For Each tbl In doc.Tables
        FormatOneTable tbl
        count = count + 1
    Next
    SafeCloseWritePlan planId

    Dim wordTableSummary, summary
    wordTableSummary = SafeHostText("GetWordTableSummary")
    summary = "Word 三线表整理完成；表格数=" & CStr(count) & vbCrLf & "Host.GetWordTableSummary: " & wordTableSummary
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostWordFormatTablesThreeLine = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""tableCount"":" & CStr(count) & _
        ",""planId"":""" & EscapeJson(planId) & """,""preview"":""" & EscapeJson(previewJson) & """}"

End Function

Sub FormatOneTable(tbl)
    On Error Resume Next
    tbl.AutoFitBehavior wdAutoFitWindow
    tbl.Range.Font.NameFarEast = "宋体"
    tbl.Range.Font.Name = "Times New Roman"
    tbl.Range.Font.Size = 10.5
    tbl.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
    tbl.Borders(wdBorderRight).LineStyle = wdLineStyleNone
    tbl.Borders(wdBorderVertical).LineStyle = wdLineStyleNone
    tbl.Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
    tbl.Borders(wdBorderTop).LineStyle = wdLineStyleSingle
    tbl.Borders(wdBorderTop).LineWidth = wdLineWidth150pt
    tbl.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
    tbl.Borders(wdBorderBottom).LineWidth = wdLineWidth150pt
    If tbl.Rows.Count >= 1 Then
        tbl.Rows(1).Range.Font.Bold = True
        tbl.Rows(1).Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
        tbl.Rows(1).Borders(wdBorderBottom).LineWidth = wdLineWidth050pt
    End If
End Sub

Function SafeHostText(methodName)
    On Error Resume Next
    If methodName = "GetWordTableSummary" Then
        SafeHostText = Host.GetWordTableSummary()
    Else
        SafeHostText = ""
    End If
    If Err.Number <> 0 Then
        SafeHostText = ""
        Err.Clear
    End If
End Function

Function SafeBeginWritePlan()
    On Error Resume Next
    SafeBeginWritePlan = Host.BeginWritePlan("")
    If Err.Number <> 0 Then SafeBeginWritePlan = ""
    Err.Clear
End Function

Sub SafeRecordWrite(actionId, paramsJson, capabilityId)
    On Error Resume Next
    Host.RecordWrite actionId, paramsJson, "office", "", "", capabilityId
    Err.Clear
End Sub

Function SafePreviewWritePlan()
    On Error Resume Next
    SafePreviewWritePlan = Host.PreviewWritePlan()
    If Err.Number <> 0 Then SafePreviewWritePlan = ""
    Err.Clear
End Function

Sub SafeCloseWritePlan(planId)
    On Error Resume Next
    Host.RollbackWritePlan planId
    Err.Clear
End Sub

Sub SafeWriteLog(message)
    On Error Resume Next
    Host.WriteLog message
    Err.Clear
End Sub

Function EscapeJson(value)
    Dim text
    text = CStr(value)
    text = Replace(text, "\", "\\")
    text = Replace(text, """", "\""")
    text = Replace(text, vbCrLf, "\n")
    text = Replace(text, vbCr, "\n")
    text = Replace(text, vbLf, "\n")
    text = Replace(text, vbTab, "\t")
    EscapeJson = text
End Function
