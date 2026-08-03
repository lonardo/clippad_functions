' 函数名: HostWordCenterAutofitTables
' 描述: 将当前 Word 文档全部表格居中对齐、按窗口自适应列宽，并统一表内字体；适合报告交付前表格快速整理
' 适用应用: Word
' 搜索范围: 全文
' 搜索对象: 无

Option Explicit

Const wdAlignParagraphCenter = 1
Const wdAutoFitWindow = 2
Const wdAlignRowCenter = 1

Function Main()
    On Error Resume Next

    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = "{""ok"":false,""code"":""E_NO_WORD_APP"",""message"":""未取得 Word 应用，请在 Word 中运行该预设""}"
        Exit Function
    End If

    Main = HostWordCenterAutofitTables(appObj)
End Function

Function HostWordCenterAutofitTables(appObj)
    On Error Resume Next

    Dim doc
    Set doc = appObj.ActiveDocument
    If Err.Number <> 0 Or TypeName(doc) = "Empty" Or TypeName(doc) = "Nothing" Then
        Err.Clear
        HostWordCenterAutofitTables = "{""ok"":false,""code"":""E_NO_DOCUMENT"",""message"":""当前没有活动 Word 文档""}"
        Exit Function
    End If

    Dim tableSummary, planId, previewJson, changed, tbl, row
    tableSummary = SafeHostText("GetWordTableSummary")
    planId = SafeBeginWritePlan()
    SafeRecordWrite "word_center_autofit_tables", "{""scope"":""activeDocument""}", "office.word.table.format"
    previewJson = SafePreviewWritePlan()

    changed = 0
    For Each tbl In doc.Tables
        tbl.AutoFitBehavior wdAutoFitWindow
        tbl.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        tbl.Range.Font.NameFarEast = "宋体"
        tbl.Range.Font.Name = "Times New Roman"
        tbl.Range.Font.Size = 10.5
        For Each row In tbl.Rows
            row.Alignment = wdAlignRowCenter
            Err.Clear
        Next
        changed = changed + 1
        Err.Clear
    Next
    SafeCloseWritePlan planId

    Dim summary
    summary = "Word 表格居中自适应完成；处理表格=" & CStr(changed) & vbCrLf & _
        "Host.GetWordTableSummary: " & tableSummary & vbCrLf & _
        "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostWordCenterAutofitTables = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""tableCount"":" & CStr(changed) & _
        ",""planId"":""" & EscapeJson(planId) & """}"
End Function

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

Function SafePrompt(message, defaultValue)
    On Error Resume Next
    SafePrompt = Host.Prompt(message, defaultValue)
    If Err.Number <> 0 Then
        SafePrompt = defaultValue
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

Function SafeValidateWritePlan(planId)
    On Error Resume Next
    SafeValidateWritePlan = Host.ValidateWritePlan(planId, False)
    If Err.Number <> 0 Then SafeValidateWritePlan = ""
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
