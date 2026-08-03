' 函数名: HostWordRemoveEmptyParagraphs
' 描述: 删除当前 Word 文档中的空段落和仅含空白字符的段落，并复制清理摘要；适合粘贴网页、PDF 转换稿和多人协作草稿
' 适用应用: Word
' 搜索范围: 全文
' 搜索对象: 无

Option Explicit

Function Main()
    On Error Resume Next

    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = "{""ok"":false,""code"":""E_NO_WORD_APP"",""message"":""未取得 Word 应用，请在 Word 中运行该预设""}"
        Exit Function
    End If

    Main = HostWordRemoveEmptyParagraphs(appObj)
End Function

Function HostWordRemoveEmptyParagraphs(appObj)
    On Error Resume Next

    Dim doc
    Set doc = appObj.ActiveDocument
    If Err.Number <> 0 Or TypeName(doc) = "Empty" Or TypeName(doc) = "Nothing" Then
        Err.Clear
        HostWordRemoveEmptyParagraphs = "{""ok"":false,""code"":""E_NO_DOCUMENT"",""message"":""当前没有活动 Word 文档""}"
        Exit Function
    End If

    Dim formatSummary, planId, previewJson, removed, scanned, index, para, textValue
    formatSummary = SafeHostText("GetWordFormatIssueSummary")
    planId = SafeBeginWritePlan()
    SafeRecordWrite "word_remove_empty_paragraphs", "{""scope"":""activeDocument""}", "office.word.cleanup"
    previewJson = SafePreviewWritePlan()

    removed = 0
    scanned = doc.Paragraphs.Count
    For index = doc.Paragraphs.Count To 1 Step -1
        Set para = doc.Paragraphs(index)
        textValue = NormalizeParagraphText(para.Range.Text)
        If Len(textValue) = 0 Then
            para.Range.Delete
            If Err.Number = 0 Then removed = removed + 1
            Err.Clear
        End If
    Next
    SafeCloseWritePlan planId

    Dim summary
    summary = "Word 空段落清理完成；扫描段落=" & CStr(scanned) & "；删除=" & CStr(removed) & vbCrLf & _
        "Host.GetWordFormatIssueSummary: " & formatSummary & vbCrLf & _
        "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostWordRemoveEmptyParagraphs = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""scanned"":" & CStr(scanned) & _
        ",""removed"":" & CStr(removed) & ",""planId"":""" & EscapeJson(planId) & """}"
End Function

Function NormalizeParagraphText(value)
    Dim text
    text = CStr(value)
    text = Replace(text, vbCr, "")
    text = Replace(text, vbLf, "")
    text = Replace(text, vbTab, "")
    text = Replace(text, ChrW(7), "")
    text = Replace(text, ChrW(160), " ")
    text = Replace(text, ChrW(12288), " ")
    NormalizeParagraphText = Trim(text)
End Function

Function SafeHostText(methodName)
    On Error Resume Next
    If methodName = "GetWordFormatIssueSummary" Then
        SafeHostText = Host.GetWordFormatIssueSummary()
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
