' 函数名: HostWordRemoveTrailingBlankParagraphs
' 描述: 删除文档末尾多余空段落，缓解“尾页空白页”问题；对标 Word 排版插件常见去空白页能力
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
    Main = HostWordRemoveTrailingBlankParagraphs(appObj)
End Function

Function HostWordRemoveTrailingBlankParagraphs(appObj)
    On Error Resume Next
    Dim doc
    Set doc = appObj.ActiveDocument
    If Err.Number <> 0 Or TypeName(doc) = "Empty" Or TypeName(doc) = "Nothing" Then
        Err.Clear
        HostWordRemoveTrailingBlankParagraphs = "{""ok"":false,""code"":""E_NO_DOCUMENT"",""message"":""当前没有活动 Word 文档""}"
        Exit Function
    End If

    Dim wordSummary, planId, previewJson, removed, index, para, textValue, keepGoing
    wordSummary = SafeHostText("GetWordContextInfo")
    planId = SafeBeginWritePlan()
    SafeRecordWrite "word_remove_trailing_blank_paragraphs", "{""scope"":""activeDocument""}", "office.word.cleanup"
    previewJson = SafePreviewWritePlan()

    removed = 0
    keepGoing = True
    Do While keepGoing And doc.Paragraphs.Count > 1
        index = doc.Paragraphs.Count
        Set para = doc.Paragraphs(index)
        textValue = CStr(para.Range.Text)
        textValue = Replace(textValue, vbCr, "")
        textValue = Replace(textValue, vbLf, "")
        textValue = Replace(textValue, vbTab, "")
        textValue = Replace(textValue, ChrW(12), "")
        textValue = Replace(textValue, ChrW(160), "")
        textValue = Replace(textValue, ChrW(12288), "")
        textValue = Trim(textValue)
        If Len(textValue) = 0 Then
            para.Range.Delete
            If Err.Number = 0 Then
                removed = removed + 1
            Else
                keepGoing = False
            End If
            Err.Clear
        Else
            keepGoing = False
        End If
    Loop
    SafeCloseWritePlan planId

    Dim summary
    summary = "Word 尾部空段落清理完成；删除=" & CStr(removed) & vbCrLf & _
        "Host.GetWordContextInfo: " & wordSummary & vbCrLf & "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostWordRemoveTrailingBlankParagraphs = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""removed"":" & CStr(removed) & _
        ",""planId"":""" & EscapeJson(planId) & """}"
End Function

Function SafeHostText(methodName)
    On Error Resume Next
    If methodName = "GetWordContextInfo" Then
        SafeHostText = Host.GetWordContextInfo()
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