' 函数名: HostWordRemoveHyperlinks
' 描述: 删除当前 Word 文档中的超链接并保留显示文字；适合外部资料整理和文档交付前清理
' 适用应用: Word
' 搜索范围: 全文
' 搜索对象: 文本

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

    Main = HostWordRemoveHyperlinks(appObj)
End Function

Function HostWordRemoveHyperlinks(appObj)
    On Error Resume Next

    Dim doc, confirmation, beforeCount, index, removed, planId
    Set doc = appObj.ActiveDocument
    If Err.Number <> 0 Or TypeName(doc) = "Empty" Or TypeName(doc) = "Nothing" Then
        Err.Clear
        HostWordRemoveHyperlinks = "{""ok"":false,""code"":""E_NO_DOCUMENT"",""message"":""当前没有活动 Word 文档""}"
        Exit Function
    End If

    beforeCount = doc.Hyperlinks.Count
    If beforeCount = 0 Then
        HostWordRemoveHyperlinks = "{""ok"":true,""message"":""当前文档没有超链接"",""removedCount"":0}"
        Exit Function
    End If
    confirmation = SafePrompt("将删除 " & CStr(beforeCount) & " 个超链接并保留文字。请输入 删除 继续", "")
    If confirmation <> "删除" Then
        HostWordRemoveHyperlinks = "{""ok"":false,""code"":""E_CONFIRM_REQUIRED"",""message"":""已取消删除超链接""}"
        Exit Function
    End If

    planId = SafeBeginWritePlan()
    SafeRecordWrite "word_remove_hyperlinks", "{""count"":" & CStr(beforeCount) & "}", "office.word.link.cleanup"
    removed = 0
    For index = doc.Hyperlinks.Count To 1 Step -1
        doc.Hyperlinks(index).Delete
        If Err.Number = 0 Then removed = removed + 1
        Err.Clear
    Next
    SafeCloseWritePlan planId

    Dim summary
    summary = "Word 超链接清理完成；删除=" & CStr(removed) & "；保留显示文字"
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostWordRemoveHyperlinks = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""removedCount"":" & CStr(removed) & "}"
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
    EscapeJson = text
End Function
