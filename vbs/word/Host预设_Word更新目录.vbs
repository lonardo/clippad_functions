' 函数名: HostWordUpdateTableOfContents
' 描述: 更新当前 Word 文档中的目录及页码；适合正文或分页调整完成后的交付前检查
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

    Main = HostWordUpdateTableOfContents(appObj)
End Function

Function HostWordUpdateTableOfContents(appObj)
    On Error Resume Next

    Dim doc, toc, updated, planId
    Set doc = appObj.ActiveDocument
    If Err.Number <> 0 Or TypeName(doc) = "Empty" Or TypeName(doc) = "Nothing" Then
        Err.Clear
        HostWordUpdateTableOfContents = "{""ok"":false,""code"":""E_NO_DOCUMENT"",""message"":""当前没有活动 Word 文档""}"
        Exit Function
    End If
    If doc.TablesOfContents.Count = 0 Then
        HostWordUpdateTableOfContents = "{""ok"":false,""code"":""E_NO_TOC"",""message"":""当前文档没有目录，请先在 Word 中插入目录""}"
        Exit Function
    End If

    planId = SafeBeginWritePlan()
    SafeRecordWrite "word_update_toc", "{""tableOfContentsCount"":" & CStr(doc.TablesOfContents.Count) & "}", "office.word.toc"
    updated = 0
    For Each toc In doc.TablesOfContents
        toc.Update
        If Err.Number = 0 Then updated = updated + 1
        Err.Clear
    Next
    SafeCloseWritePlan planId

    Dim summary
    summary = "Word 目录更新完成；更新目录数=" & CStr(updated)
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostWordUpdateTableOfContents = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""updatedCount"":" & CStr(updated) & "}"
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
