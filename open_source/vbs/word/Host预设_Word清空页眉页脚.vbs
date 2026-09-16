' 函数名: HostWordClearHeaderFooter
' 描述: 清空当前 Word 文档全部节的页眉页脚内容；对标稻壳/公文工具中的页眉页脚一键清理
' 适用应用: Word
' 搜索范围: 全文
' 搜索对象: 无

Option Explicit

Const wdHeaderFooterPrimary = 1
Const wdHeaderFooterFirstPage = 2
Const wdHeaderFooterEvenPages = 3

Function Main()
    On Error Resume Next
    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = "{""ok"":false,""code"":""E_NO_WORD_APP"",""message"":""未取得 Word 应用，请在 Word 中运行该预设""}"
        Exit Function
    End If
    Main = HostWordClearHeaderFooter(appObj)
End Function

Function HostWordClearHeaderFooter(appObj)
    On Error Resume Next
    Dim doc
    Set doc = appObj.ActiveDocument
    If Err.Number <> 0 Or TypeName(doc) = "Empty" Or TypeName(doc) = "Nothing" Then
        Err.Clear
        HostWordClearHeaderFooter = "{""ok"":false,""code"":""E_NO_DOCUMENT"",""message"":""当前没有活动 Word 文档""}"
        Exit Function
    End If

    Dim confirmation
    confirmation = SafePrompt("将清空全部页眉页脚。请输入 清空 继续", "")
    If confirmation <> "清空" Then
        HostWordClearHeaderFooter = "{""ok"":false,""code"":""E_CONFIRM_REQUIRED"",""message"":""已取消清空页眉页脚""}"
        Exit Function
    End If

    Dim wordSummary, planId, previewJson, section, cleared, scanned
    wordSummary = SafeHostText("GetWordContextInfo")
    planId = SafeBeginWritePlan()
    SafeRecordWrite "word_clear_header_footer", "{""scope"":""activeDocument""}", "office.word.headerfooter.cleanup"
    previewJson = SafePreviewWritePlan()

    cleared = 0
    scanned = 0
    For Each section In doc.Sections
        scanned = scanned + 1
        cleared = cleared + ClearHeaderFooterPart(section.Headers(wdHeaderFooterPrimary))
        cleared = cleared + ClearHeaderFooterPart(section.Headers(wdHeaderFooterFirstPage))
        cleared = cleared + ClearHeaderFooterPart(section.Headers(wdHeaderFooterEvenPages))
        cleared = cleared + ClearHeaderFooterPart(section.Footers(wdHeaderFooterPrimary))
        cleared = cleared + ClearHeaderFooterPart(section.Footers(wdHeaderFooterFirstPage))
        cleared = cleared + ClearHeaderFooterPart(section.Footers(wdHeaderFooterEvenPages))
    Next
    SafeCloseWritePlan planId

    Dim summary
    summary = "Word 页眉页脚清空完成；处理节=" & CStr(scanned) & "；清理区域=" & CStr(cleared) & vbCrLf & _
        "Host.GetWordContextInfo: " & wordSummary & vbCrLf & "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostWordClearHeaderFooter = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""sections"":" & CStr(scanned) & _
        ",""clearedParts"":" & CStr(cleared) & ",""planId"":""" & EscapeJson(planId) & """}"
End Function

Function ClearHeaderFooterPart(part)
    On Error Resume Next
    ClearHeaderFooterPart = 0
    If TypeName(part) = "Nothing" Or TypeName(part) = "Empty" Then Exit Function
    part.Range.Delete
    If Err.Number = 0 Then ClearHeaderFooterPart = 1
    Err.Clear
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