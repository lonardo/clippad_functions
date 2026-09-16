' 函数名: HostWordSetupPageNumberFooter
' 描述: 为当前 Word 文档各节页脚设置居中页码，并可选写入页脚说明文字；适合报告、方案和交付文档快速补齐页码
' 适用应用: Word
' 搜索范围: 全文
' 搜索对象: 无

Option Explicit

Const wdAlignParagraphCenter = 1
Const wdHeaderFooterPrimary = 1
Const wdFieldPage = 33
Const wdCollapseEnd = 0

Function Main()
    On Error Resume Next

    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = "{""ok"":false,""code"":""E_NO_WORD_APP"",""message"":""未取得 Word 应用，请在 Word 中运行该预设""}"
        Exit Function
    End If

    Main = HostWordSetupPageNumberFooter(appObj)
End Function

Function HostWordSetupPageNumberFooter(appObj)
    On Error Resume Next

    Dim doc
    Set doc = appObj.ActiveDocument
    If Err.Number <> 0 Or TypeName(doc) = "Empty" Or TypeName(doc) = "Nothing" Then
        Err.Clear
        HostWordSetupPageNumberFooter = "{""ok"":false,""code"":""E_NO_DOCUMENT"",""message"":""当前没有活动 Word 文档""}"
        Exit Function
    End If

    Dim footerText, wordSummary, planId, previewJson, section, footer, changed
    Dim confirmation
    confirmation = SafePrompt("将重写各节页脚为居中页码（会覆盖原页脚）。请输入 设置 继续", "")
    If confirmation <> "设置" Then
        HostWordSetupPageNumberFooter = "{""ok"":false,""code"":""E_CONFIRM_REQUIRED"",""message"":""已取消页码页脚设置""}"
        Exit Function
    End If
    footerText = SafePrompt("可选：输入页脚说明文字（可留空，仅设置页码）", "")
    wordSummary = SafeHostText("GetWordContextInfo")
    planId = SafeBeginWritePlan()
    SafeRecordWrite "word_setup_page_number_footer", "{""hasFooterText"":" & LCase(CStr(Len(footerText) > 0)) & "}", "office.word.headerfooter"
    previewJson = SafePreviewWritePlan()

    changed = 0
    For Each section In doc.Sections
        Set footer = section.Footers(wdHeaderFooterPrimary)
        footer.Range.Delete
        footer.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        footer.Range.Font.NameFarEast = "宋体"
        footer.Range.Font.Name = "Times New Roman"
        footer.Range.Font.Size = 10.5
        If Len(footerText) > 0 Then
            footer.Range.Text = footerText & "  "
            footer.Range.Collapse wdCollapseEnd
        End If
        footer.Range.Fields.Add footer.Range, wdFieldPage
        changed = changed + 1
        Err.Clear
    Next
    SafeCloseWritePlan planId

    Dim summary
    summary = "Word 页码页脚设置完成；处理节数=" & CStr(changed) & "；页脚说明长度=" & CStr(Len(footerText)) & vbCrLf & _
        "Host.GetWordContextInfo: " & wordSummary & vbCrLf & _
        "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostWordSetupPageNumberFooter = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""sectionCount"":" & CStr(changed) & _
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
