' 函数名: HostWordNormalizeBodyParagraphs
' 描述: 将正文段落统一为宋体小四、首行缩进 2 字符、1.5 倍行距；对标公文排版插件的正文一键规范
' 适用应用: Word
' 搜索范围: 全文
' 搜索对象: 无

Option Explicit

Const wdAlignParagraphJustify = 3
Const wdLineSpace1pt5 = 1
Const wdCollapseStart = 1
Const wdCharacter = 1

Function Main()
    On Error Resume Next
    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = "{""ok"":false,""code"":""E_NO_WORD_APP"",""message"":""未取得 Word 应用，请在 Word 中运行该预设""}"
        Exit Function
    End If
    Main = HostWordNormalizeBodyParagraphs(appObj)
End Function

Function HostWordNormalizeBodyParagraphs(appObj)
    On Error Resume Next
    Dim doc
    Set doc = appObj.ActiveDocument
    If Err.Number <> 0 Or TypeName(doc) = "Empty" Or TypeName(doc) = "Nothing" Then
        Err.Clear
        HostWordNormalizeBodyParagraphs = "{""ok"":false,""code"":""E_NO_DOCUMENT"",""message"":""当前没有活动 Word 文档""}"
        Exit Function
    End If

    Dim formatSummary, planId, previewJson, para, changed, scanned, styleName, textValue
    formatSummary = SafeHostText("GetWordFormatIssueSummary")
    planId = SafeBeginWritePlan()
    SafeRecordWrite "word_normalize_body_paragraphs", "{""scope"":""activeDocument""}", "office.word.format.body"
    previewJson = SafePreviewWritePlan()

    changed = 0
    scanned = 0
    For Each para In doc.Paragraphs
        scanned = scanned + 1
        styleName = ""
        styleName = CStr(para.Style)
        textValue = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
        If Len(textValue) = 0 Then
            ' skip empty
        ElseIf InStr(1, styleName, "标题", vbTextCompare) > 0 Or InStr(1, styleName, "Heading", vbTextCompare) > 0 Then
            ' skip headings
        Else
            para.Range.Font.NameFarEast = "宋体"
            para.Range.Font.Name = "Times New Roman"
            para.Range.Font.Size = 12
            para.Format.Alignment = wdAlignParagraphJustify
            para.Format.LineSpacingRule = wdLineSpace1pt5
            para.Format.CharacterUnitFirstLineIndent = 2
            para.Format.SpaceBefore = 0
            para.Format.SpaceAfter = 0
            changed = changed + 1
        End If
        Err.Clear
    Next
    SafeCloseWritePlan planId

    Dim summary
    summary = "Word 正文段落规范完成；扫描=" & CStr(scanned) & "；处理=" & CStr(changed) & vbCrLf & _
        "Host.GetWordFormatIssueSummary: " & formatSummary & vbCrLf & "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostWordNormalizeBodyParagraphs = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""scanned"":" & CStr(scanned) & _
        ",""changed"":" & CStr(changed) & ",""planId"":""" & EscapeJson(planId) & """}"
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