' 函数名: HostWordNormalizeHeadingLevels
' 描述: 识别 Word 文档中常见中文标题编号并套用一二三级标题排版；适合报告、制度和公文初稿快速规范层级
' 适用应用: Word
' 搜索范围: 全文
' 搜索对象: 无

Option Explicit

Const wdAlignParagraphLeft = 0
Const wdAlignParagraphCenter = 1
Const wdOutlineLevelBodyText = 10

Function Main()
    On Error Resume Next

    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = "{""ok"":false,""code"":""E_NO_WORD_APP"",""message"":""未取得 Word 应用，请在 Word 中运行该预设""}"
        Exit Function
    End If

    Main = HostWordNormalizeHeadingLevels(appObj)
End Function

Function HostWordNormalizeHeadingLevels(appObj)
    On Error Resume Next

    Dim doc
    Set doc = appObj.ActiveDocument
    If Err.Number <> 0 Or TypeName(doc) = "Empty" Or TypeName(doc) = "Nothing" Then
        Err.Clear
        HostWordNormalizeHeadingLevels = "{""ok"":false,""code"":""E_NO_DOCUMENT"",""message"":""当前没有活动 Word 文档""}"
        Exit Function
    End If

    Dim wordContext, formatIssueSummary, planId, previewJson
    wordContext = SafeHostText("GetWordContextInfo")
    formatIssueSummary = SafeHostText("GetWordFormatIssueSummary")
    planId = SafeBeginWritePlan()
    SafeRecordWrite "word_normalize_heading_levels", "{""scope"":""activeDocument""}", "office.word.headingFormat"
    previewJson = SafePreviewWritePlan()

    Dim para, textValue, level, level1Count, level2Count, level3Count
    level1Count = 0
    level2Count = 0
    level3Count = 0
    For Each para In doc.Paragraphs
        textValue = CleanParagraphText(para.Range.Text)
        level = DetectHeadingLevel(textValue)
        If level > 0 Then
            ApplyHeadingFormat appObj, para, level
            If level = 1 Then level1Count = level1Count + 1
            If level = 2 Then level2Count = level2Count + 1
            If level = 3 Then level3Count = level3Count + 1
        End If
        If Err.Number <> 0 Then Err.Clear
    Next
    SafeCloseWritePlan planId

    Dim summary
    summary = "Word 标题层级样式规范完成；一级=" & CStr(level1Count) & _
        "；二级=" & CStr(level2Count) & "；三级=" & CStr(level3Count) & vbCrLf & _
        "Host.GetWordContextInfo: " & wordContext & vbCrLf & _
        "Host.GetWordFormatIssueSummary: " & formatIssueSummary & vbCrLf & _
        "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostWordNormalizeHeadingLevels = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""level1"":" & CStr(level1Count) & _
        ",""level2"":" & CStr(level2Count) & ",""level3"":" & CStr(level3Count) & "}"

End Function

Function CleanParagraphText(value)
    Dim text
    text = CStr(value)
    text = Replace(text, vbCr, "")
    text = Replace(text, vbLf, "")
    text = Replace(text, ChrW(160), " ")
    text = Replace(text, ChrW(12288), " ")
    CleanParagraphText = Trim(text)
End Function

Function DetectHeadingLevel(textValue)
    DetectHeadingLevel = 0
    If Len(textValue) = 0 Or Len(textValue) > 80 Then Exit Function
    If IsLevelOneHeading(textValue) Then
        DetectHeadingLevel = 1
    ElseIf IsLevelTwoHeading(textValue) Then
        DetectHeadingLevel = 2
    ElseIf IsLevelThreeHeading(textValue) Then
        DetectHeadingLevel = 3
    End If
End Function

Function IsLevelOneHeading(textValue)
    Dim firstChar, secondChar
    firstChar = Left(textValue, 1)
    secondChar = Mid(textValue, 2, 1)
    IsLevelOneHeading = (InStr("一二三四五六七八九十", firstChar) > 0 And (secondChar = "、" Or secondChar = "." Or secondChar = "．"))
End Function

Function IsLevelTwoHeading(textValue)
    IsLevelTwoHeading = ((Left(textValue, 1) = "（" Or Left(textValue, 1) = "(") And _
        InStr("一二三四五六七八九十", Mid(textValue, 2, 1)) > 0 And _
        (Mid(textValue, 3, 1) = "）" Or Mid(textValue, 3, 1) = ")"))
End Function

Function IsLevelThreeHeading(textValue)
    Dim i, ch
    i = 1
    Do While i <= Len(textValue)
        ch = Mid(textValue, i, 1)
        If InStr("0123456789", ch) = 0 Then Exit Do
        i = i + 1
    Loop
    IsLevelThreeHeading = (i > 1 And i <= Len(textValue) And (Mid(textValue, i, 1) = "." Or Mid(textValue, i, 1) = "、" Or Mid(textValue, i, 1) = "．"))
End Function

Sub ApplyHeadingFormat(appObj, para, level)
    On Error Resume Next
    para.Format.FirstLineIndent = 0
    para.Format.SpaceBefore = 6
    para.Format.SpaceAfter = 6
    para.OutlineLevel = level
    If level = 1 Then
        para.Range.Font.NameFarEast = "黑体"
        para.Range.Font.Name = "Arial"
        para.Range.Font.Size = 16
        para.Range.Font.Bold = True
        para.Format.Alignment = wdAlignParagraphCenter
    ElseIf level = 2 Then
        para.Range.Font.NameFarEast = "楷体"
        para.Range.Font.Name = "Times New Roman"
        para.Range.Font.Size = 14
        para.Range.Font.Bold = True
        para.Format.Alignment = wdAlignParagraphLeft
    Else
        para.Range.Font.NameFarEast = "仿宋"
        para.Range.Font.Name = "Times New Roman"
        para.Range.Font.Size = 12
        para.Range.Font.Bold = True
        para.Format.Alignment = wdAlignParagraphLeft
    End If
    Err.Clear
End Sub

Function SafeHostText(methodName)
    On Error Resume Next
    If methodName = "GetWordContextInfo" Then
        SafeHostText = Host.GetWordContextInfo()
    ElseIf methodName = "GetWordFormatIssueSummary" Then
        SafeHostText = Host.GetWordFormatIssueSummary()
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
