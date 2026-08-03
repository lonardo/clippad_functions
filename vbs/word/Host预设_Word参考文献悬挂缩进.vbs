' 函数名: HostWordFormatReferences
' 描述: 从“参考文献”标题之后批量套用宋体小五、固定段距和悬挂缩进；适合论文、报告和申报材料排版
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

    Main = HostWordFormatReferences(appObj)
End Function

Function HostWordFormatReferences(appObj)
    On Error Resume Next

    Dim doc
    Set doc = appObj.ActiveDocument
    If Err.Number <> 0 Or TypeName(doc) = "Empty" Or TypeName(doc) = "Nothing" Then
        Err.Clear
        HostWordFormatReferences = "{""ok"":false,""code"":""E_NO_DOCUMENT"",""message"":""当前没有活动 Word 文档""}"
        Exit Function
    End If

    Dim reviewSummary, planId, previewJson
    reviewSummary = SafeHostText("GetWordSummary")
    planId = SafeBeginWritePlan()
    SafeRecordWrite "word_format_references", "{""scope"":""afterReferencesHeading""}", "office.word.referenceFormat"
    previewJson = SafePreviewWritePlan()

    Dim para, textValue, inReferences, formattedCount
    inReferences = False
    formattedCount = 0
    For Each para In doc.Paragraphs
        textValue = CleanParagraphText(para.Range.Text)
        If IsReferencesHeading(textValue) Then
            inReferences = True
            FormatReferencesHeading para
        ElseIf inReferences And Len(textValue) > 0 Then
            FormatReferenceParagraph appObj, para
            formattedCount = formattedCount + 1
        End If
        If Err.Number <> 0 Then Err.Clear
    Next
    SafeCloseWritePlan planId

    If Not inReferences Then
        HostWordFormatReferences = "{""ok"":false,""code"":""E_NO_REFERENCES_HEADING"",""message"":""未找到参考文献标题""}"
        Exit Function
    End If

    Dim summary
    summary = "Word 参考文献悬挂缩进完成；条目=" & CStr(formattedCount) & vbCrLf & _
        "Host.GetWordSummary: " & reviewSummary & vbCrLf & _
        "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostWordFormatReferences = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""formattedCount"":" & CStr(formattedCount) & "}"

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

Function IsReferencesHeading(textValue)
    Dim normalized
    normalized = Replace(CStr(textValue), " ", "")
    IsReferencesHeading = (normalized = "参考文献" Or normalized = "参考资料")
End Function

Sub FormatReferencesHeading(para)
    On Error Resume Next
    para.Range.Font.NameFarEast = "黑体"
    para.Range.Font.Name = "Arial"
    para.Range.Font.Size = 14
    para.Range.Font.Bold = True
    para.Format.FirstLineIndent = 0
    para.Format.SpaceBefore = 12
    para.Format.SpaceAfter = 6
    Err.Clear
End Sub

Sub FormatReferenceParagraph(appObj, para)
    On Error Resume Next
    Dim hang
    hang = appObj.CentimetersToPoints(0.74)
    para.Range.Font.NameFarEast = "宋体"
    para.Range.Font.Name = "Times New Roman"
    para.Range.Font.Size = 9
    para.Range.Font.Bold = False
    para.Format.LeftIndent = hang
    para.Format.FirstLineIndent = -hang
    para.Format.SpaceBefore = 0
    para.Format.SpaceAfter = 0
    Err.Clear
End Sub

Function SafeHostText(methodName)
    On Error Resume Next
    If methodName = "GetWordSummary" Then
        SafeHostText = Host.GetWordSummary()
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
