' 函数名: HostWordDocumentFormat
' 描述: 对当前 Word 文档执行常见公文基础排版：页面边距、正文字体段距、标题居中、表格自适应，并记录 Host 写入计划
' 适用应用: Word
' 搜索范围: 全文
' 搜索对象: 无

Option Explicit

Const wdAlignParagraphLeft = 0
Const wdAlignParagraphCenter = 1
Const wdLineSpaceSingle = 0
Const wdAutoFitWindow = 2

Function Main()
    On Error Resume Next

    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = "{""ok"":false,""code"":""E_NO_WORD_APP"",""message"":""未取得 Word 应用，请在 Word 中运行该预设""}"
        Exit Function
    End If

    Main = HostWordDocumentFormat(appObj)
End Function

Function HostWordDocumentFormat(appObj)
    On Error Resume Next

    Dim doc
    Set doc = Nothing
    Err.Clear
    Set doc = appObj.ActiveDocument
    If Err.Number <> 0 Or TypeName(doc) = "Empty" Or TypeName(doc) = "Nothing" Then
        Err.Clear
        HostWordDocumentFormat = "{""ok"":false,""code"":""E_NO_DOCUMENT"",""message"":""当前没有活动 Word 文档""}"
        Exit Function
    End If

    Dim planId, previewJson, validateJson
    planId = SafeBeginWritePlan()
    SafeRecordWrite "word_document_format", "{""scope"":""activeDocument"",""preset"":""official_document_basic""}", "office.word.format"
    validateJson = SafeValidateWritePlan(planId)
    previewJson = SafePreviewWritePlan()

    Dim oldScreenUpdating, hadScreenUpdating
    hadScreenUpdating = False
    Err.Clear
    oldScreenUpdating = appObj.ScreenUpdating
    If Err.Number = 0 Then
        hadScreenUpdating = True
        appObj.ScreenUpdating = False
    Else
        Err.Clear
    End If

    Dim changedParagraphs, changedTables, headingCount
    changedParagraphs = 0
    changedTables = 0
    headingCount = 0

    Err.Clear
    doc.PageSetup.TopMargin = appObj.CentimetersToPoints(3.7)
    doc.PageSetup.BottomMargin = appObj.CentimetersToPoints(3.5)
    doc.PageSetup.LeftMargin = appObj.CentimetersToPoints(2.8)
    doc.PageSetup.RightMargin = appObj.CentimetersToPoints(2.6)
    Err.Clear

    Dim para, textValue
    For Each para In doc.Paragraphs
        textValue = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
        If Len(textValue) > 0 Then
            changedParagraphs = changedParagraphs + 1
            para.Range.Font.NameFarEast = "仿宋"
            para.Range.Font.Name = "Times New Roman"
            para.Range.Font.Size = 12
            para.Format.Alignment = wdAlignParagraphLeft
            para.Format.LineSpacingRule = wdLineSpaceSingle
            para.Format.SpaceBefore = 0
            para.Format.SpaceAfter = 0
            para.Format.FirstLineIndent = appObj.CentimetersToPoints(0.74)
            If LooksLikeHeading(textValue) Then
                headingCount = headingCount + 1
                para.Range.Font.NameFarEast = "黑体"
                para.Range.Font.Bold = True
                para.Format.Alignment = wdAlignParagraphCenter
                para.Format.FirstLineIndent = 0
            End If
        End If
    Next

    Dim tbl
    For Each tbl In doc.Tables
        changedTables = changedTables + 1
        tbl.Range.Font.NameFarEast = "仿宋"
        tbl.Range.Font.Size = 10.5
        tbl.AutoFitBehavior wdAutoFitWindow
    Next

    If hadScreenUpdating Then appObj.ScreenUpdating = oldScreenUpdating
    SafeCloseWritePlan planId

    Dim summary
    summary = "Word 公文基础排版完成；段落=" & CStr(changedParagraphs) & "，标题=" & CStr(headingCount) & "，表格=" & CStr(changedTables)
    SafeWriteLog summary
    Host.WriteClipboard summary

    HostWordDocumentFormat = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""planId"":""" & EscapeJson(planId) & _
        """,""preview"":""" & EscapeJson(previewJson) & """,""validation"":""" & EscapeJson(validateJson) & """}"

End Function

Function LooksLikeHeading(textValue)
    LooksLikeHeading = False
    If Len(textValue) <= 40 Then
        If InStr(textValue, "关于") > 0 Or InStr(textValue, "通知") > 0 Or InStr(textValue, "报告") > 0 Or InStr(textValue, "方案") > 0 Then
            LooksLikeHeading = True
        End If
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
