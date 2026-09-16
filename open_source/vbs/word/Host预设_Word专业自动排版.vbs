' 函数名: HostWordProfessionalFormatter
' Description: Rule-based Word auto formatter inspired by Word-Formatter-Pro; preview and apply safely.
' 描述: 参考 Word-Formatter-Pro 的规则化思路，自动识别标题、层级标题、正文、列表和表格并统一页面与段落格式；支持预览后确认应用，业务文字保持不变
' 适用应用: Word
' 搜索范围: 全文
' 搜索对象: 无

Option Explicit

Const wdPaperA4 = 7
Const wdAlignParagraphLeft = 0
Const wdAlignParagraphCenter = 1
Const wdAlignParagraphJustify = 3
Const wdLineSpaceSingle = 0
Const wdLineSpace1pt5 = 1
Const wdHeaderFooterPrimary = 1
Const wdAutoFitWindow = 2
Const wdAlignRowCenter = 1
Const wdOutlineLevelBodyText = 10

Function Main()
    On Error Resume Next
    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or IsMissingObject(appObj) Then
        Err.Clear
        Main = FailureJson("E_NO_WORD_APP", "未取得 Word 应用，请在 Word 中运行此预设 / Run this preset from Word")
        Exit Function
    End If
    Main = HostWordProfessionalFormatter(appObj)
End Function

Function HostWordProfessionalFormatter(appObj)
    On Error Resume Next
    Dim doc
    Set doc = appObj.ActiveDocument
    If Err.Number <> 0 Or IsMissingObject(doc) Then
        Err.Clear
        HostWordProfessionalFormatter = FailureJson("E_NO_DOCUMENT", "当前没有活动 Word 文档 / No active Word document")
        Exit Function
    End If
    If IsDocumentProtected(doc) Then
        HostWordProfessionalFormatter = FailureJson("E_DOCUMENT_PROTECTED", "当前文档受保护，未修改 / Protected document was not changed")
        Exit Function
    End If

    Dim mode, scopeMode, scopeRange
    mode = NormalizeMode(SafePrompt("运行模式 / Mode：预览 preview 或应用 apply", "apply"))
    scopeMode = NormalizeScope(SafePrompt("处理范围 / Scope：全文 document 或选区 selection", "document"))
    Set scopeRange = GetScopeRange(appObj, doc, scopeMode)
    If IsMissingObject(scopeRange) Then
        HostWordProfessionalFormatter = FailureJson("E_SCOPE_UNAVAILABLE", "无法取得处理范围，未修改 / Unable to get target scope")
        Exit Function
    End If

    Dim titleCount, headingCount, bodyCount, listCount, tableCount
    CountTargets scopeRange, titleCount, headingCount, bodyCount, listCount, tableCount

    Dim issueSummary, contextSummary, previewText, planId, planPreview, planValidation
    issueSummary = SafeHostText("GetWordFormatIssueSummary")
    contextSummary = SafeHostText("GetWordContextInfo")
    previewText = BuildPreviewText(scopeMode, titleCount, headingCount, bodyCount, listCount, tableCount, issueSummary)
    planId = SafeBeginWritePlan()
    SafeRecordWrite "word_professional_formatter", BuildPlanParams(scopeMode, titleCount, headingCount, bodyCount, listCount, tableCount), "office.word.format.professional"
    planPreview = SafePreviewWritePlan()
    planValidation = SafeValidateWritePlan(planId)
    SafeRollbackWritePlan planId

    If mode <> "apply" Then
        Host.WriteClipboard previewText
        SafeWriteLog previewText
        HostWordProfessionalFormatter = SuccessJson(False, scopeMode, titleCount, headingCount, bodyCount, listCount, tableCount, planPreview, planValidation, previewText, contextSummary)
        Exit Function
    End If

    Dim confirmation
    confirmation = Trim(SafePrompt(previewText & vbCrLf & vbCrLf & "确认应用格式且保留正文文字？输入 应用 / APPLY", "应用"))
    If StrComp(confirmation, "应用", vbTextCompare) <> 0 And StrComp(confirmation, "apply", vbTextCompare) <> 0 Then
        HostWordProfessionalFormatter = FailureJson("E_CONFIRM_REQUIRED", "未确认应用，文档未修改 / Not confirmed; document was not changed")
        Exit Function
    End If

    Dim originalText, oldScreenUpdating, hadScreenUpdating
    originalText = CStr(doc.Content.Text)
    hadScreenUpdating = False
    Err.Clear
    oldScreenUpdating = appObj.ScreenUpdating
    If Err.Number = 0 Then
        hadScreenUpdating = True
        appObj.ScreenUpdating = False
    Else
        Err.Clear
    End If

    Dim changedTitle, changedHeading, changedBody, changedList, changedTable, changedSection, failureCount
    changedTitle = 0: changedHeading = 0: changedBody = 0: changedList = 0: changedTable = 0: changedSection = 0: failureCount = 0
    FormatDocumentStyles doc, failureCount
    FormatSections appObj, doc, changedSection, failureCount
    FormatParagraphs appObj, scopeRange, changedTitle, changedHeading, changedBody, changedList, failureCount
    FormatTables scopeRange, changedTable, failureCount
    If hadScreenUpdating Then appObj.ScreenUpdating = oldScreenUpdating

    If CStr(doc.Content.Text) <> originalText Or failureCount > 0 Then
        HostWordProfessionalFormatter = FailureWithStateJson("E_FORMAT_INCOMPLETE", "排版未完整完成，或检测到正文文字变化，已停止 / Formatting was incomplete or text changed", (CStr(doc.Content.Text) = originalText), failureCount)
        Exit Function
    End If

    planId = SafeBeginWritePlan()
    SafeRecordWrite "word_professional_formatter_apply", BuildPlanParams(scopeMode, changedTitle, changedHeading, changedBody, changedList, changedTable), "office.word.format.professional"
    planPreview = SafePreviewWritePlan()
    planValidation = SafeValidateWritePlan(planId)
    SafeRollbackWritePlan planId

    Dim summary
    summary = "Word 专业自动排版完成；标题=" & CStr(changedTitle) & "，层级标题=" & CStr(changedHeading) & "，正文=" & CStr(changedBody) & "，列表=" & CStr(changedList) & "，表格=" & CStr(changedTable) & "，节=" & CStr(changedSection) & "；正文文字保持不变" & vbCrLf & _
        "Word professional formatting completed; title=" & CStr(changedTitle) & ", headings=" & CStr(changedHeading) & ", body=" & CStr(changedBody) & ", lists=" & CStr(changedList) & ", tables=" & CStr(changedTable) & ", sections=" & CStr(changedSection) & "; text preserved=true"
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostWordProfessionalFormatter = SuccessJson(True, scopeMode, changedTitle, changedHeading, changedBody, changedList, changedTable, planPreview, planValidation, summary, contextSummary)
End Function

Sub FormatDocumentStyles(doc, ByRef failureCount)
    On Error Resume Next
    SetStyleFontIfExists doc, Array("Normal", "正文"), "宋体", "Times New Roman", 10.5
    SetStyleFontIfExists doc, Array("Heading 1", "标题 1"), "黑体", "Arial", 16
    SetStyleFontIfExists doc, Array("Heading 2", "标题 2"), "黑体", "Arial", 14
    SetStyleFontIfExists doc, Array("Heading 3", "标题 3"), "楷体", "Times New Roman", 12
    If Err.Number <> 0 Then failureCount = failureCount + 1
    Err.Clear
End Sub

Sub SetStyleFontIfExists(doc, names, farEastName, asciiName, fontSize)
    On Error Resume Next
    Dim i, st
    For i = LBound(names) To UBound(names)
        Err.Clear
        Set st = doc.Styles(names(i))
        If Err.Number = 0 And Not IsMissingObject(st) Then
            st.Font.NameFarEast = farEastName: st.Font.NameAscii = asciiName: st.Font.Size = fontSize
            Exit Sub
        End If
    Next
    Err.Clear
End Sub

Sub FormatSections(appObj, doc, ByRef changedSection, ByRef failureCount)
    On Error Resume Next
    Dim sec
    For Each sec In doc.Sections
        Err.Clear
        sec.PageSetup.PaperSize = wdPaperA4
        sec.PageSetup.TopMargin = appObj.CentimetersToPoints(2.54)
        sec.PageSetup.BottomMargin = appObj.CentimetersToPoints(2.54)
        sec.PageSetup.LeftMargin = appObj.CentimetersToPoints(2.80)
        sec.PageSetup.RightMargin = appObj.CentimetersToPoints(2.60)
        sec.PageSetup.HeaderDistance = appObj.CentimetersToPoints(1.50)
        sec.PageSetup.FooterDistance = appObj.CentimetersToPoints(1.75)
        If Err.Number <> 0 Then failureCount = failureCount + 1 Else changedSection = changedSection + 1
        Err.Clear
        FormatHeaderFooter sec.Headers(wdHeaderFooterPrimary).Range
        FormatHeaderFooter sec.Footers(wdHeaderFooterPrimary).Range
    Next
End Sub

Sub FormatHeaderFooter(rng)
    On Error Resume Next
    If IsMissingObject(rng) Then Exit Sub
    rng.Font.NameFarEast = "宋体": rng.Font.Name = "Times New Roman": rng.Font.Size = 9
    rng.ParagraphFormat.SpaceBefore = 0: rng.ParagraphFormat.SpaceAfter = 0
    Err.Clear
End Sub

Sub FormatParagraphs(appObj, scopeRange, ByRef changedTitle, ByRef changedHeading, ByRef changedBody, ByRef changedList, ByRef failureCount)
    On Error Resume Next
    Dim para, textValue, nonEmptyIndex, headingKind
    nonEmptyIndex = 0
    For Each para In scopeRange.Paragraphs
        textValue = CleanParagraphText(para.Range.Text)
        If Len(textValue) > 0 Then
            nonEmptyIndex = nonEmptyIndex + 1
            headingKind = DetectHeadingKind(para, textValue, nonEmptyIndex)
            Err.Clear
            If headingKind = 1 Then
                ApplyTitleFormat para: changedTitle = changedTitle + 1
            ElseIf headingKind > 1 Then
                ApplyHeadingFormat para, headingKind: changedHeading = changedHeading + 1
            ElseIf IsListParagraph(para) Then
                ApplyListFormat appObj, para: changedList = changedList + 1
            Else
                ApplyBodyFormat appObj, para: changedBody = changedBody + 1
            End If
            If Err.Number <> 0 Then failureCount = failureCount + 1
            Err.Clear
        End If
    Next
End Sub

Sub ApplyTitleFormat(para)
    On Error Resume Next
    With para.Range.Font
        .NameFarEast = "黑体": .Name = "Arial": .Size = 16: .Bold = True: .Italic = False
    End With
    With para.Format
        .Alignment = wdAlignParagraphCenter: .FirstLineIndent = 0: .LeftIndent = 0: .RightIndent = 0
        .SpaceBefore = 0: .SpaceAfter = 6: .LineSpacingRule = wdLineSpaceSingle
        .KeepWithNext = True: .KeepTogether = True
    End With
End Sub

Sub ApplyHeadingFormat(para, headingKind)
    On Error Resume Next
    Dim farEastName, fontSize, beforeSpace, afterSpace
    farEastName = "黑体": fontSize = 14: beforeSpace = 8: afterSpace = 3
    If headingKind = 3 Then farEastName = "楷体": fontSize = 12: beforeSpace = 6
    If headingKind >= 4 Then farEastName = "仿宋": fontSize = 12: beforeSpace = 3: afterSpace = 2
    With para.Range.Font
        .NameFarEast = farEastName: .Name = "Times New Roman": .Size = fontSize: .Bold = True: .Italic = False
    End With
    With para.Format
        .Alignment = wdAlignParagraphLeft: .FirstLineIndent = 0: .LeftIndent = 0: .RightIndent = 0
        .SpaceBefore = beforeSpace: .SpaceAfter = afterSpace: .LineSpacingRule = wdLineSpaceSingle
        .KeepWithNext = True: .KeepTogether = True
    End With
End Sub

Sub ApplyBodyFormat(appObj, para)
    On Error Resume Next
    With para.Range.Font
        .NameFarEast = "宋体": .Name = "Times New Roman": .Size = 10.5: .Bold = False: .Italic = False
    End With
    With para.Format
        .Alignment = wdAlignParagraphJustify: .FirstLineIndent = appObj.CentimetersToPoints(0.74)
        .LeftIndent = 0: .RightIndent = 0: .SpaceBefore = 0: .SpaceAfter = 0
        .LineSpacingRule = wdLineSpace1pt5: .KeepWithNext = False: .KeepTogether = False
    End With
End Sub

Sub ApplyListFormat(appObj, para)
    On Error Resume Next
    With para.Range.Font
        .NameFarEast = "宋体": .Name = "Times New Roman": .Size = 10.5
    End With
    With para.Format
        .Alignment = wdAlignParagraphLeft: .FirstLineIndent = appObj.CentimetersToPoints(-0.50)
        .LeftIndent = appObj.CentimetersToPoints(0.74): .RightIndent = 0
        .SpaceBefore = 0: .SpaceAfter = 0: .LineSpacingRule = wdLineSpaceSingle
        .KeepWithNext = False: .KeepTogether = False
    End With
End Sub

Sub FormatTables(scopeRange, ByRef changedTable, ByRef failureCount)
    On Error Resume Next
    Dim tbl
    For Each tbl In scopeRange.Tables
        Err.Clear
        tbl.AllowAutoFit = True: tbl.AutoFitBehavior wdAutoFitWindow: tbl.Rows.Alignment = wdAlignRowCenter
        tbl.Range.Font.NameFarEast = "宋体": tbl.Range.Font.Name = "Times New Roman": tbl.Range.Font.Size = 10.5
        tbl.Range.ParagraphFormat.SpaceBefore = 0: tbl.Range.ParagraphFormat.SpaceAfter = 0
        If tbl.Rows.Count > 0 Then tbl.Rows(1).Range.Font.Bold = True
        If Err.Number <> 0 Then failureCount = failureCount + 1 Else changedTable = changedTable + 1
        Err.Clear
    Next
End Sub

Sub CountTargets(scopeRange, ByRef titleCount, ByRef headingCount, ByRef bodyCount, ByRef listCount, ByRef tableCount)
    On Error Resume Next
    titleCount = 0: headingCount = 0: bodyCount = 0: listCount = 0: tableCount = 0
    Dim para, textValue, nonEmptyIndex, headingKind, tbl
    nonEmptyIndex = 0
    For Each para In scopeRange.Paragraphs
        textValue = CleanParagraphText(para.Range.Text)
        If Len(textValue) > 0 Then
            nonEmptyIndex = nonEmptyIndex + 1
            headingKind = DetectHeadingKind(para, textValue, nonEmptyIndex)
            If headingKind = 1 Then
                titleCount = titleCount + 1
            ElseIf headingKind > 1 Then
                headingCount = headingCount + 1
            ElseIf IsListParagraph(para) Then
                listCount = listCount + 1
            Else
                bodyCount = bodyCount + 1
            End If
        End If
    Next
    For Each tbl In scopeRange.Tables
        tableCount = tableCount + 1
    Next
End Sub

Function DetectHeadingKind(para, textValue, nonEmptyIndex)
    On Error Resume Next
    DetectHeadingKind = 0
    If nonEmptyIndex = 1 And Len(textValue) <= 80 And Not IsNumberedHeadingText(textValue) Then DetectHeadingKind = 1: Exit Function
    Dim outlineLevel, styleText
    outlineLevel = wdOutlineLevelBodyText: outlineLevel = CLng(para.OutlineLevel): styleText = CStr(para.Style)
    If outlineLevel >= 1 And outlineLevel < 10 Then DetectHeadingKind = outlineLevel + 1: Exit Function
    If InStr(1, styleText, "Heading", vbTextCompare) > 0 Or InStr(1, styleText, "标题", vbTextCompare) > 0 Then DetectHeadingKind = 2: Exit Function
    If MatchesPattern(textValue, "第*章*") Or MatchesPattern(textValue, "一、*") Or MatchesPattern(textValue, "二、*") Or MatchesPattern(textValue, "三、*") Or MatchesPattern(textValue, "四、*") Or MatchesPattern(textValue, "五、*") Or MatchesPattern(textValue, "六、*") Or MatchesPattern(textValue, "七、*") Or MatchesPattern(textValue, "八、*") Or MatchesPattern(textValue, "九、*") Or MatchesPattern(textValue, "十、*") Then
        DetectHeadingKind = 2
    ElseIf Left(textValue, 1) = "（" And InStr(textValue, "）") > 1 Then
        DetectHeadingKind = 3
    ElseIf IsArabicHeading(textValue) Then
        DetectHeadingKind = 4
    End If
    Err.Clear
End Function

Function IsNumberedHeadingText(textValue)
    IsNumberedHeadingText = (MatchesPattern(textValue, "第*章*") Or MatchesPattern(textValue, "一、*") Or (Left(textValue, 1) = "（" And InStr(textValue, "）") > 1) Or IsArabicHeading(textValue))
End Function

Function IsArabicHeading(textValue)
    Dim i, ch, dotPos
    IsArabicHeading = False
    dotPos = InStr(textValue, "."): If dotPos = 0 Then dotPos = InStr(textValue, "．")
    If dotPos <= 1 Or dotPos > 4 Then Exit Function
    For i = 1 To dotPos - 1
        ch = Mid(textValue, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next
    IsArabicHeading = True
End Function

Function MatchesPattern(textValue, pattern)
    On Error Resume Next
    MatchesPattern = (textValue Like pattern)
    Err.Clear
End Function

Function IsListParagraph(para)
    On Error Resume Next
    IsListParagraph = (CLng(para.Range.ListFormat.ListType) <> 0)
    Err.Clear
End Function

Function GetScopeRange(appObj, doc, scopeMode)
    On Error Resume Next
    Set GetScopeRange = Nothing
    If scopeMode = "selection" Then Set GetScopeRange = appObj.Selection.Range.Duplicate Else Set GetScopeRange = doc.Content.Duplicate
    If Err.Number <> 0 Then Set GetScopeRange = Nothing: Err.Clear
End Function

Function CleanParagraphText(value)
    Dim text
    text = CStr(value)
    text = Replace(text, vbCr, ""): text = Replace(text, vbLf, "")
    text = Replace(text, ChrW(7), ""): text = Replace(text, ChrW(160), " "): text = Replace(text, ChrW(12288), " ")
    CleanParagraphText = Trim(text)
End Function

Function NormalizeMode(value)
    Dim text
    text = LCase(Trim(CStr(value)))
    If text = "apply" Or text = "应用" Or text = "a" Then NormalizeMode = "apply" Else NormalizeMode = "preview"
End Function

Function NormalizeScope(value)
    Dim text
    text = LCase(Trim(CStr(value)))
    If text = "selection" Or text = "选区" Or text = "s" Then NormalizeScope = "selection" Else NormalizeScope = "document"
End Function

Function BuildPreviewText(scopeMode, titleCount, headingCount, bodyCount, listCount, tableCount, issueSummary)
    BuildPreviewText = "Word 专业排版预览（尚未写入）/ Word professional formatting preview (no changes yet)" & vbCrLf & _
        "范围 / Scope=" & scopeMode & "；标题 / title=" & CStr(titleCount) & "；层级标题 / headings=" & CStr(headingCount) & "；正文 / body=" & CStr(bodyCount) & "；列表 / lists=" & CStr(listCount) & "；表格 / tables=" & CStr(tableCount) & vbCrLf & _
        "规则 / Rules：A4 页面、宋体正文 10.5 磅、1.5 倍行距、标题识别、表格窗口自适应；正文文字保持不变 / text preserved"
    If Len(issueSummary) > 0 Then BuildPreviewText = BuildPreviewText & vbCrLf & "当前排版摘要 / Current issue summary: " & issueSummary
End Function

Function BuildPlanParams(scopeMode, titleCount, headingCount, bodyCount, listCount, tableCount)
    BuildPlanParams = "{""scope"":""" & EscapeJson(scopeMode) & """,""title"":" & CStr(titleCount) & ",""headings"":" & CStr(headingCount) & ",""body"":" & CStr(bodyCount) & ",""lists"":" & CStr(listCount) & ",""tables"":" & CStr(tableCount) & "}"
End Function

Function SuccessJson(applied, scopeMode, titleCount, headingCount, bodyCount, listCount, tableCount, planPreview, planValidation, message, contextSummary)
    SuccessJson = "{""ok"":true,""applied"":" & JsonBool(applied) & ",""scope"":""" & EscapeJson(scopeMode) & """,""counts"":{""title"":" & CStr(titleCount) & ",""headings"":" & CStr(headingCount) & ",""body"":" & CStr(bodyCount) & ",""lists"":" & CStr(listCount) & ",""tables"":" & CStr(tableCount) & "},""writePlanPreview"":""" & EscapeJson(planPreview) & """,""writePlanValidation"":""" & EscapeJson(planValidation) & """,""message"":""" & EscapeJson(message) & """,""context"":""" & EscapeJson(contextSummary) & """}"
End Function

Function FailureJson(code, message)
    FailureJson = "{""ok"":false,""code"":""" & EscapeJson(code) & """,""message"":""" & EscapeJson(message) & """}"
End Function

Function FailureWithStateJson(code, message, contentPreserved, failureCount)
    FailureWithStateJson = "{""ok"":false,""code"":""" & EscapeJson(code) & """,""contentPreserved"":" & JsonBool(contentPreserved) & ",""failureCount"":" & CStr(failureCount) & ",""message"":""" & EscapeJson(message) & """}"
End Function

Function IsDocumentProtected(doc)
    On Error Resume Next
    IsDocumentProtected = False
    If CLng(doc.ProtectionType) <> -1 Then IsDocumentProtected = True
    Err.Clear
End Function

Function IsMissingObject(value)
    On Error Resume Next
    IsMissingObject = (TypeName(value) = "Nothing" Or TypeName(value) = "Empty")
    Err.Clear
End Function

Function SafePrompt(message, defaultValue)
    On Error Resume Next
    SafePrompt = Host.Prompt(message, defaultValue)
    If Err.Number <> 0 Then SafePrompt = defaultValue
    Err.Clear
End Function

Function SafeHostText(methodName)
    On Error Resume Next
    SafeHostText = ""
    If methodName = "GetWordFormatIssueSummary" Then
        SafeHostText = Host.GetWordFormatIssueSummary()
    ElseIf methodName = "GetWordContextInfo" Then
        SafeHostText = Host.GetWordContextInfo()
    End If
    If Err.Number <> 0 Then SafeHostText = ""
    Err.Clear
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

Sub SafeRollbackWritePlan(planId)
    On Error Resume Next
    Host.RollbackWritePlan planId
    Err.Clear
End Sub

Sub SafeWriteLog(message)
    On Error Resume Next
    Host.WriteLog message
    Err.Clear
End Sub

Function JsonBool(value)
    If value Then JsonBool = "true" Else JsonBool = "false"
End Function

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

