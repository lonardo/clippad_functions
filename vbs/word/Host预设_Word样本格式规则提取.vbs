' 函数名: HostWordExtractSampleFormatRules
' 描述: 从用户选择的 Word 样本文档只读提取标题、正文、表格和页眉页脚的可迁移格式规则，并统计当前文档的潜在影响范围；不会修改任何文档
' 适用应用: Word
' 搜索范围: 全文
' 搜索对象: 无

Option Explicit

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

    Main = HostWordExtractSampleFormatRules(appObj)
End Function

Function HostWordExtractSampleFormatRules(appObj)
    On Error Resume Next

    Dim targetDoc
    Set targetDoc = appObj.ActiveDocument
    If Err.Number <> 0 Or TypeName(targetDoc) = "Empty" Or TypeName(targetDoc) = "Nothing" Then
        Err.Clear
        HostWordExtractSampleFormatRules = "{""ok"":false,""code"":""E_NO_DOCUMENT"",""message"":""当前没有活动 Word 文档""}"
        Exit Function
    End If

    Dim samplePath
    samplePath = Trim(SafeSelectFile("选择规范 Word 样本文档（本次只读取规则，不会修改样本或当前文档）", "Word 文档|*.doc;*.docx;*.docm"))
    If Len(samplePath) = 0 Then
        HostWordExtractSampleFormatRules = "{""ok"":false,""code"":""E_SAMPLE_REQUIRED"",""message"":""未选择样本文档；未修改当前文档""}"
        Exit Function
    End If

    Dim sampleDoc
    Set sampleDoc = Nothing
    Err.Clear
    Set sampleDoc = appObj.Documents.Open(samplePath, False, True)
    If Err.Number <> 0 Or TypeName(sampleDoc) = "Empty" Or TypeName(sampleDoc) = "Nothing" Then
        Dim openError
        openError = Err.Description
        Err.Clear
        HostWordExtractSampleFormatRules = "{""ok"":false,""code"":""E_SAMPLE_OPEN"",""message"":""无法以只读方式打开样本文档：""" & EscapeJson(openError) & """}"
        Exit Function
    End If

    Dim titleRule, headingRule, bodyRule, tableRule, headerFooterRule
    titleRule = ExtractFirstNonEmptyParagraphRule(sampleDoc)
    headingRule = ExtractHeadingRule(sampleDoc)
    bodyRule = ExtractBodyRule(sampleDoc)
    tableRule = ExtractTableRule(sampleDoc)
    headerFooterRule = ExtractHeaderFooterRule(sampleDoc)

    Dim targetTitleCount, targetHeadingCount, targetBodyCount
    CountTargetParagraphRoles targetDoc, targetTitleCount, targetHeadingCount, targetBodyCount

    Dim targetTables, targetSections, sampleTables, sampleSections
    targetTables = SafeCount(targetDoc.Tables.Count)
    targetSections = SafeCount(targetDoc.Sections.Count)
    sampleTables = SafeCount(sampleDoc.Tables.Count)
    sampleSections = SafeCount(sampleDoc.Sections.Count)

    Dim caution
    caution = "正式应用前必须先预览并确认；本次仅提取规则。"
    If targetTables = 0 And sampleTables > 0 Then caution = caution & " 当前文档没有表格，表格规则将无法应用。"
    If targetSections = 0 Or sampleSections = 0 Then caution = caution & " 页眉页脚规则需要至少一个有效节。"

    Dim summary
    summary = "Word 样本格式规则提取完成（只读）；样本=" & sampleDoc.Name & _
        "；目标标题候选=" & CStr(targetTitleCount) & "；标题层级候选=" & CStr(targetHeadingCount) & _
        "；正文候选=" & CStr(targetBodyCount) & "；目标表格=" & CStr(targetTables) & _
        "；目标节=" & CStr(targetSections) & vbCrLf & _
        "规则：标题=" & RuleLabel(titleRule) & "；层级标题=" & RuleLabel(headingRule) & _
        "；正文=" & RuleLabel(bodyRule) & "；表格=" & RuleLabel(tableRule) & _
        "；页眉页脚=" & RuleLabel(headerFooterRule) & vbCrLf & _
        "提示：" & caution
    Host.WriteClipboard summary
    SafeWriteLog summary

    On Error Resume Next
    sampleDoc.Close False
    Err.Clear

    HostWordExtractSampleFormatRules = "{""ok"":true,""readOnly"":true,""samplePath"":""" & EscapeJson(samplePath) & _
        """,""impact"":{""titleCandidates"":" & CStr(targetTitleCount) & ",""headingCandidates"":" & CStr(targetHeadingCount) & _
        ",""bodyCandidates"":" & CStr(targetBodyCount) & ",""tables"":" & CStr(targetTables) & ",""sections"":" & CStr(targetSections) & _
        "},""rules"":{""title"":" & titleRule & ",""heading"":" & headingRule & ",""body"":" & bodyRule & _
        ",""table"":" & tableRule & ",""headerFooter"":" & headerFooterRule & _
        "},""requiresPreview"":true,""requiresConfirmation"":true,""message"":""" & EscapeJson(summary) & """}"
End Function

Function ExtractFirstNonEmptyParagraphRule(doc)
    Dim para, textValue
    ExtractFirstNonEmptyParagraphRule = "null"
    For Each para In doc.Paragraphs
        textValue = CleanParagraphText(para.Range.Text)
        If Len(textValue) > 0 Then
            ExtractFirstNonEmptyParagraphRule = BuildParagraphRule(para, "title")
            Exit Function
        End If
    Next
End Function

Function ExtractHeadingRule(doc)
    Dim para, textValue
    ExtractHeadingRule = "null"
    For Each para In doc.Paragraphs
        textValue = CleanParagraphText(para.Range.Text)
        If Len(textValue) > 0 And IsHeadingParagraph(para) Then
            ExtractHeadingRule = BuildParagraphRule(para, "heading")
            Exit Function
        End If
    Next
End Function

Function ExtractBodyRule(doc)
    Dim para, textValue
    ExtractBodyRule = "null"
    For Each para In doc.Paragraphs
        textValue = CleanParagraphText(para.Range.Text)
        If Len(textValue) > 0 And Not IsHeadingParagraph(para) Then
            ExtractBodyRule = BuildParagraphRule(para, "body")
            Exit Function
        End If
    Next
End Function

Function ExtractTableRule(doc)
    On Error Resume Next
    ExtractTableRule = "null"
    If doc.Tables.Count <= 0 Then Exit Function
    Dim tbl
    Set tbl = doc.Tables(1)
    ExtractTableRule = "{""present"":true,""rows"":" & CStr(SafeCount(tbl.Rows.Count)) & _
        ",""columns"":" & CStr(SafeCount(tbl.Columns.Count)) & _
        ",""allowAutoFit"":" & JsonBool(CBool(tbl.AllowAutoFit)) & _
        ",""alignment"":" & CStr(SafeLong(tbl.Rows.Alignment)) & "}"
    If Err.Number <> 0 Then
        ExtractTableRule = "{""present"":true,""detailsAvailable"":false}"
        Err.Clear
    End If
End Function

Function ExtractHeaderFooterRule(doc)
    On Error Resume Next
    ExtractHeaderFooterRule = "null"
    If doc.Sections.Count <= 0 Then Exit Function
    Dim sectionObj, primaryHeader, primaryFooter
    Set sectionObj = doc.Sections(1)
    Set primaryHeader = sectionObj.Headers(1)
    Set primaryFooter = sectionObj.Footers(1)
    ExtractHeaderFooterRule = "{""present"":true,""headerHasText"":" & JsonBool(Len(CleanParagraphText(primaryHeader.Range.Text)) > 0) & _
        ",""footerHasText"":" & JsonBool(Len(CleanParagraphText(primaryFooter.Range.Text)) > 0) & "}"
    If Err.Number <> 0 Then
        ExtractHeaderFooterRule = "{""present"":true,""detailsAvailable"":false}"
        Err.Clear
    End If
End Function

Function BuildParagraphRule(para, roleName)
    On Error Resume Next
    BuildParagraphRule = "{""role"":""" & EscapeJson(roleName) & _
        """,""fontEastAsia"":""" & EscapeJson(CStr(para.Range.Font.NameFarEast)) & _
        """,""fontLatin"":""" & EscapeJson(CStr(para.Range.Font.Name)) & _
        """,""fontSize"":" & JsonNumber(para.Range.Font.Size) & _
        ",""bold"":" & JsonBool(CBool(para.Range.Font.Bold)) & _
        ",""alignment"":" & CStr(SafeLong(para.Format.Alignment)) & _
        ",""spaceBefore"":" & JsonNumber(para.Format.SpaceBefore) & _
        ",""spaceAfter"":" & JsonNumber(para.Format.SpaceAfter) & _
        ",""firstLineIndent"":" & JsonNumber(para.Format.FirstLineIndent) & _
        ",""lineSpacingRule"":" & CStr(SafeLong(para.Format.LineSpacingRule)) & "}"
    If Err.Number <> 0 Then
        BuildParagraphRule = "{""role"":""" & EscapeJson(roleName) & """,""detailsAvailable"":false}"
        Err.Clear
    End If
End Function

Sub CountTargetParagraphRoles(doc, ByRef titleCount, ByRef headingCount, ByRef bodyCount)
    Dim para, textValue, nonEmptyCount
    titleCount = 0
    headingCount = 0
    bodyCount = 0
    nonEmptyCount = 0
    For Each para In doc.Paragraphs
        textValue = CleanParagraphText(para.Range.Text)
        If Len(textValue) > 0 Then
            nonEmptyCount = nonEmptyCount + 1
            If nonEmptyCount = 1 Then titleCount = titleCount + 1
            If IsHeadingParagraph(para) Then
                headingCount = headingCount + 1
            Else
                bodyCount = bodyCount + 1
            End If
        End If
    Next
End Sub

Function IsHeadingParagraph(para)
    On Error Resume Next
    IsHeadingParagraph = False
    If CLng(para.OutlineLevel) <> wdOutlineLevelBodyText Then
        IsHeadingParagraph = True
    ElseIf InStr(1, CStr(para.Style), "标题", vbTextCompare) > 0 Or InStr(1, CStr(para.Style), "Heading", vbTextCompare) > 0 Then
        IsHeadingParagraph = True
    End If
    If Err.Number <> 0 Then
        IsHeadingParagraph = False
        Err.Clear
    End If
End Function

Function RuleLabel(ruleJson)
    If ruleJson = "null" Then
        RuleLabel = "未发现"
    Else
        RuleLabel = "已提取"
    End If
End Function

Function CleanParagraphText(value)
    Dim text
    text = CStr(value)
    text = Replace(text, vbCr, "")
    text = Replace(text, vbLf, "")
    text = Replace(text, ChrW(7), "")
    text = Replace(text, ChrW(160), " ")
    text = Replace(text, ChrW(12288), " ")
    CleanParagraphText = Trim(text)
End Function

Function SafeSelectFile(title, filter)
    On Error Resume Next
    SafeSelectFile = Host.SelectFile(title, filter)
    If Err.Number <> 0 Then
        SafeSelectFile = ""
        Err.Clear
    End If
End Function

Function SafeCount(value)
    On Error Resume Next
    SafeCount = CLng(value)
    If Err.Number <> 0 Then
        SafeCount = 0
        Err.Clear
    End If
End Function

Function SafeLong(value)
    On Error Resume Next
    SafeLong = CLng(value)
    If Err.Number <> 0 Then
        SafeLong = 0
        Err.Clear
    End If
End Function

Function JsonNumber(value)
    On Error Resume Next
    JsonNumber = Replace(CStr(CDbl(value)), ",", ".")
    If Err.Number <> 0 Then
        JsonNumber = "0"
        Err.Clear
    End If
End Function

Function JsonBool(value)
    If value Then
        JsonBool = "true"
    Else
        JsonBool = "false"
    End If
End Function

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
