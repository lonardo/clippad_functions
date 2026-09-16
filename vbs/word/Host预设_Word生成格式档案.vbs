' 函数名: HostWordGenerateFormatProfile
' 描述: 从用户确认的 Word 样本文档提取格式规则，生成带来源、版本、适用范围和运行时样本路径的版本化格式档案；不会修改文档内容
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
        Main = FailureJson("E_NO_WORD_APP", "未取得 Word 应用，请在 Word 中运行该预设")
        Exit Function
    End If
    Main = HostWordGenerateFormatProfile(appObj)
End Function

Function HostWordGenerateFormatProfile(appObj)
    On Error Resume Next
    Dim samplePath
    samplePath = Trim(SafeSelectFile("选择要登记为格式档案的 Word 样本文档（以只读方式打开）", "Word 文档|*.doc;*.docx;*.docm"))
    If Len(samplePath) = 0 Then
        HostWordGenerateFormatProfile = FailureJson("E_SAMPLE_REQUIRED", "未选择样本文档；未生成格式档案")
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
        HostWordGenerateFormatProfile = FailureJson("E_SAMPLE_OPEN", "无法以只读方式打开样本文档：" & openError)
        Exit Function
    End If

    Dim profileId, displayName, version, documentTypes, domain, sourceLabel, authority, confirmation
    profileId = NormalizeId(SafePrompt("格式档案 ID（例如 legal.user_template.v1）", "legal.user_template"))
    displayName = LimitText(SafePrompt("格式档案名称", "用户模板格式档案"), 80)
    version = NormalizeVersion(SafePrompt("格式档案版本（例如 1.0.0）", "1.0.0"))
    documentTypes = LimitText(SafePrompt("适用文书类型，用逗号分隔（例如 起诉状,答辩状）", "用户确认的法律文书"), 160)
    domain = NormalizeDomain(SafePrompt("领域：法律、公文、合同、报告或自定义", "法律"))
    sourceLabel = LimitText(SafePrompt("模板来源说明（例如 律所内部模板）", "用户提供的本地模板"), 120)
    authority = LimitText(SafePrompt("来源机构/确认人（可留空）", ""), 120)
    confirmation = Trim(SafePrompt("确认已核对模板来源、版本和适用范围？请输入：确认；否则档案保持草稿", ""))

    Dim outputFolder
    outputFolder = Trim(SafeSelectFolder("选择格式档案保存文件夹"))
    If Len(outputFolder) = 0 Then
        SafeCloseDocument sampleDoc
        HostWordGenerateFormatProfile = FailureJson("E_OUTPUT_FOLDER_REQUIRED", "未选择档案保存文件夹；未生成格式档案")
        Exit Function
    End If

    Dim fileName, outputPath
    fileName = Host.SanitizeFileName(profileId & "_" & version, "format_profile")
    fileName = Host.EnsureFileExtension(fileName, ".json")
    outputPath = Host.CombinePath(outputFolder, fileName)
    If Host.PathExists(outputPath) Then
        SafeCloseDocument sampleDoc
        HostWordGenerateFormatProfile = FailureJson("E_OUTPUT_EXISTS", "目标格式档案已存在，请提高版本号或更换文件名：" & outputPath)
        Exit Function
    End If

    Dim titlePara, headingPara, bodyPara, tableObj, sectionObj
    Set titlePara = FindFirstNonEmptyParagraph(sampleDoc.Content)
    Set headingPara = FindFirstHeadingParagraph(sampleDoc.Content)
    Set bodyPara = FindFirstBodyParagraph(sampleDoc.Content, titlePara)
    Set tableObj = FindFirstTable(sampleDoc)
    Set sectionObj = FindFirstSection(sampleDoc)

    Dim status, isConfirmed
    isConfirmed = False
    If StrComp(confirmation, "确认", vbTextCompare) = 0 Then
        isConfirmed = True
    End If
    If isConfirmed Then
        status = "user_confirmed"
    Else
        status = "draft"
    End If

    Dim profileJson, writeOk, sampleInfo
    sampleInfo = SafeHostText("GetFileInfo", samplePath)
    profileJson = BuildProfileJson(profileId, displayName, domain, status, version, sourceLabel, authority, documentTypes, samplePath, sampleDoc.Name, sampleInfo, titlePara, headingPara, bodyPara, tableObj, sectionObj)
    writeOk = Host.WriteTextFile(outputPath, profileJson, False)
    If Not CBool(writeOk) Then
        SafeCloseDocument sampleDoc
        HostWordGenerateFormatProfile = FailureJson("E_PROFILE_WRITE", "格式档案写入失败：" & outputPath)
        Exit Function
    End If

    Dim summary
    summary = "Word 格式档案已生成；状态=" & status & "；版本=" & version & "；类型=" & documentTypes & _
        "；样本=" & sampleDoc.Name & "；路径=" & outputPath & vbCrLf & _
        "规则来源：仅来自用户选择的样本文档；业务正文、页眉页脚文字和页码字段不会写入档案。"
    Host.WriteClipboard summary
    SafeWriteLog summary
    SafeCloseDocument sampleDoc
    HostWordGenerateFormatProfile = "{""ok"":true,""profilePath"":""" & EscapeJson(outputPath) & _
        """,""profileId"":""" & EscapeJson(profileId) & """,""status"":""" & status & _
        """,""profileVersion"":""" & version & """,""confirmedByUser"":" & JsonBool(isConfirmed) & _
        ",""message"":""" & EscapeJson(summary) & """}"
End Function

Function BuildProfileJson(profileId, displayName, domain, status, version, sourceLabel, authority, documentTypes, samplePath, sampleName, sampleInfo, titlePara, headingPara, bodyPara, tableObj, sectionObj)
    Dim json, sourceJson, runtimeJson, rulesJson, guardrailsJson
    sourceJson = "{" & JsonQuote("kind") & ":" & JsonQuote("user_template") & _
        "," & JsonQuote("label") & ":" & JsonQuote(sourceLabel) & _
        "," & JsonQuote("version") & ":" & JsonQuote(version) & _
        "," & JsonQuote("authority") & ":" & JsonQuote(authority)
    If status = "user_confirmed" Then
        sourceJson = sourceJson & "," & JsonQuote("confirmedByUser") & ":true}"
    Else
        sourceJson = sourceJson & "," & JsonQuote("confirmedByUser") & ":false}"
    End If
    runtimeJson = "{" & JsonQuote("samplePath") & ":" & JsonQuote(samplePath) & _
        "," & JsonQuote("sampleDocumentName") & ":" & JsonQuote(sampleName) & _
        "," & JsonQuote("sampleFileInfo") & ":" & JsonQuote(sampleInfo) & "}"
    rulesJson = "{" & JsonQuote("title") & ":" & BuildRuleJson(titlePara, "title", True) & _
        "," & JsonQuote("heading") & ":" & BuildRuleJson(headingPara, "heading", True) & _
        "," & JsonQuote("body") & ":" & BuildRuleJson(bodyPara, "body", True) & _
        "," & JsonQuote("table") & ":" & BuildTableRuleJson(tableObj) & _
        "," & JsonQuote("page") & ":" & BuildPageRuleJson(sectionObj) & _
        "," & JsonQuote("headerFooter") & ":" & BuildHeaderRuleJson(sectionObj) & "}"
    guardrailsJson = "{" & JsonQuote("preserveBusinessContent") & ":true," & _
        JsonQuote("requiresPreview") & ":true," & JsonQuote("requiresConfirmation") & ":true," & _
        JsonQuote("prohibitedOperations") & ":[" & JsonQuote("不要生成或修改业务事实") & "," & _
        JsonQuote("不要生成或修改法律请求、论证或证据结论") & "," & _
        JsonQuote("不要复制样本业务文字、页码字段或附件内容") & "]}"
    json = "{" & JsonQuote("profileVersion") & ":" & JsonQuote(version) & _
        "," & JsonQuote("id") & ":" & JsonQuote(profileId) & _
        "," & JsonQuote("domain") & ":" & JsonQuote(domain) & _
        "," & JsonQuote("displayName") & ":" & JsonQuote(displayName) & _
        "," & JsonQuote("status") & ":" & JsonQuote(status) & _
        "," & JsonQuote("source") & ":" & sourceJson & _
        "," & JsonQuote("applicableDocumentTypes") & ":[" & JsonStringArray(documentTypes) & "]" & _
        "," & JsonQuote("runtime") & ":" & runtimeJson & _
        "," & JsonQuote("formatRules") & ":" & rulesJson & _
        "," & JsonQuote("guardrails") & ":" & guardrailsJson & "}"
    BuildProfileJson = json
End Function

Function BuildRuleJson(para, roleName, requiredRule)
    If IsNothing(para) Then
        BuildRuleJson = "{" & JsonQuote("mode") & ":" & JsonQuote("disabled") & _
            "," & JsonQuote("fields") & ":[]," & JsonQuote("required") & ":" & JsonBool(requiredRule) & "}"
    Else
        BuildRuleJson = "{" & JsonQuote("mode") & ":" & JsonQuote("explicit") & _
            "," & JsonQuote("fields") & ":[" & JsonQuote("font") & "," & JsonQuote("size") & "," & _
            JsonQuote("alignment") & "," & JsonQuote("spacing") & "," & JsonQuote("indent") & "]" & _
            "," & JsonQuote("required") & ":" & JsonBool(requiredRule) & _
            "," & JsonQuote("values") & ":" & BuildParagraphValues(para, roleName) & "}"
    End If
End Function

Function BuildParagraphValues(para, roleName)
    On Error Resume Next
    BuildParagraphValues = "{" & JsonQuote("role") & ":" & JsonQuote(roleName) & _
        "," & JsonQuote("fontEastAsia") & ":" & JsonQuote(CStr(para.Range.Font.NameFarEast)) & _
        "," & JsonQuote("fontLatin") & ":" & JsonQuote(CStr(para.Range.Font.Name)) & _
        "," & JsonQuote("fontSize") & ":" & JsonNumber(para.Range.Font.Size) & _
        "," & JsonQuote("bold") & ":" & JsonBool(CBool(para.Range.Font.Bold)) & _
        "," & JsonQuote("alignment") & ":" & CStr(SafeLong(para.Format.Alignment)) & _
        "," & JsonQuote("spaceBefore") & ":" & JsonNumber(para.Format.SpaceBefore) & _
        "," & JsonQuote("spaceAfter") & ":" & JsonNumber(para.Format.SpaceAfter) & _
        "," & JsonQuote("firstLineIndent") & ":" & JsonNumber(para.Format.FirstLineIndent) & _
        "," & JsonQuote("lineSpacingRule") & ":" & CStr(SafeLong(para.Format.LineSpacingRule)) & "}"
    If Err.Number <> 0 Then
        BuildParagraphValues = "{" & JsonQuote("role") & ":" & JsonQuote(roleName) & _
            "," & JsonQuote("detailsAvailable") & ":false}"
        Err.Clear
    End If
End Function

Function BuildTableRuleJson(tableObj)
    If IsNothing(tableObj) Then
        BuildTableRuleJson = "{" & JsonQuote("mode") & ":" & JsonQuote("disabled") & _
            "," & JsonQuote("fields") & ":[]," & JsonQuote("required") & ":false}"
    Else
        BuildTableRuleJson = "{" & JsonQuote("mode") & ":" & JsonQuote("explicit") & _
            "," & JsonQuote("fields") & ":[" & JsonQuote("font") & "," & JsonQuote("alignment") & "," & JsonQuote("borders") & "," & JsonQuote("cellSpacing") & "]" & _
            "," & JsonQuote("required") & ":false," & JsonQuote("values") & ":{" & _
            JsonQuote("rows") & ":" & CStr(SafeCount(tableObj.Rows.Count)) & _
            "," & JsonQuote("columns") & ":" & CStr(SafeCount(tableObj.Columns.Count)) & _
            "," & JsonQuote("fontEastAsia") & ":" & JsonQuote(CStr(tableObj.Range.Font.NameFarEast)) & _
            "," & JsonQuote("fontSize") & ":" & JsonNumber(tableObj.Range.Font.Size) & "}}"
    End If
End Function

Function BuildPageRuleJson(sectionObj)
    If IsNothing(sectionObj) Then
        BuildPageRuleJson = "{" & JsonQuote("mode") & ":" & JsonQuote("disabled") & _
            "," & JsonQuote("fields") & ":[]," & JsonQuote("required") & ":false}"
    Else
        BuildPageRuleJson = "{" & JsonQuote("mode") & ":" & JsonQuote("explicit") & _
            "," & JsonQuote("fields") & ":[" & JsonQuote("paper") & "," & JsonQuote("margins") & "," & JsonQuote("orientation") & "]" & _
            "," & JsonQuote("required") & ":false," & JsonQuote("values") & ":{" & _
            JsonQuote("topMargin") & ":" & JsonNumber(sectionObj.PageSetup.TopMargin) & _
            "," & JsonQuote("bottomMargin") & ":" & JsonNumber(sectionObj.PageSetup.BottomMargin) & _
            "," & JsonQuote("leftMargin") & ":" & JsonNumber(sectionObj.PageSetup.LeftMargin) & _
            "," & JsonQuote("rightMargin") & ":" & JsonNumber(sectionObj.PageSetup.RightMargin) & _
            "," & JsonQuote("orientation") & ":" & CStr(SafeLong(sectionObj.PageSetup.Orientation)) & "}}"
    End If
End Function

Function BuildHeaderRuleJson(sectionObj)
    If IsNothing(sectionObj) Then
        BuildHeaderRuleJson = "{" & JsonQuote("mode") & ":" & JsonQuote("disabled") & _
            "," & JsonQuote("fields") & ":[]," & JsonQuote("required") & ":false}"
    Else
        BuildHeaderRuleJson = "{" & JsonQuote("mode") & ":" & JsonQuote("explicit") & _
            "," & JsonQuote("fields") & ":[" & JsonQuote("font") & "," & JsonQuote("alignment") & "," & JsonQuote("spacing") & "]" & _
            "," & JsonQuote("required") & ":false," & JsonQuote("values") & ":{" & _
            JsonQuote("headerTextExcluded") & ":true," & JsonQuote("footerTextExcluded") & ":true}}"
    End If
End Function

Function JsonStringArray(value)
    Dim parts, i, item, result
    parts = Split(Replace(CStr(value), "，", ","), ",")
    result = ""
    For i = 0 To UBound(parts)
        item = Trim(parts(i))
        If Len(item) > 0 Then
            If Len(result) > 0 Then result = result & ","
            result = result & """" & EscapeJson(item) & """"
        End If
    Next
    If Len(result) = 0 Then result = """用户确认文档"""
    JsonStringArray = result
End Function

Function NormalizeId(value)
    Dim text
    text = LCase(Trim(CStr(value)))
    text = Replace(text, " ", "_")
    If Len(text) = 0 Then text = "custom.user_template"
    NormalizeId = Left(text, 80)
End Function

Function NormalizeVersion(value)
    Dim text
    text = Trim(CStr(value))
    If Not IsVersionLike(text) Then text = "1.0.0"
    NormalizeVersion = text
End Function

Function IsVersionLike(value)
    Dim parts
    parts = Split(CStr(value), ".")
    IsVersionLike = (UBound(parts) = 2 And IsNumeric(parts(0)) And IsNumeric(parts(1)) And IsNumeric(parts(2)))
End Function

Function NormalizeDomain(value)
    Dim text
    text = LCase(Trim(CStr(value)))
    If InStr(text, "法律") > 0 Or text = "legal" Then
        NormalizeDomain = "legal"
    ElseIf InStr(text, "公文") > 0 Or text = "official" Then
        NormalizeDomain = "official"
    ElseIf InStr(text, "合同") > 0 Or text = "contract" Then
        NormalizeDomain = "contract"
    ElseIf InStr(text, "报告") > 0 Or text = "report" Then
        NormalizeDomain = "report"
    Else
        NormalizeDomain = "custom"
    End If
End Function

Function FindFirstNonEmptyParagraph(sourceRange)
    Dim para
    Set FindFirstNonEmptyParagraph = Nothing
    For Each para In sourceRange.Paragraphs
        If Len(CleanParagraphText(para.Range.Text)) > 0 Then Set FindFirstNonEmptyParagraph = para: Exit Function
    Next
End Function

Function FindFirstHeadingParagraph(sourceRange)
    Dim para
    Set FindFirstHeadingParagraph = Nothing
    For Each para In sourceRange.Paragraphs
        If Len(CleanParagraphText(para.Range.Text)) > 0 And IsHeadingParagraph(para) Then Set FindFirstHeadingParagraph = para: Exit Function
    Next
End Function

Function FindFirstBodyParagraph(sourceRange, titlePara)
    Dim para
    Set FindFirstBodyParagraph = Nothing
    For Each para In sourceRange.Paragraphs
        If Len(CleanParagraphText(para.Range.Text)) > 0 Then
            If (IsNothing(titlePara) Or SafeLong(para.Range.Start) <> SafeLong(titlePara.Range.Start)) And Not IsHeadingParagraph(para) Then Set FindFirstBodyParagraph = para: Exit Function
        End If
    Next
End Function

Function FindFirstTable(doc)
    Set FindFirstTable = Nothing
    If SafeCount(doc.Tables.Count) > 0 Then Set FindFirstTable = doc.Tables(1)
End Function

Function FindFirstSection(doc)
    Set FindFirstSection = Nothing
    If SafeCount(doc.Sections.Count) > 0 Then Set FindFirstSection = doc.Sections(1)
End Function

Function IsHeadingParagraph(para)
    On Error Resume Next
    IsHeadingParagraph = (CLng(para.OutlineLevel) <> wdOutlineLevelBodyText)
    If Err.Number <> 0 Then IsHeadingParagraph = (InStr(1, CStr(para.Style), "标题", vbTextCompare) > 0)
    Err.Clear
End Function

Function IsNothing(value)
    IsNothing = (TypeName(value) = "Nothing" Or TypeName(value) = "Empty")
End Function

Function CleanParagraphText(value)
    Dim text
    text = Replace(Replace(Replace(CStr(value), vbCr, ""), vbLf, ""), ChrW(7), "")
    CleanParagraphText = Trim(text)
End Function

Function SafeSelectFile(title, filter)
    On Error Resume Next
    SafeSelectFile = Host.SelectFile(title, filter)
    If Err.Number <> 0 Then SafeSelectFile = ""
    Err.Clear
End Function

Function SafeSelectFolder(title)
    On Error Resume Next
    SafeSelectFolder = Host.SelectFolder(title)
    If Err.Number <> 0 Then SafeSelectFolder = ""
    Err.Clear
End Function

Function SafePrompt(message, defaultValue)
    On Error Resume Next
    SafePrompt = Host.Prompt(message, defaultValue)
    If Err.Number <> 0 Then SafePrompt = defaultValue
    Err.Clear
End Function

Function SafeHostText(methodName, value)
    On Error Resume Next
    If methodName = "GetFileInfo" Then SafeHostText = Host.GetFileInfo(value) Else SafeHostText = ""
    If Err.Number <> 0 Then SafeHostText = ""
    Err.Clear
End Function

Function LimitText(value, maxLength)
    LimitText = Left(CStr(value), CLng(maxLength))
End Function

Sub SafeCloseDocument(doc)
    On Error Resume Next
    doc.Close False
    Err.Clear
End Sub

Sub SafeWriteLog(message)
    On Error Resume Next
    Host.WriteLog message
    Err.Clear
End Sub

Function SafeCount(value)
    On Error Resume Next
    SafeCount = CLng(value)
    If Err.Number <> 0 Then SafeCount = 0
    Err.Clear
End Function

Function SafeLong(value)
    On Error Resume Next
    SafeLong = CLng(value)
    If Err.Number <> 0 Then SafeLong = 0
    Err.Clear
End Function

Function JsonNumber(value)
    On Error Resume Next
    JsonNumber = Replace(CStr(CDbl(value)), ",", ".")
    If Err.Number <> 0 Then JsonNumber = "0"
    Err.Clear
End Function

Function JsonBool(value)
    If value Then JsonBool = "true" Else JsonBool = "false"
End Function

Function JsonQuote(value)
    JsonQuote = """" & EscapeJson(value) & """"
End Function

Function FailureJson(code, message)
    FailureJson = "{""ok"":false,""code"":""" & EscapeJson(code) & """,""message"":""" & EscapeJson(message) & """}"
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
