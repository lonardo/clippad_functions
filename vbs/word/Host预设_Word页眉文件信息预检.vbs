' 函数名: HostWordHeaderFileInfoPreflight
' 描述: 预检各节页眉/页脚现状后，确认词“写入”才把文件名、路径或日期信息写入页眉；默认不改页脚，可预览后取消
' 适用应用: Word
' 搜索范围: 全文
' 搜索对象: 无

Option Explicit

Const wdHeaderFooterPrimary = 1
Const wdHeaderFooterFirstPage = 2
Const wdHeaderFooterEvenPages = 3
Const wdAlignParagraphCenter = 1
Const wdAlignParagraphLeft = 0
Const wdAlignParagraphRight = 2

Function Main()
    On Error Resume Next
    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = FailureJson("E_NO_WORD_APP", "未取得 Word 应用，请在 Word 中运行该预设")
        Exit Function
    End If
    Main = HostWordHeaderFileInfoPreflight(appObj)
End Function

Function HostWordHeaderFileInfoPreflight(appObj)
    On Error Resume Next

    Dim doc
    Set doc = appObj.ActiveDocument
    If Err.Number <> 0 Or TypeName(doc) = "Empty" Or TypeName(doc) = "Nothing" Then
        Err.Clear
        HostWordHeaderFileInfoPreflight = FailureJson("E_NO_DOCUMENT", "当前没有活动 Word 文档")
        Exit Function
    End If

    Dim modeText, modeCode, alignText, alignCode
    modeText = Trim(SafePrompt("写入内容：文件名 / 文件名和路径 / 文件名和日期", "文件名"))
    modeCode = NormalizeHeaderMode(modeText)
    If Len(modeCode) = 0 Then
        HostWordHeaderFileInfoPreflight = FailureJson("E_MODE_INVALID", "写入内容仅支持：文件名 / 文件名和路径 / 文件名和日期")
        Exit Function
    End If
    alignText = Trim(SafePrompt("页眉对齐：左 / 中 / 右", "中"))
    alignCode = NormalizeAlign(alignText)

    Dim fileName, fullPath, infoText
    fileName = "未命名文档"
    fullPath = ""
    On Error Resume Next
    If Len(CStr(doc.Name)) > 0 Then fileName = CStr(doc.Name)
    If Len(CStr(doc.FullName)) > 0 Then fullPath = CStr(doc.FullName)
    Err.Clear
    infoText = BuildHeaderInfoText(modeCode, fileName, fullPath)

    Dim sectionCount, nonemptyHeaders, nonemptyFooters, sampleHeader
    sectionCount = 0
    nonemptyHeaders = 0
    nonemptyFooters = 0
    sampleHeader = ""
    CollectHeaderFooterStats doc, sectionCount, nonemptyHeaders, nonemptyFooters, sampleHeader

    Dim wordSummary, formatSummary, previewText
    wordSummary = SafeHostText("GetWordContextInfo")
    formatSummary = SafeHostText("GetWordFormatIssueSummary")
    previewText = "Word 页眉文件信息预检（尚未写入）" & vbCrLf & _
        "文档=" & fileName & "；节数=" & CStr(sectionCount) & vbCrLf & _
        "已有非空页眉区域=" & CStr(nonemptyHeaders) & "；非空页脚区域=" & CStr(nonemptyFooters) & "（页脚默认不动）" & vbCrLf & _
        "将写入页眉=" & infoText & "；对齐=" & AlignLabel(alignCode) & vbCrLf & _
        "样例现有页眉=" & sampleHeader & vbCrLf & _
        "Host.GetWordContextInfo: " & wordSummary & vbCrLf & _
        "Host.GetWordFormatIssueSummary: " & formatSummary

    Dim planId, planPreview
    planId = SafeBeginWritePlan()
    SafeRecordWrite "word_header_file_info_preflight", "{""mode"":""" & EscapeJson(modeCode) & """,""sections"":" & CStr(sectionCount) & "}", "office.word.headerfooter.info"
    planPreview = SafePreviewWritePlan()

    Dim confirmText
    confirmText = Trim(SafePrompt(previewText & vbCrLf & vbCrLf & "以上尚未写入。确认写入页眉文件信息，请输入：写入", ""))
    If StrComp(confirmText, "写入", vbTextCompare) <> 0 Then
        SafeRollbackWritePlan planId
        HostWordHeaderFileInfoPreflight = FailureJson("E_CONFIRM_REQUIRED", "未输入“写入”，已取消且未修改文档")
        Exit Function
    End If

    Dim written
    written = ApplyHeaderInfo(doc, infoText, alignCode)
    SafeRollbackWritePlan planId

    Dim summary
    summary = "Word 页眉文件信息已写入；处理节=" & CStr(sectionCount) & "；写入页眉区域=" & CStr(written) & "；内容=" & infoText & "；页脚未改" & vbCrLf & _
        "Preview: " & planPreview
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostWordHeaderFileInfoPreflight = "{""ok"":true,""changed"":true,""sections"":" & CStr(sectionCount) & ",""writtenHeaders"":" & CStr(written) & ",""mode"":""" & EscapeJson(modeCode) & """,""message"":""" & EscapeJson(summary) & """}"
End Function

Function NormalizeHeaderMode(text)
    Dim t
    t = LCase(Trim(CStr(text)))
    If t = "文件名" Or t = "name" Or t = "filename" Then
        NormalizeHeaderMode = "name"
    ElseIf InStr(1, t, "路径", vbTextCompare) > 0 Or t = "path" Then
        NormalizeHeaderMode = "name_path"
    ElseIf InStr(1, t, "日期", vbTextCompare) > 0 Or t = "date" Then
        NormalizeHeaderMode = "name_date"
    Else
        NormalizeHeaderMode = ""
    End If
End Function

Function NormalizeAlign(text)
    Dim t
    t = LCase(Trim(CStr(text)))
    If t = "左" Or t = "left" Or t = "l" Then
        NormalizeAlign = wdAlignParagraphLeft
    ElseIf t = "右" Or t = "right" Or t = "r" Then
        NormalizeAlign = wdAlignParagraphRight
    Else
        NormalizeAlign = wdAlignParagraphCenter
    End If
End Function

Function AlignLabel(alignCode)
    If alignCode = wdAlignParagraphLeft Then
        AlignLabel = "左"
    ElseIf alignCode = wdAlignParagraphRight Then
        AlignLabel = "右"
    Else
        AlignLabel = "中"
    End If
End Function

Function BuildHeaderInfoText(modeCode, fileName, fullPath)
    If modeCode = "name_path" Then
        If Len(fullPath) > 0 Then
            BuildHeaderInfoText = fullPath
        Else
            BuildHeaderInfoText = fileName & "（未保存，无完整路径）"
        End If
    ElseIf modeCode = "name_date" Then
        BuildHeaderInfoText = fileName & "  " & FormatDateTime(Now, vbGeneralDate)
    Else
        BuildHeaderInfoText = fileName
    End If
End Function

Sub CollectHeaderFooterStats(doc, ByRef sectionCount, ByRef nonemptyHeaders, ByRef nonemptyFooters, ByRef sampleHeader)
    On Error Resume Next
    Dim section, part, textValue
    sectionCount = 0
    nonemptyHeaders = 0
    nonemptyFooters = 0
    sampleHeader = "（空）"
    For Each section In doc.Sections
        sectionCount = sectionCount + 1
        Set part = section.Headers(wdHeaderFooterPrimary)
        textValue = HeaderFooterText(part)
        If Len(NormalizeWs(textValue)) > 0 Then
            nonemptyHeaders = nonemptyHeaders + 1
            If sampleHeader = "（空）" Then sampleHeader = Host.LimitText(NormalizeWs(textValue), 40, "...")
        End If
        Set part = section.Headers(wdHeaderFooterFirstPage)
        If Len(NormalizeWs(HeaderFooterText(part))) > 0 Then nonemptyHeaders = nonemptyHeaders + 1
        Set part = section.Headers(wdHeaderFooterEvenPages)
        If Len(NormalizeWs(HeaderFooterText(part))) > 0 Then nonemptyHeaders = nonemptyHeaders + 1

        Set part = section.Footers(wdHeaderFooterPrimary)
        If Len(NormalizeWs(HeaderFooterText(part))) > 0 Then nonemptyFooters = nonemptyFooters + 1
        Set part = section.Footers(wdHeaderFooterFirstPage)
        If Len(NormalizeWs(HeaderFooterText(part))) > 0 Then nonemptyFooters = nonemptyFooters + 1
        Set part = section.Footers(wdHeaderFooterEvenPages)
        If Len(NormalizeWs(HeaderFooterText(part))) > 0 Then nonemptyFooters = nonemptyFooters + 1
    Next
End Sub

Function HeaderFooterText(part)
    On Error Resume Next
    HeaderFooterText = ""
    If TypeName(part) = "Nothing" Or TypeName(part) = "Empty" Then Exit Function
    HeaderFooterText = part.Range.Text
    If Err.Number <> 0 Then
        HeaderFooterText = ""
        Err.Clear
    End If
End Function

Function ApplyHeaderInfo(doc, infoText, alignCode)
    On Error Resume Next
    Dim section, written, part
    written = 0
    For Each section In doc.Sections
        Set part = section.Headers(wdHeaderFooterPrimary)
        If WriteHeaderPart(part, infoText, alignCode) Then written = written + 1
        ' only primary by default to avoid surprising first/even duplicates unless already linked differently
    Next
    ApplyHeaderInfo = written
End Function

Function WriteHeaderPart(part, infoText, alignCode)
    On Error Resume Next
    WriteHeaderPart = False
    If TypeName(part) = "Nothing" Or TypeName(part) = "Empty" Then Exit Function
    part.Range.Text = infoText & vbCr
    part.Range.ParagraphFormat.Alignment = alignCode
    If Err.Number = 0 Then WriteHeaderPart = True
    Err.Clear
End Function

Function NormalizeWs(value)
    Dim text
    text = CStr(value)
    text = Replace(text, Chr(7), "")
    text = Replace(text, vbCr, "")
    text = Replace(text, vbLf, "")
    text = Replace(text, ChrW(160), " ")
    NormalizeWs = Trim(text)
End Function

Function SafeHostText(methodName)
    On Error Resume Next
    If methodName = "GetWordContextInfo" Then
        SafeHostText = Host.GetWordContextInfo()
    ElseIf methodName = "GetWordFormatIssueSummary" Then
        SafeHostText = Host.GetWordFormatIssueSummary()
    ElseIf methodName = "GetWordReviewSummary" Then
        SafeHostText = Host.GetWordReviewSummary()
    Else
        SafeHostText = ""
    End If
    If Err.Number <> 0 Then SafeHostText = ""
    Err.Clear
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

Function SafePrompt(message, defaultValue)
    On Error Resume Next
    SafePrompt = Host.Prompt(message, defaultValue)
    If Err.Number <> 0 Then SafePrompt = defaultValue
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
