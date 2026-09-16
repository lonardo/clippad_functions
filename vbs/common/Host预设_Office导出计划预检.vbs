' 函数名: HostOfficeExportPlanPreflight
' 描述: 预检当前 Office 文档 PDF 导出计划，一键确认后按 avoid 策略真实导出；可改预览
' 适用应用: Word|Excel|PowerPoint
' 搜索范围: 全文
' 搜索对象: 无
' 作用范围: 全部
' 需要参数: 否
'
' 实现约束（给后续维护/模型使用）：
' - 先汇总 Host 导出预检（路径/canRun/canExport/体量），默认动作为“预览”；只有明确输入“导出”才会写入 PDF。
' - 真导出走各应用 ExportAsFixedFormat，路径来自 Host.GetExportPlan(..., "avoid")，不覆盖已有同名 PDF。
' - 写计划只记录/展示 Host 元数据；RollbackWritePlan 不能撤销已生成的 PDF 文件。

Option Explicit

Const wdExportFormatPDF = 17
Const xlTypePDF = 0
Const ppSaveAsPDF = 32

Function Main()
    On Error Resume Next

    Dim appObj
    Set appObj = Nothing
    Err.Clear
    Set appObj = Host.GetApplication()
    Err.Clear

    Main = HostOfficeExportPlanPreflight(appObj)
End Function

Function HostOfficeExportPlanPreflight(appObj)
    On Error Resume Next

    Dim defaultName, overwritePolicy
    defaultName = "office_export.pdf"
    overwritePolicy = "avoid"

    Dim hostInfo, docInfo, contextInfo, preflightJson, exportJson
    hostInfo = SafeHostText("GetHostInfo")
    docInfo = SafeHostText("GetDocumentInfo")
    contextInfo = SafeHostText("GetContextInfo")
    preflightJson = SafeHostText("GetRunPreflightPlan")
    exportJson = SafeExportPlan(defaultName, overwritePolicy)

    Dim appType, hasPath, canRun, canExport, exportPath, sizeHint, docName, riskNotes
    appType = FirstNonEmpty(ExtractJsonString(docInfo, "appType"), ExtractJsonString(contextInfo, "appType"), DetectAppType(appObj), "未知")
    hasPath = ExtractJsonBoolean(docInfo, "hasPath")
    canRun = ExtractJsonBoolean(preflightJson, "canRun")
    If Not JsonHasKey(preflightJson, "canRun") Then canRun = True
    canExport = ExtractJsonBoolean(exportJson, "canExport")
    If Not JsonHasKey(exportJson, "canExport") Then canExport = True
    exportPath = ExtractJsonString(exportJson, "path")
    If Len(exportPath) = 0 Then exportPath = "(未解析路径)"
    sizeHint = BuildSizeHint(appObj, appType)
    docName = DetectDocName(appObj, appType)
    riskNotes = BuildRiskNotes(hasPath, canRun, canExport, overwritePolicy)

    Dim previewText, planId, planPreview, confirmRaw, doExport
    previewText = "Office 导出计划预检（尚未导出）" & vbCrLf & _
        "应用类型: " & appType & vbCrLf & _
        "文档名称: " & docName & vbCrLf & _
        "是否已保存路径: " & YesNoLabel(hasPath) & vbCrLf & _
        "体量提示: " & sizeHint & vbCrLf & _
        "建议导出文件: " & defaultName & vbCrLf & _
        "导出路径计划: " & exportPath & vbCrLf & _
        "覆盖策略: " & overwritePolicy & "（避免覆盖已有文件）" & vbCrLf & _
        "运行预检 canRun: " & YesNoLabel(canRun) & vbCrLf & _
        "导出预检 canExport: " & YesNoLabel(canExport) & vbCrLf & _
        "风险与建议: " & riskNotes & vbCrLf & _
        "默认动作=预览（源文档和文件均不改）" & vbCrLf & _
        "Host.GetHostInfo: " & hostInfo & vbCrLf & _
        "Host.GetDocumentInfo: " & docInfo & vbCrLf & _
        "Host.GetContextInfo: " & contextInfo & vbCrLf & _
        "Host.GetRunPreflightPlan: " & preflightJson & vbCrLf & _
        "Host.GetExportPlan: " & exportJson & vbCrLf & vbCrLf & _
        "确认导出：请输入 导出。" & vbCrLf & _
        "仅预览：直接确认或输入 预览" & vbCrLf & _
        "覆盖写法：导出 | 预览"

    planId = SafeBeginWritePlan()
    SafeRecordWrite "office_export_plan", "{""format"":""pdf"",""path"":""" & EscapeJson(exportPath) & """,""policy"":""" & EscapeJson(overwritePolicy) & """}", "office.export.pdf"
    planPreview = SafePreviewWritePlan()
    CloseWritePlan planId

    confirmRaw = SafePrompt(previewText, "预览")
    doExport = False
    If Not ParseExportConfirm(confirmRaw, doExport) Then
        HostOfficeExportPlanPreflight = FailureJson("E_CONFIRM_REQUIRED", "未确认导出/预览，已取消")
        Exit Function
    End If

    If Not doExport Then
        Host.WriteClipboard previewText
        SafeWriteLog previewText
        HostOfficeExportPlanPreflight = "{""ok"":true,""readonly"":true,""appType"":""" & EscapeJson(appType) & """,""hasPath"":" & LCase(CStr(hasPath)) & _
            ",""canRun"":" & LCase(CStr(canRun)) & ",""canExport"":" & LCase(CStr(canExport)) & _
            ",""exportPath"":""" & EscapeJson(exportPath) & """,""message"":""" & EscapeJson(previewText) & """}"
        Exit Function
    End If

    If TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        HostOfficeExportPlanPreflight = FailureJson("E_NO_APP", "未取得 Office 应用，无法导出")
        Exit Function
    End If
    If Not canExport Or Len(Trim(exportPath)) = 0 Or Left(exportPath, 1) = "(" Then
        HostOfficeExportPlanPreflight = FailureJson("E_EXPORT_PATH", "导出路径不可用：" & exportPath)
        Exit Function
    End If
    If Not hasPath Then
        HostOfficeExportPlanPreflight = FailureJson("E_DOC_NOT_SAVED", "当前文档尚未保存到磁盘，请先保存后再导出")
        Exit Function
    End If

    planId = SafeBeginWritePlan()
    SafeRecordWrite "office_export_pdf", "{""path"":""" & EscapeJson(exportPath) & """,""appType"":""" & EscapeJson(appType) & """}", "office.export.pdf"
    planPreview = SafePreviewWritePlan()

    Dim exportError
    exportError = ""
    If Not ExportPdfByApp(appObj, appType, exportPath, exportError) Then
        CloseWritePlan planId
        HostOfficeExportPlanPreflight = FailureJson("E_EXPORT_FAILED", "PDF 导出失败：" & exportError)
        Exit Function
    End If
    CloseWritePlan planId

    Dim summary
    summary = "Office PDF 导出完成；应用=" & appType & "，文档=" & docName & "，路径=" & exportPath & _
        "，源文档内容未改" & vbCrLf & "Preview: " & planPreview
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostOfficeExportPlanPreflight = "{""ok"":true,""exported"":true,""sourceUnchanged"":true,""appType"":""" & EscapeJson(appType) & _
        """,""exportPath"":""" & EscapeJson(exportPath) & """,""writePlanPreview"":""" & EscapeJson(planPreview) & _
        """,""message"":""" & EscapeJson(summary) & """}"
End Function

Function ParseExportConfirm(rawText, ByRef doExport)
    Dim text, parts, head
    doExport = False
    text = Trim(CStr(rawText))
    If Len(text) = 0 Then
        ParseExportConfirm = False
        Exit Function
    End If
    If InStr(1, text, "|", vbBinaryCompare) > 0 Then
        parts = Split(text, "|")
        head = Trim(CStr(parts(0)))
    Else
        head = text
    End If

    If StrComp(head, "导出", vbTextCompare) = 0 Or StrComp(head, "export", vbTextCompare) = 0 Then
        doExport = True
        ParseExportConfirm = True
    ElseIf StrComp(head, "预览", vbTextCompare) = 0 Or StrComp(head, "preview", vbTextCompare) = 0 Then
        doExport = False
        ParseExportConfirm = True
    Else
        ParseExportConfirm = False
    End If
End Function

Function ExportPdfByApp(appObj, appType, exportPath, ByRef exportError)
    On Error Resume Next
    Dim docObj
    exportError = ""
    ExportPdfByApp = False

    If StrComp(appType, "Word", vbTextCompare) = 0 Then
        Set docObj = appObj.ActiveDocument
        If Err.Number <> 0 Or TypeName(docObj) = "Empty" Or TypeName(docObj) = "Nothing" Then
            exportError = "无活动 Word 文档"
            Err.Clear
            Exit Function
        End If
        Err.Clear
        docObj.ExportAsFixedFormat exportPath, wdExportFormatPDF, False, 0, 0, 1, 1, 0, True, True, 1, True, True, False
    ElseIf StrComp(appType, "Excel", vbTextCompare) = 0 Then
        Set docObj = appObj.ActiveWorkbook
        If Err.Number <> 0 Or TypeName(docObj) = "Empty" Or TypeName(docObj) = "Nothing" Then
            exportError = "无活动 Excel 工作簿"
            Err.Clear
            Exit Function
        End If
        Err.Clear
        docObj.ExportAsFixedFormat xlTypePDF, exportPath, 0, 0, False, True, False
    ElseIf StrComp(appType, "PowerPoint", vbTextCompare) = 0 Then
        Set docObj = appObj.ActivePresentation
        If Err.Number <> 0 Or TypeName(docObj) = "Empty" Or TypeName(docObj) = "Nothing" Then
            exportError = "无活动 PowerPoint 演示文稿"
            Err.Clear
            Exit Function
        End If
        Err.Clear
        docObj.ExportAsFixedFormat exportPath, ppSaveAsPDF, 2, 0, 0, 0, 0, 0, False, True, 1, False, True, False
    Else
        exportError = "暂不支持的应用类型：" & appType
        Exit Function
    End If

    If Err.Number <> 0 Then
        exportError = Err.Description
        Err.Clear
        ExportPdfByApp = False
    Else
        ExportPdfByApp = True
    End If
End Function

Function SafeExportPlan(defaultName, overwritePolicy)
    On Error Resume Next
    SafeExportPlan = Host.GetExportPlan("pdf", defaultName, overwritePolicy)
    If Err.Number <> 0 Then SafeExportPlan = ""
    Err.Clear
End Function

Function SafeHostText(methodName)
    On Error Resume Next
    Err.Clear
    Select Case methodName
        Case "GetHostInfo"
            SafeHostText = Host.GetHostInfo()
        Case "GetDocumentInfo"
            SafeHostText = Host.GetDocumentInfo()
        Case "GetContextInfo"
            SafeHostText = Host.GetContextInfo()
        Case "GetRunPreflightPlan"
            SafeHostText = Host.GetRunPreflightPlan("", True, False, False)
        Case Else
            SafeHostText = ""
    End Select
    If Err.Number <> 0 Then
        SafeHostText = ""
        Err.Clear
    End If
End Function

Function DetectAppType(appObj)
    On Error Resume Next
    Dim n
    DetectAppType = "未知"
    If TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then Exit Function
    n = ""
    n = CStr(appObj.Name)
    If Err.Number <> 0 Then
        Err.Clear
        n = TypeName(appObj)
    End If
    n = LCase(n)
    If InStr(1, n, "word", vbTextCompare) > 0 Then
        DetectAppType = "Word"
    ElseIf InStr(1, n, "excel", vbTextCompare) > 0 Then
        DetectAppType = "Excel"
    ElseIf InStr(1, n, "power", vbTextCompare) > 0 Then
        DetectAppType = "PowerPoint"
    Else
        DetectAppType = n
    End If
End Function

Function DetectDocName(appObj, appType)
    On Error Resume Next
    Dim nameText
    nameText = ""
    If TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        DetectDocName = "(无活动文档)"
        Exit Function
    End If

    If StrComp(appType, "Word", vbTextCompare) = 0 Then
        nameText = CStr(appObj.ActiveDocument.Name)
    ElseIf StrComp(appType, "Excel", vbTextCompare) = 0 Then
        nameText = CStr(appObj.ActiveWorkbook.Name)
    ElseIf StrComp(appType, "PowerPoint", vbTextCompare) = 0 Then
        nameText = CStr(appObj.ActivePresentation.Name)
    Else
        nameText = ""
        nameText = CStr(appObj.ActiveDocument.Name)
        If Err.Number <> 0 Or Len(Trim(nameText)) = 0 Then
            Err.Clear
            nameText = CStr(appObj.ActiveWorkbook.Name)
        End If
        If Err.Number <> 0 Or Len(Trim(nameText)) = 0 Then
            Err.Clear
            nameText = CStr(appObj.ActivePresentation.Name)
        End If
    End If

    If Err.Number <> 0 Or Len(Trim(nameText)) = 0 Then
        Err.Clear
        DetectDocName = "(未命名/不可读)"
    Else
        DetectDocName = nameText
    End If
End Function

Function BuildSizeHint(appObj, appType)
    On Error Resume Next
    Dim pages, slides, usedRange, rows, cols
    BuildSizeHint = "未知"
    If TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then Exit Function

    If StrComp(appType, "Word", vbTextCompare) = 0 Then
        pages = appObj.ActiveDocument.ComputeStatistics(2)
        If Err.Number <> 0 Then
            Err.Clear
            pages = appObj.ActiveDocument.Range.Information(4)
        End If
        If Err.Number = 0 Then
            BuildSizeHint = "约 " & CStr(pages) & " 页"
        End If
        Err.Clear
    ElseIf StrComp(appType, "Excel", vbTextCompare) = 0 Then
        Set usedRange = appObj.ActiveSheet.UsedRange
        If Err.Number = 0 Then
            rows = usedRange.Rows.Count
            cols = usedRange.Columns.Count
            BuildSizeHint = "活动表约 " & CStr(rows) & " 行 x " & CStr(cols) & " 列"
        End If
        Err.Clear
    ElseIf StrComp(appType, "PowerPoint", vbTextCompare) = 0 Then
        slides = appObj.ActivePresentation.Slides.Count
        If Err.Number = 0 Then
            BuildSizeHint = "约 " & CStr(slides) & " 页幻灯片"
        End If
        Err.Clear
    End If

    If BuildSizeHint = "未知" Then
        BuildSizeHint = "未能估算体量（请结合导出预检）"
    End If
End Function

Function BuildRiskNotes(hasPath, canRun, canExport, overwritePolicy)
    Dim notes
    notes = ""
    If Not hasPath Then notes = notes & "当前文档可能未保存到磁盘，正式导出前建议先保存。"
    If Not canRun Then
        If Len(notes) > 0 Then notes = notes & " "
        notes = notes & "运行预检 canRun=否，检查是否有活动文档/权限。"
    End If
    If Not canExport Then
        If Len(notes) > 0 Then notes = notes & " "
        notes = notes & "导出预检 canExport=否，暂不建议导出。"
    End If
    If StrComp(CStr(overwritePolicy), "avoid", vbTextCompare) = 0 Then
        If Len(notes) > 0 Then notes = notes & " "
        notes = notes & "默认避免覆盖同名 PDF。"
    End If
    If Len(notes) = 0 Then notes = "预检未见明显阻断项；确认后将生成外部 PDF。"
    BuildRiskNotes = notes
End Function

Function FirstNonEmpty(a, b, c, d)
    If Len(Trim(CStr(a))) > 0 Then
        FirstNonEmpty = CStr(a)
    ElseIf Len(Trim(CStr(b))) > 0 Then
        FirstNonEmpty = CStr(b)
    ElseIf Len(Trim(CStr(c))) > 0 Then
        FirstNonEmpty = CStr(c)
    Else
        FirstNonEmpty = CStr(d)
    End If
End Function

Function JsonHasKey(jsonText, keyName)
    Dim marker
    marker = """" & CStr(keyName) & """"
    JsonHasKey = (InStr(1, CStr(jsonText), marker, vbTextCompare) > 0)
End Function

Function ExtractJsonBoolean(jsonText, keyName)
    Dim marker, keyPos, colonPos, valueText
    ExtractJsonBoolean = False
    marker = """" & CStr(keyName) & """"
    keyPos = InStr(1, CStr(jsonText), marker, vbTextCompare)
    If keyPos <= 0 Then Exit Function
    colonPos = InStr(keyPos + Len(marker), CStr(jsonText), ":")
    If colonPos <= 0 Then Exit Function
    valueText = LCase(LTrim(Mid(CStr(jsonText), colonPos + 1)))
    ExtractJsonBoolean = (Left(valueText, 4) = "true")
End Function

Function ExtractJsonString(jsonText, keyName)
    Dim marker, keyPos, colonPos, quotePos, i, ch, nextCh, result
    ExtractJsonString = ""
    marker = """" & CStr(keyName) & """"
    keyPos = InStr(1, CStr(jsonText), marker, vbTextCompare)
    If keyPos <= 0 Then Exit Function
    colonPos = InStr(keyPos + Len(marker), CStr(jsonText), ":")
    If colonPos <= 0 Then Exit Function
    quotePos = InStr(colonPos + 1, CStr(jsonText), """")
    If quotePos <= 0 Then Exit Function
    result = ""
    i = quotePos + 1
    Do While i <= Len(jsonText)
        ch = Mid(jsonText, i, 1)
        If ch = """" Then Exit Do
        If ch = "\" And i < Len(jsonText) Then
            nextCh = Mid(jsonText, i + 1, 1)
            result = result & nextCh
            i = i + 2
        Else
            result = result & ch
            i = i + 1
        End If
    Loop
    ExtractJsonString = result
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

Sub CloseWritePlan(planId)
    On Error Resume Next
    Host.RollbackWritePlan planId
    Err.Clear
End Sub

Sub SafeWriteLog(message)
    On Error Resume Next
    Host.WriteLog message
    Err.Clear
End Sub

Function YesNoLabel(flag)
    If flag Then
        YesNoLabel = "是"
    Else
        YesNoLabel = "否"
    End If
End Function
