' 函数名: HostExcelFormulaExplainAuditPack
' 描述: 公式审计增强场景包：扫描错误/外链/易波动公式，输出问题明细（含人话说明）与处理摘要，不改源表
' 适用应用: Excel
' 搜索范围: 当前范围
' 搜索对象: 无
' 作用范围: 当前范围
' 输出类型: 表

Option Explicit

Function Main()
    On Error Resume Next
    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = FailureJson("E_NO_EXCEL_APP", "未取得 Excel 应用，请在 Excel 中运行该预设")
        Exit Function
    End If
    Main = HostExcelFormulaExplainAuditPack(appObj)
End Function

Function HostExcelFormulaExplainAuditPack(appObj)
    On Error Resume Next
    Dim sourceRange, sourceSheet, sourceName, outputMode
    Set sourceRange = appObj.Selection.CurrentRegion
    If Err.Number <> 0 Or TypeName(sourceRange) = "Empty" Or TypeName(sourceRange) = "Nothing" Then
        Err.Clear
        HostExcelFormulaExplainAuditPack = FailureJson("E_NO_RANGE", "请先选中要审计的数据区域")
        Exit Function
    End If
    If sourceRange.Rows.Count < 1 Or sourceRange.Columns.Count < 1 Then
        HostExcelFormulaExplainAuditPack = FailureJson("E_RANGE_TOO_SMALL", "当前区域为空，无法审计公式")
        Exit Function
    End If
    Set sourceSheet = sourceRange.Worksheet
    sourceName = sourceSheet.Name
    outputMode = NormalizeOutputMode(SafePrompt("输出模式：NewSheet 或 NewWorkbook（默认 NewSheet）", "NewSheet"))

    Dim hostSummary, typeStats, preflightJson
    hostSummary = ""
    typeStats = ""
    preflightJson = ""
    On Error Resume Next
    hostSummary = CStr(Host.GetExcelFormulaIssueSummary())
    If Err.Number <> 0 Then hostSummary = "": Err.Clear
    typeStats = CStr(Host.GetTypeStats())
    If Err.Number <> 0 Then typeStats = "": Err.Clear
    preflightJson = CStr(Host.GetRunPreflightPlan("Excel", True, True, False))
    If Err.Number <> 0 Then preflightJson = "": Err.Clear

    Dim formulaCount, errorCount, externalCount, volatileCount, issueCount
    Dim addrArr(), kindArr(), explainArr(), formulaArr()
    issueCount = 0
    ReDim addrArr(0)
    ReDim kindArr(0)
    ReDim explainArr(0)
    ReDim formulaArr(0)
    ScanFormulaIssues sourceRange, formulaCount, errorCount, externalCount, volatileCount, addrArr, kindArr, explainArr, formulaArr, issueCount

    Dim previewText, planId, planPreview
    previewText = "Excel 公式审计增强场景包预览（尚未写入）" & vbCrLf & _
        "命令ID：excel.formula_explain_audit" & vbCrLf & _
        "源表=" & sourceName & "；区域=" & sourceRange.Address & vbCrLf & _
        "输出模式=" & outputMode & vbCrLf & _
        "公式单元格=" & CStr(formulaCount) & "；问题数=" & CStr(issueCount) & vbCrLf & _
        "错误=" & CStr(errorCount) & "；外部链接=" & CStr(externalCount) & "；易波动=" & CStr(volatileCount) & vbCrLf & _
        "将生成：公式问题明细 + 处理摘要；sourceUnchanged=true；不改源表格式"
    If Len(hostSummary) > 0 Then previewText = previewText & vbCrLf & "Host摘要=" & hostSummary

    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_formula_explain_audit", "{""source"":""" & EscapeJson(sourceRange.Address) & """,""issues"":" & CStr(issueCount) & "}", "office.excel.formulaExplainAudit"
    planPreview = SafePreviewWritePlan()
    SafeRollbackWritePlan planId

    If Not SafeConfirmStep(previewText & vbCrLf & vbCrLf & "确认生成公式审计报告？", "excel.formula_explain_audit") Then
        HostExcelFormulaExplainAuditPack = FailureJson("E_CONFIRM_REQUIRED", "用户取消或未确认，未修改工作簿")
        Exit Function
    End If

    Dim createdSheets(), createdCount, detailSheet, summarySheet
    createdCount = 0
    ReDim createdSheets(4)

    Set detailSheet = CreateOutputSheet(appObj, sourceName, "公式问题明细", outputMode)
    If detailSheet Is Nothing Then
        HostExcelFormulaExplainAuditPack = FailureJson("E_OUTPUT_SHEET", "无法创建公式问题明细工作表")
        Exit Function
    End If
    createdCount = createdCount + 1: Set createdSheets(createdCount) = detailSheet

    If Not WriteFormulaIssueSheet(detailSheet, addrArr, kindArr, explainArr, formulaArr, issueCount) Then
        RollbackCreatedSheets appObj, createdSheets, createdCount
        HostExcelFormulaExplainAuditPack = FailureJson("E_FORMULA_WRITE", "公式问题明细写入失败，已删除未完成输出")
        Exit Function
    End If

    Set summarySheet = CreateOutputSheet(appObj, sourceName, "处理摘要", outputMode)
    If summarySheet Is Nothing Then
        RollbackCreatedSheets appObj, createdSheets, createdCount
        HostExcelFormulaExplainAuditPack = FailureJson("E_OUTPUT_SHEET", "无法创建处理摘要工作表")
        Exit Function
    End If
    createdCount = createdCount + 1: Set createdSheets(createdCount) = summarySheet

    Dim summaryLines, impactJson, summary, aigenHint
    aigenHint = "AIGEN验收骨架：请解释前20个问题单元格的业务影响，并给出是否允许交付的判断；不要直接改源表。"
    summaryLines = "场景名=公式审计增强场景包" & vbCrLf & _
        "命令ID=excel.formula_explain_audit" & vbCrLf & _
        "源表名称=" & sourceName & vbCrLf & _
        "源区域=" & sourceRange.Address & vbCrLf & _
        "输出模式=" & outputMode & vbCrLf & _
        "输出位置=" & detailSheet.Name & "," & summarySheet.Name & vbCrLf & _
        "处理前行数=" & CStr(sourceRange.Rows.Count) & vbCrLf & _
        "处理后行数=" & CStr(sourceRange.Rows.Count) & vbCrLf & _
        "参数=scan=errors,external,volatile" & vbCrLf & _
        "公式数=" & CStr(formulaCount) & vbCrLf & _
        "异常/问题数=" & CStr(issueCount) & vbCrLf & _
        "错误=" & CStr(errorCount) & "；外部链接=" & CStr(externalCount) & "；易波动=" & CStr(volatileCount) & vbCrLf & _
        "Host.GetExcelFormulaIssueSummary=" & hostSummary & vbCrLf & _
        "Host.GetTypeStats=" & typeStats & vbCrLf & _
        "AIGEN提示=" & aigenHint & vbCrLf & _
        "执行时间=" & Now & vbCrLf & _
        "结果=成功"
    impactJson = SafeHostText("ExcelWriteImpactSummary", summaryLines, summarySheet.Name)
    If Not ExtractJsonBoolean(impactJson, "ok") Then WriteSummaryFallback summarySheet, summaryLines

    summary = "Excel 公式审计增强完成：明细=" & detailSheet.Name & "；问题=" & CStr(issueCount) & "；公式=" & CStr(formulaCount) & "；sourceUnchanged=true"
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostExcelFormulaExplainAuditPack = "{""ok"":true,""sourceUnchanged"":true,""commandId"":""excel.formula_explain_audit"",""detailSheet"":""" & EscapeJson(detailSheet.Name) & """,""summarySheet"":""" & EscapeJson(summarySheet.Name) & """,""formulaCount"":" & CStr(formulaCount) & ",""issueCount"":" & CStr(issueCount) & ",""errorCount"":" & CStr(errorCount) & ",""externalLinkCount"":" & CStr(externalCount) & ",""volatileCount"":" & CStr(volatileCount) & ",""writePlanPreview"":""" & EscapeJson(planPreview) & """,""message"":""" & EscapeJson(summary) & """}"
End Function

Sub ScanFormulaIssues(sourceRange, ByRef formulaCount, ByRef errorCount, ByRef externalCount, ByRef volatileCount, ByRef addrArr, ByRef kindArr, ByRef explainArr, ByRef formulaArr, ByRef issueCount)
    On Error Resume Next
    Dim cell, formulaText, kinds, explain, sampleCap
    formulaCount = 0
    errorCount = 0
    externalCount = 0
    volatileCount = 0
    issueCount = 0
    sampleCap = 5000
    ReDim addrArr(sampleCap)
    ReDim kindArr(sampleCap)
    ReDim explainArr(sampleCap)
    ReDim formulaArr(sampleCap)

    For Each cell In sourceRange.Cells
        formulaText = ""
        If cell.HasFormula Then formulaText = CStr(cell.Formula)
        If Len(formulaText) > 0 Then
            formulaCount = formulaCount + 1
            kinds = ""
            explain = ""
            If IsError(cell.Value) Then
                errorCount = errorCount + 1
                kinds = AppendToken(kinds, "公式错误")
                explain = AppendToken(explain, ExplainFormulaError(cell, formulaText))
            End If
            If FormulaHasExternalLink(formulaText) Then
                externalCount = externalCount + 1
                kinds = AppendToken(kinds, "外部链接")
                explain = AppendToken(explain, "公式引用了其他工作簿，换机或断网后可能失效")
            End If
            If FormulaHasVolatileFunction(formulaText) Then
                volatileCount = volatileCount + 1
                kinds = AppendToken(kinds, "易波动函数")
                explain = AppendToken(explain, "含 NOW/TODAY/RAND/OFFSET/INDIRECT 等，重算结果可能变化")
            End If
            If Len(kinds) > 0 Then
                If issueCount < sampleCap Then
                    addrArr(issueCount) = cell.Address(False, False)
                    kindArr(issueCount) = kinds
                    explainArr(issueCount) = explain
                    formulaArr(issueCount) = formulaText
                    issueCount = issueCount + 1
                End If
            End If
        End If
        If Err.Number <> 0 Then Err.Clear
    Next
End Sub

Function WriteFormulaIssueSheet(sheetObj, ByRef addrArr, ByRef kindArr, ByRef explainArr, ByRef formulaArr, issueCount)
    On Error Resume Next
    Dim i
    WriteFormulaIssueSheet = False
    sheetObj.Cells(1, 1).Value = "单元格"
    sheetObj.Cells(1, 2).Value = "问题类型"
    sheetObj.Cells(1, 3).Value = "人话说明"
    sheetObj.Cells(1, 4).Value = "公式"
    sheetObj.Cells(1, 5).Value = "建议"
    If issueCount = 0 Then
        sheetObj.Cells(2, 1).Value = "(无)"
        sheetObj.Cells(2, 2).Value = "无问题"
        sheetObj.Cells(2, 3).Value = "当前区域未发现公式错误、外部链接或易波动函数"
        sheetObj.Cells(2, 4).Value = ""
        sheetObj.Cells(2, 5).Value = "可继续交付前人工抽查关键汇总"
    Else
        For i = 0 To issueCount - 1
            sheetObj.Cells(i + 2, 1).Value = addrArr(i)
            sheetObj.Cells(i + 2, 2).Value = kindArr(i)
            sheetObj.Cells(i + 2, 3).Value = explainArr(i)
            sheetObj.Cells(i + 2, 4).Value = "'" & formulaArr(i)
            sheetObj.Cells(i + 2, 5).Value = SuggestForFormulaIssue(kindArr(i))
        Next
    End If
    sheetObj.Rows(1).Font.Bold = True
    sheetObj.Columns.AutoFit
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    WriteFormulaIssueSheet = True
End Function

Function AppendToken(existing, token)
    If Len(existing) = 0 Then
        AppendToken = token
    Else
        AppendToken = existing & "；" & token
    End If
End Function

Function ExplainFormulaError(cell, formulaText)
    Dim v
    On Error Resume Next
    v = CStr(cell.Text)
    If Len(v) = 0 Then v = "#ERROR"
    ExplainFormulaError = "单元格显示 " & v & "，公式未能算出有效结果"
End Function

Function SuggestForFormulaIssue(kindText)
    Dim t, buf
    t = CStr(kindText)
    buf = ""
    If InStr(1, t, "公式错误", vbTextCompare) > 0 Then buf = AppendToken(buf, "先修错误引用/除零/名称后再交付")
    If InStr(1, t, "外部链接", vbTextCompare) > 0 Then buf = AppendToken(buf, "改为当前工作簿引用或提供外部文件")
    If InStr(1, t, "易波动", vbTextCompare) > 0 Then buf = AppendToken(buf, "交付前转数值快照，避免打开即变")
    If Len(buf) = 0 Then buf = "人工复核"
    SuggestForFormulaIssue = buf
End Function

Function FormulaHasExternalLink(formulaText)
    Dim t
    t = CStr(formulaText)
    FormulaHasExternalLink = (InStr(t, "[") > 0 And InStr(t, "]") > InStr(t, "[")) Or (InStr(1, t, "http", vbTextCompare) > 0)
End Function

Function FormulaHasVolatileFunction(formulaText)
    Dim t
    t = UCase(CStr(formulaText))
    FormulaHasVolatileFunction = (InStr(t, "NOW(") > 0 Or InStr(t, "TODAY(") > 0 Or InStr(t, "RAND(") > 0 Or InStr(t, "RANDBETWEEN(") > 0 Or InStr(t, "OFFSET(") > 0 Or InStr(t, "INDIRECT(") > 0 Or InStr(t, "INFO(") > 0 Or InStr(t, "CELL(") > 0)
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

Function ExtractJsonNumber(jsonText, keyName)
    Dim marker, keyPos, colonPos, valueText, ch, i, buf
    ExtractJsonNumber = 0
    marker = """" & CStr(keyName) & """"
    keyPos = InStr(1, CStr(jsonText), marker, vbTextCompare)
    If keyPos <= 0 Then Exit Function
    colonPos = InStr(keyPos + Len(marker), CStr(jsonText), ":")
    If colonPos <= 0 Then Exit Function
    valueText = LTrim(Mid(CStr(jsonText), colonPos + 1))
    buf = ""
    For i = 1 To Len(valueText)
        ch = Mid(valueText, i, 1)
        If (ch >= "0" And ch <= "9") Or ch = "-" Or ch = "." Then
            buf = buf & ch
        ElseIf Len(buf) > 0 Then
            Exit For
        End If
    Next
    If Len(buf) > 0 Then ExtractJsonNumber = CDbl(buf)
End Function

Function ParseJsonIntArray(jsonText, keyName)
    Dim marker, keyPos, bracketStart, bracketEnd, body, parts, i, n, values()
    ParseJsonIntArray = Array()
    marker = """" & CStr(keyName) & """"
    keyPos = InStr(1, CStr(jsonText), marker, vbTextCompare)
    If keyPos <= 0 Then Exit Function
    bracketStart = InStr(keyPos, CStr(jsonText), "[")
    If bracketStart <= 0 Then Exit Function
    bracketEnd = InStr(bracketStart + 1, CStr(jsonText), "]")
    If bracketEnd <= bracketStart Then Exit Function
    body = Trim(Mid(CStr(jsonText), bracketStart + 1, bracketEnd - bracketStart - 1))
    If Len(body) = 0 Then Exit Function
    parts = Split(body, ",")
    n = -1
    ReDim values(UBound(parts))
    For i = 0 To UBound(parts)
        If Len(Trim(parts(i))) > 0 Then
            n = n + 1
            values(n) = CLng(Trim(parts(i)))
        End If
    Next
    If n >= 0 Then
        ReDim Preserve values(n)
        ParseJsonIntArray = values
    End If
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
    If Err.Number <> 0 Then
        SafePrompt = CStr(defaultValue)
        Err.Clear
    End If
End Function

Function SafeConfirmStep(message, stepInfo)
    On Error Resume Next
    Dim confirmed, answer
    confirmed = Host.ConfirmStep(message, stepInfo)
    If Err.Number = 0 Then
        SafeConfirmStep = CBool(confirmed)
        Exit Function
    End If
    Err.Clear
    answer = Trim(SafePrompt(message & vbCrLf & vbCrLf & "确认继续请输入：确认", ""))
    SafeConfirmStep = (StrComp(answer, "确认", vbTextCompare) = 0)
End Function

Function SafeBeginWritePlan()
    On Error Resume Next
    SafeBeginWritePlan = Host.BeginWritePlan("")
    If Err.Number <> 0 Then SafeBeginWritePlan = ""
    Err.Clear
End Function

Sub SafeRecordWrite(actionId, paramsJson, capabilityId)
    On Error Resume Next
    Host.RecordWrite actionId, paramsJson, "office", "excel.sheet.delete", "{}", capabilityId
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

Function SafeHostText(methodName, arg1, arg2, arg3, arg4)
    On Error Resume Next
    SafeHostText = ""
    If methodName = "ExcelGetHeaderMap" Then
        SafeHostText = Host.ExcelGetHeaderMap(arg1)
    ElseIf methodName = "ExcelCreateOutputSheet" Then
        SafeHostText = Host.ExcelCreateOutputSheet(arg1, arg2, arg3, arg4)
    ElseIf methodName = "ExcelWriteImpactSummary" Then
        SafeHostText = Host.ExcelWriteImpactSummary(arg1, arg2)
    ElseIf methodName = "ExcelPreviewRowSamples" Then
        SafeHostText = Host.ExcelPreviewRowSamples(arg1, arg2)
    ElseIf methodName = "ExcelDedupePlan" Then
        SafeHostText = Host.ExcelDedupePlan(arg1, arg2)
    ElseIf methodName = "ExcelValidatePlan" Then
        SafeHostText = Host.ExcelValidatePlan(arg1)
    End If
    If Err.Number <> 0 Then SafeHostText = ""
    Err.Clear
End Function

Function NormalizeCellText(value)
    Dim text
    text = CStr(value)
    text = Replace(text, ChrW(160), " ")
    text = Replace(text, ChrW(12288), " ")
    NormalizeCellText = Trim(text)
End Function

Function NormalizeKey(value)
    NormalizeKey = LCase(NormalizeCellText(value))
End Function

Function FindHeaderColumn(sourceRange, fieldName)
    Dim colIndex
    FindHeaderColumn = 0
    For colIndex = 1 To sourceRange.Columns.Count
        If StrComp(NormalizeCellText(sourceRange.Cells(1, colIndex).Text), NormalizeCellText(fieldName), vbTextCompare) = 0 Then
            FindHeaderColumn = colIndex
            Exit Function
        End If
    Next
End Function

Function SplitCsvFields(text)
    Dim cleaned, parts, i, n, values()
    cleaned = Replace(Replace(Trim(CStr(text)), "，", ","), "；", ",")
    If Len(cleaned) = 0 Then
        SplitCsvFields = Array()
        Exit Function
    End If
    parts = Split(cleaned, ",")
    n = -1
    ReDim values(UBound(parts))
    For i = 0 To UBound(parts)
        If Len(Trim(parts(i))) > 0 Then
            n = n + 1
            values(n) = Trim(parts(i))
        End If
    Next
    If n < 0 Then
        SplitCsvFields = Array()
    Else
        ReDim Preserve values(n)
        SplitCsvFields = values
    End If
End Function

Function ResolveHeaderColumns(sourceRange, fieldNames)
    Dim i, colIndex, n, cols()
    ResolveHeaderColumns = Array()
    If Not IsArray(fieldNames) Then Exit Function
    On Error Resume Next
    If UBound(fieldNames) < 0 Then Exit Function
    Err.Clear
    n = -1
    ReDim cols(UBound(fieldNames))
    For i = 0 To UBound(fieldNames)
        colIndex = FindHeaderColumn(sourceRange, fieldNames(i))
        If colIndex <= 0 Then
            ResolveHeaderColumns = Array()
            Exit Function
        End If
        n = n + 1
        cols(n) = colIndex
    Next
    If n >= 0 Then
        ReDim Preserve cols(n)
        ResolveHeaderColumns = cols
    End If
End Function

Function UniqueSheetName(workbook, baseName)
    Dim candidate, index, exists, ws
    index = 1
    Do
        If index = 1 Then
            candidate = Left(CStr(baseName), 31)
        Else
            candidate = Left(CStr(baseName), 28) & CStr(index)
        End If
        exists = False
        For Each ws In workbook.Worksheets
            If StrComp(ws.Name, candidate, vbTextCompare) = 0 Then exists = True
        Next
        If Not exists Then
            UniqueSheetName = candidate
            Exit Function
        End If
        index = index + 1
    Loop
End Function

Function BuildTimestamp()
    Dim d
    d = Now
    BuildTimestamp = Year(d) & Right("0" & Month(d), 2) & Right("0" & Day(d), 2) & "_" & Right("0" & Hour(d), 2) & Right("0" & Minute(d), 2) & Right("0" & Second(d), 2)
End Function

Function BuildOutputSheetName(sourceSheetName, functionName)
    Dim baseName
    baseName = CStr(sourceSheetName) & "_" & CStr(functionName) & "_" & BuildTimestamp()
    baseName = Replace(baseName, "[", "(")
    baseName = Replace(baseName, "]", ")")
    baseName = Replace(baseName, ":", "-")
    baseName = Replace(baseName, "\\", "-")
    baseName = Replace(baseName, "/", "-")
    baseName = Replace(baseName, "?", "")
    baseName = Replace(baseName, "*", "")
    If Len(baseName) > 31 Then baseName = Left(baseName, 31)
    BuildOutputSheetName = baseName
End Function

Function CreateOutputSheet(appObj, sourceSheetName, functionName, outputMode)
    On Error Resume Next
    Dim hostJson, sheetName, workbook, sheetObj, modeText
    Set CreateOutputSheet = Nothing
    modeText = NormalizeOutputMode(outputMode)
    hostJson = SafeHostText("ExcelCreateOutputSheet", functionName, modeText, sourceSheetName, True)
    sheetName = ExtractJsonString(hostJson, "sheetName")
    If Len(sheetName) = 0 Then sheetName = ExtractJsonString(hostJson, "name")
    If Len(sheetName) > 0 Then
        Set sheetObj = Nothing
        Set sheetObj = appObj.ActiveWorkbook.Worksheets(sheetName)
        If Not sheetObj Is Nothing Then
            Set CreateOutputSheet = sheetObj
            Exit Function
        End If
    End If

    Set workbook = appObj.ActiveWorkbook
    If modeText = "NewWorkbook" Then
        Set workbook = appObj.Workbooks.Add
    End If
    sheetName = UniqueSheetName(workbook, BuildOutputSheetName(sourceSheetName, functionName))
    Set sheetObj = workbook.Worksheets.Add
    sheetObj.Name = sheetName
    If Err.Number <> 0 Then
        Err.Clear
        Set CreateOutputSheet = Nothing
        Exit Function
    End If
    Set CreateOutputSheet = sheetObj
End Function

Sub SafeDeleteSheet(appObj, sheetObj)
    On Error Resume Next
    Dim prev
    If sheetObj Is Nothing Then Exit Sub
    prev = appObj.DisplayAlerts
    appObj.DisplayAlerts = False
    sheetObj.Delete
    appObj.DisplayAlerts = prev
    Err.Clear
End Sub

Function NormalizeOutputMode(text)
    Dim t
    t = LCase(Trim(CStr(text)))
    If t = "newworkbook" Or t = "workbook" Or t = "新工作簿" Then
        NormalizeOutputMode = "NewWorkbook"
    Else
        NormalizeOutputMode = "NewSheet"
    End If
End Function

Sub WriteSummaryFallback(sheetObj, summaryLines)
    Dim lines, i
    lines = Split(CStr(summaryLines), vbCrLf)
    sheetObj.Cells(1, 1).Value = "项目"
    sheetObj.Cells(1, 2).Value = "内容"
    For i = 0 To UBound(lines)
        sheetObj.Cells(i + 2, 1).Value = "行" & CStr(i + 1)
        sheetObj.Cells(i + 2, 2).Value = lines(i)
    Next
    sheetObj.Rows(1).Font.Bold = True
    sheetObj.Columns.AutoFit
End Sub

Sub RollbackCreatedSheets(appObj, createdSheets, createdCount)
    Dim i
    For i = createdCount To 1 Step -1
        SafeDeleteSheet appObj, createdSheets(i)
        Set createdSheets(i) = Nothing
    Next
End Sub

Function IsBlankCell(value)
    IsBlankCell = (Len(NormalizeCellText(value)) = 0)
End Function

Function FindWorksheet(workbook, sheetName)
    On Error Resume Next
    Dim ws
    Set FindWorksheet = Nothing
    For Each ws In workbook.Worksheets
        If StrComp(ws.Name, Trim(CStr(sheetName)), vbTextCompare) = 0 Then
            Set FindWorksheet = ws
            Exit Function
        End If
    Next
End Function

Function FindOtherSheetName(workbook, currentName)
    Dim ws
    FindOtherSheetName = ""
    For Each ws In workbook.Worksheets
        If StrComp(ws.Name, currentName, vbTextCompare) <> 0 Then
            FindOtherSheetName = ws.Name
            Exit Function
        End If
    Next
End Function

Function FindKeyRow(dataRange, keyColumn, keyText)
    Dim rowIndex
    FindKeyRow = 0
    For rowIndex = 2 To dataRange.Rows.Count
        If NormalizeKey(dataRange.Cells(rowIndex, keyColumn).Text) = keyText Then
            FindKeyRow = rowIndex
            Exit Function
        End If
    Next
End Function

Function CountKeyOccurrences(dataRange, keyColumn, keyText)
    Dim rowIndex, total
    total = 0
    For rowIndex = 2 To dataRange.Rows.Count
        If NormalizeKey(dataRange.Cells(rowIndex, keyColumn).Text) = keyText Then total = total + 1
    Next
    CountKeyOccurrences = total
End Function

Function TargetAbsoluteRow(targetRange, relativeRow)
    If relativeRow > 0 Then TargetAbsoluteRow = targetRange.Row + relativeRow - 1 Else TargetAbsoluteRow = 0
End Function
Function DigitsOnly(value)
    Dim i, ch, buf
    buf = ""
    For i = 1 To Len(CStr(value))
        ch = Mid(CStr(value), i, 1)
        If ch >= "0" And ch <= "9" Then buf = buf & ch
    Next
    DigitsOnly = buf
End Function
