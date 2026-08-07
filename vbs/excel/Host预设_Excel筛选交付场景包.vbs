' 函数名: HostExcelFilterDeliveryPack
' 描述: 筛选交付场景包：支持 1-3 个条件与 AND/OR，预览命中后生成筛选副本（含_来源行）与处理摘要，不改源表
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
    Main = HostExcelFilterDeliveryPack(appObj)
End Function

Function HostExcelFilterDeliveryPack(appObj)
    On Error Resume Next
    Dim sourceRange, sourceSheet, sourceName, outputMode
    Set sourceRange = appObj.Selection.CurrentRegion
    If Err.Number <> 0 Or TypeName(sourceRange) = "Empty" Or TypeName(sourceRange) = "Nothing" Then
        Err.Clear
        HostExcelFilterDeliveryPack = FailureJson("E_NO_RANGE", "请先选中带表头的数据区域")
        Exit Function
    End If
    If sourceRange.Rows.Count < 2 Then
        HostExcelFilterDeliveryPack = FailureJson("E_RANGE_TOO_SMALL", "当前区域至少需要一行表头和一行数据")
        Exit Function
    End If
    Set sourceSheet = sourceRange.Worksheet
    sourceName = sourceSheet.Name
    outputMode = NormalizeOutputMode(SafePrompt("输出模式：NewSheet 或 NewWorkbook（默认 NewSheet）", "NewSheet"))

    Dim logicText, logicCode
    logicText = Trim(SafePrompt("条件关系：AND 或 OR", "AND"))
    logicCode = NormalizeLogic(logicText)

    Dim fields(3), keywords(3), modes(3), columns(3)
    Dim condCount, idx, defaultField
    condCount = 0
    defaultField = CStr(sourceRange.Cells(1, 1).Text)
    For idx = 1 To 3
        fields(idx) = Trim(SafePrompt("条件" & CStr(idx) & " 字段名（留空结束）", IIf(idx = 1, defaultField, "")))
        If Len(fields(idx)) = 0 Then Exit For
        columns(idx) = FindHeaderColumn(sourceRange, fields(idx))
        If columns(idx) <= 0 Then
            HostExcelFilterDeliveryPack = FailureJson("E_FIELD_NOT_FOUND", "未找到筛选字段：" & fields(idx))
            Exit Function
        End If
        keywords(idx) = SafePrompt("条件" & CStr(idx) & " 关键词", "")
        modes(idx) = NormalizeMatchMode(SafePrompt("条件" & CStr(idx) & " 匹配方式：精确 或 包含", "精确"))
        condCount = condCount + 1
    Next
    If condCount = 0 Then
        HostExcelFilterDeliveryPack = FailureJson("E_NO_CONDITION", "至少需要 1 个筛选条件")
        Exit Function
    End If

    Dim rowIndex, matchCount
    matchCount = 0
    For rowIndex = 2 To sourceRange.Rows.Count
        If RowMatchesConditions(sourceRange, rowIndex, columns, keywords, modes, condCount, logicCode) Then matchCount = matchCount + 1
    Next

    Dim previewText, planId, planPreview
    previewText = "Excel 筛选交付场景包预览（尚未写入）" & vbCrLf & _
        "命令ID：excel.filter_delivery_pack" & vbCrLf & _
        "源表=" & sourceName & "；区域=" & sourceRange.Address & vbCrLf & _
        "输出模式=" & outputMode & vbCrLf & _
        "条件=" & BuildConditionSummary(fields, keywords, modes, condCount, logicCode) & vbCrLf & _
        "数据行=" & CStr(sourceRange.Rows.Count - 1) & "；命中=" & CStr(matchCount) & vbCrLf & _
        "将生成：筛选结果（含_来源行） + 处理摘要；sourceUnchanged=true"

    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_filter_delivery_pack", "{""source"":""" & EscapeJson(sourceRange.Address) & """,""logic"":""" & EscapeJson(logicCode) & """,""conditions"":" & CStr(condCount) & ",""matches"":" & CStr(matchCount) & "}", "office.excel.filterDeliveryPack"
    planPreview = SafePreviewWritePlan()
    SafeRollbackWritePlan planId

    If Not SafeConfirmStep(previewText & vbCrLf & vbCrLf & "确认生成筛选交付结果？", "excel.filter_delivery_pack") Then
        HostExcelFilterDeliveryPack = FailureJson("E_CONFIRM_REQUIRED", "用户取消或未确认，未修改工作簿")
        Exit Function
    End If

    Dim createdSheets(), createdCount, filterSheet, summarySheet, outRows
    createdCount = 0
    ReDim createdSheets(4)
    outRows = 0

    Set filterSheet = CreateOutputSheet(appObj, sourceName, "筛选结果", outputMode)
    If filterSheet Is Nothing Then
        HostExcelFilterDeliveryPack = FailureJson("E_OUTPUT_SHEET", "无法创建筛选结果工作表")
        Exit Function
    End If
    createdCount = createdCount + 1: Set createdSheets(createdCount) = filterSheet

    If Not WriteFilteredSheet(sourceRange, columns, keywords, modes, condCount, logicCode, filterSheet, outRows) Then
        RollbackCreatedSheets appObj, createdSheets, createdCount
        HostExcelFilterDeliveryPack = FailureJson("E_FILTER_WRITE", "筛选结果写入失败，已删除未完成输出")
        Exit Function
    End If

    Set summarySheet = CreateOutputSheet(appObj, sourceName, "处理摘要", outputMode)
    If summarySheet Is Nothing Then
        RollbackCreatedSheets appObj, createdSheets, createdCount
        HostExcelFilterDeliveryPack = FailureJson("E_OUTPUT_SHEET", "无法创建处理摘要工作表")
        Exit Function
    End If
    createdCount = createdCount + 1: Set createdSheets(createdCount) = summarySheet

    Dim summaryLines, impactJson, summary
    summaryLines = "场景名=筛选交付场景包" & vbCrLf & _
        "命令ID=excel.filter_delivery_pack" & vbCrLf & _
        "源表名称=" & sourceName & vbCrLf & _
        "源区域=" & sourceRange.Address & vbCrLf & _
        "输出模式=" & outputMode & vbCrLf & _
        "输出位置=" & filterSheet.Name & "," & summarySheet.Name & vbCrLf & _
        "处理前行数=" & CStr(sourceRange.Rows.Count - 1) & vbCrLf & _
        "处理后行数=" & CStr(outRows) & vbCrLf & _
        "参数=logic=" & logicCode & ";conditions=" & CStr(condCount) & vbCrLf & _
        "条件=" & BuildConditionSummary(fields, keywords, modes, condCount, logicCode) & vbCrLf & _
        "命中行数=" & CStr(outRows) & vbCrLf & _
        "执行时间=" & Now & vbCrLf & _
        "结果=成功"
    impactJson = SafeHostText("ExcelWriteImpactSummary", summaryLines, summarySheet.Name)
    If Not ExtractJsonBoolean(impactJson, "ok") Then WriteSummaryFallback summarySheet, summaryLines

    summary = "Excel 筛选交付完成：输出=" & filterSheet.Name & "；命中=" & CStr(outRows) & "；条件数=" & CStr(condCount) & "；sourceUnchanged=true"
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostExcelFilterDeliveryPack = "{""ok"":true,""sourceUnchanged"":true,""commandId"":""excel.filter_delivery_pack"",""outputSheet"":""" & EscapeJson(filterSheet.Name) & """,""summarySheet"":""" & EscapeJson(summarySheet.Name) & """,""matchedRows"":" & CStr(outRows) & ",""conditions"":" & CStr(condCount) & ",""writePlanPreview"":""" & EscapeJson(planPreview) & """,""message"":""" & EscapeJson(summary) & """}"
End Function

Function RowMatchesConditions(sourceRange, rowIndex, ByRef columns, ByRef keywords, ByRef modes, condCount, logicCode)
    Dim i, matched, cellText, oneMatch
    If logicCode = "OR" Then
        matched = False
        For i = 1 To condCount
            cellText = CStr(sourceRange.Cells(rowIndex, columns(i)).Text)
            If IsMatch(cellText, keywords(i), modes(i)) Then
                matched = True
                Exit For
            End If
        Next
    Else
        matched = True
        For i = 1 To condCount
            cellText = CStr(sourceRange.Cells(rowIndex, columns(i)).Text)
            oneMatch = IsMatch(cellText, keywords(i), modes(i))
            If Not oneMatch Then
                matched = False
                Exit For
            End If
        Next
    End If
    RowMatchesConditions = matched
End Function

Function BuildConditionSummary(ByRef fields, ByRef keywords, ByRef modes, condCount, logicCode)
    Dim i, buf
    buf = "logic=" & logicCode
    For i = 1 To condCount
        buf = buf & ";[" & CStr(i) & "]" & fields(i) & " " & MatchModeLabel(modes(i)) & " " & keywords(i)
    Next
    BuildConditionSummary = buf
End Function

Function WriteFilteredSheet(sourceRange, ByRef columns, ByRef keywords, ByRef modes, condCount, logicCode, filterSheet, ByRef outRows)
    On Error Resume Next
    Dim c, r, outR, totalCols
    WriteFilteredSheet = False
    totalCols = sourceRange.Columns.Count
    For c = 1 To totalCols
        filterSheet.Cells(1, c).Value = sourceRange.Cells(1, c).Text
    Next
    filterSheet.Cells(1, totalCols + 1).Value = "_来源行"
    outR = 1
    For r = 2 To sourceRange.Rows.Count
        If RowMatchesConditions(sourceRange, r, columns, keywords, modes, condCount, logicCode) Then
            outR = outR + 1
            For c = 1 To totalCols
                filterSheet.Cells(outR, c).Value = sourceRange.Cells(r, c).Value
            Next
            filterSheet.Cells(outR, totalCols + 1).Value = sourceRange.Row + r - 1
        End If
    Next
    outRows = outR - 1
    filterSheet.Rows(1).Font.Bold = True
    filterSheet.Columns.AutoFit
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    WriteFilteredSheet = True
End Function


Function IsNothing(value)
    IsNothing = (TypeName(value) = "Nothing" Or TypeName(value) = "Empty")
End Function

Function ValuesEqual(sourceValue, targetValue)
    ValuesEqual = (StrComp(NormalizeCellText(sourceValue), NormalizeCellText(targetValue), vbTextCompare) = 0)
End Function

Function NormalizeExportFormat(value)
    Dim t
    t = LCase(Trim(CStr(value)))
    If t = "tsv" Or t = "tab" Or t = "制表符" Then
        NormalizeExportFormat = "tsv"
    Else
        NormalizeExportFormat = "csv"
    End If
End Function

Function FormatFileStamp(value)
    Dim d
    d = CDate(value)
    FormatFileStamp = Year(d) & Right("0" & Month(d), 2) & Right("0" & Day(d), 2) & "_" & Right("0" & Hour(d), 2) & Right("0" & Minute(d), 2) & Right("0" & Second(d), 2)
End Function

Function NormalizeLogic(value)
    Dim t
    t = UCase(Trim(CStr(value)))
    If t = "OR" Or t = "或" Or t = "或者" Then
        NormalizeLogic = "OR"
    Else
        NormalizeLogic = "AND"
    End If
End Function

Function NormalizeMatchMode(value)
    Dim t
    t = LCase(Trim(CStr(value)))
    If t = "contains" Or t = "包含" Or t = "like" Or t = "fuzzy" Then
        NormalizeMatchMode = "contains"
    Else
        NormalizeMatchMode = "exact"
    End If
End Function

Function MatchModeLabel(value)
    If NormalizeMatchMode(value) = "contains" Then
        MatchModeLabel = "包含"
    Else
        MatchModeLabel = "精确"
    End If
End Function

Function IsMatch(cellText, keyword, matchMode)
    Dim leftText, rightText
    leftText = NormalizeCellText(cellText)
    rightText = NormalizeCellText(keyword)
    If Len(rightText) = 0 Then
        IsMatch = (Len(leftText) = 0)
        Exit Function
    End If
    If NormalizeMatchMode(matchMode) = "contains" Then
        IsMatch = (InStr(1, leftText, rightText, vbTextCompare) > 0)
    Else
        IsMatch = (StrComp(leftText, rightText, vbTextCompare) = 0)
    End If
End Function

Function NormalizeMaskKind(value)
    Dim t
    t = LCase(Trim(CStr(value)))
    If t = "phone" Or t = "mobile" Or t = "手机" Or t = "手机号" Then
        NormalizeMaskKind = "phone"
    ElseIf t = "id" Or t = "idcard" Or t = "身份证" Then
        NormalizeMaskKind = "idcard"
    ElseIf t = "email" Or t = "mail" Or t = "邮箱" Then
        NormalizeMaskKind = "email"
    Else
        NormalizeMaskKind = "auto"
    End If
End Function

Function MaskKindLabel(value)
    Dim k
    k = NormalizeMaskKind(value)
    If k = "phone" Then
        MaskKindLabel = "手机号"
    ElseIf k = "idcard" Then
        MaskKindLabel = "身份证"
    ElseIf k = "email" Then
        MaskKindLabel = "邮箱"
    Else
        MaskKindLabel = "自动"
    End If
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

Function LooksLikeIdCard(value)
    Dim digits, text
    text = UCase(Trim(CStr(value)))
    digits = DigitsOnly(text)
    LooksLikeIdCard = (Len(digits) = 15 Or Len(digits) = 18 Or (Len(text) = 18 And Right(text, 1) = "X"))
End Function

Function MaskSensitiveValue(value, maskKind)
    Dim text, digits, kind, atPos, localPart, domainPart
    MaskSensitiveValue = ""
    text = NormalizeCellText(value)
    If Len(text) = 0 Then Exit Function
    kind = NormalizeMaskKind(maskKind)
    digits = DigitsOnly(text)
    If kind = "auto" Then
        If InStr(text, "@") > 0 Then
            kind = "email"
        ElseIf LooksLikeIdCard(text) Then
            kind = "idcard"
        ElseIf Len(digits) = 11 And Left(digits, 1) = "1" Then
            kind = "phone"
        Else
            Exit Function
        End If
    End If
    If kind = "phone" Then
        If Len(digits) < 7 Then Exit Function
        MaskSensitiveValue = Left(digits, 3) & "****" & Right(digits, 4)
    ElseIf kind = "idcard" Then
        If Len(digits) < 8 Then Exit Function
        If Len(digits) >= 18 Then
            MaskSensitiveValue = Left(digits, 4) & "**********" & Right(digits, 4)
        Else
            MaskSensitiveValue = Left(digits, 3) & "*********" & Right(digits, 3)
        End If
    ElseIf kind = "email" Then
        atPos = InStr(text, "@")
        If atPos <= 1 Then Exit Function
        localPart = Left(text, atPos - 1)
        domainPart = Mid(text, atPos)
        If Len(localPart) <= 1 Then
            MaskSensitiveValue = "*" & domainPart
        Else
            MaskSensitiveValue = Left(localPart, 1) & "***" & domainPart
        End If
    End If
End Function

Function RowHasContent(rowRange)
    Dim cell
    RowHasContent = False
    For Each cell In rowRange.Cells
        If Len(NormalizeCellText(cell.Text)) > 0 Then
            RowHasContent = True
            Exit Function
        End If
    Next
End Function

Function IIf(condition, trueValue, falseValue)
    If condition Then
        IIf = trueValue
    Else
        IIf = falseValue
    End If
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

