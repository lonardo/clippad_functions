' 函数名: HostExcelMultiSheetMergePack
' 描述: 多表合并场景包：先预览工作表列表与行数，确认后合并当前工作簿 UsedRange 到新表并保留来源表名，附带处理摘要
' 适用应用: Excel
' 搜索范围: 全文
' 搜索对象: 无
' 作用范围: 全文
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
    Main = HostExcelMultiSheetMergePack(appObj)
End Function

Function HostExcelMultiSheetMergePack(appObj)
    On Error Resume Next
    Dim workbook, activeName, outputMode
    Set workbook = appObj.ActiveWorkbook
    If Err.Number <> 0 Or TypeName(workbook) = "Empty" Or TypeName(workbook) = "Nothing" Then
        Err.Clear
        HostExcelMultiSheetMergePack = FailureJson("E_NO_WORKBOOK", "当前没有活动工作簿")
        Exit Function
    End If
    If workbook.Worksheets.Count < 2 Then
        HostExcelMultiSheetMergePack = FailureJson("E_SINGLE_SHEET", "当前工作簿少于两个工作表，无需合并")
        Exit Function
    End If
    activeName = workbook.ActiveSheet.Name
    outputMode = NormalizeOutputMode(SafePrompt("输出模式：NewSheet 或 NewWorkbook（默认 NewSheet）", "NewSheet"))

    Dim ws, rng, previewLines, sheetCount, dataRows, sampleCount, rowCount
    previewLines = ""
    sheetCount = 0
    dataRows = 0
    sampleCount = 0
    For Each ws In workbook.Worksheets
        Set rng = ws.UsedRange
        If Err.Number <> 0 Then
            Err.Clear
            rowCount = 0
        ElseIf rng Is Nothing Then
            rowCount = 0
        ElseIf rng.Rows.Count <= 1 Then
            rowCount = 0
        Else
            rowCount = CountContentRows(rng)
        End If
        sheetCount = sheetCount + 1
        dataRows = dataRows + rowCount
        If sampleCount < 12 Then
            If Len(previewLines) > 0 Then previewLines = previewLines & vbCrLf
            previewLines = previewLines & "- " & ws.Name & " : " & CStr(rowCount) & " 行"
            sampleCount = sampleCount + 1
        End If
    Next
    Dim previewText, planId, planPreview
    previewText = "Excel 多表合并场景包预览（尚未写入）" & vbCrLf & _
        "命令ID：excel.multi_sheet_merge_pack" & vbCrLf & _
        "工作簿工作表数=" & CStr(sheetCount) & "；预估数据行=" & CStr(dataRows) & vbCrLf & _
        "输出模式=" & outputMode & vbCrLf & _
        "工作表列表（最多 12 个）：" & vbCrLf & previewLines & vbCrLf & _
        "将生成：合并结果（含来源工作表列） + 处理摘要；sourceUnchanged=true"
    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_multi_sheet_merge_pack", "{""sheetCount"":" & CStr(sheetCount) & ",""estimatedRows"":" & CStr(dataRows) & ",""outputMode"":""" & EscapeJson(outputMode) & """}", "office.excel.multiSheetMergePack"
    planPreview = SafePreviewWritePlan()
    SafeRollbackWritePlan planId
    If Not SafeConfirmStep(previewText & vbCrLf & vbCrLf & "确认合并当前工作簿多表？", "excel.multi_sheet_merge_pack") Then
        HostExcelMultiSheetMergePack = FailureJson("E_CONFIRM_REQUIRED", "用户取消或未确认，未修改工作簿")
        Exit Function
    End If
    Dim createdSheets(), createdCount, mergeSheet, summarySheet, mergedSheets, mergedRows, headerWritten, colCount
    createdCount = 0
    ReDim createdSheets(4)
    mergedSheets = 0
    mergedRows = 0
    headerWritten = False
    colCount = 0
    Set mergeSheet = CreateOutputSheet(appObj, activeName, "多表合并", outputMode)
    If mergeSheet Is Nothing Then
        HostExcelMultiSheetMergePack = FailureJson("E_OUTPUT_SHEET", "无法创建多表合并输出工作表")
        Exit Function
    End If
    createdCount = createdCount + 1: Set createdSheets(createdCount) = mergeSheet
    If Not MergeWorkbookToSheet(workbook, mergeSheet, mergedSheets, mergedRows, headerWritten, colCount) Then
        RollbackCreatedSheets appObj, createdSheets, createdCount
        HostExcelMultiSheetMergePack = FailureJson("E_MERGE_WRITE", "多表合并写入失败，已删除未完成输出")
        Exit Function
    End If
    Set summarySheet = CreateOutputSheet(appObj, activeName, "处理摘要", outputMode)
    If summarySheet Is Nothing Then
        RollbackCreatedSheets appObj, createdSheets, createdCount
        HostExcelMultiSheetMergePack = FailureJson("E_OUTPUT_SHEET", "无法创建处理摘要工作表")
        Exit Function
    End If
    createdCount = createdCount + 1: Set createdSheets(createdCount) = summarySheet
    Dim summaryLines, impactJson, summary
    summaryLines = "场景名=多表合并场景包" & vbCrLf & _
        "命令ID=excel.multi_sheet_merge_pack" & vbCrLf & _
        "源表名称=" & activeName & vbCrLf & _
        "源区域=UsedRange" & vbCrLf & _
        "输出模式=" & outputMode & vbCrLf & _
        "输出位置=" & mergeSheet.Name & "," & summarySheet.Name & vbCrLf & _
        "处理前行数=" & CStr(dataRows) & vbCrLf & _
        "处理后行数=" & CStr(mergedRows) & vbCrLf & _
        "参数=mergedSheets=" & CStr(mergedSheets) & ";columns=" & CStr(colCount) & vbCrLf & _
        "合并工作表数=" & CStr(mergedSheets) & vbCrLf & _
        "合并数据行=" & CStr(mergedRows) & vbCrLf & _
        "执行时间=" & Now & vbCrLf & _
        "结果=成功"
    impactJson = SafeHostText("ExcelWriteImpactSummary", summaryLines, summarySheet.Name)
    If Not ExtractJsonBoolean(impactJson, "ok") Then WriteSummaryFallback summarySheet, summaryLines
    summary = "Excel 多表合并完成：输出=" & mergeSheet.Name & "；合并表=" & CStr(mergedSheets) & "；行=" & CStr(mergedRows) & "；sourceUnchanged=true"
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostExcelMultiSheetMergePack = "{""ok"":true,""sourceUnchanged"":true,""commandId"":""excel.multi_sheet_merge_pack"",""outputSheet"":""" & EscapeJson(mergeSheet.Name) & """,""summarySheet"":""" & EscapeJson(summarySheet.Name) & """,""mergedSheets"":" & CStr(mergedSheets) & ",""mergedRows"":" & CStr(mergedRows) & ",""writePlanPreview"":""" & EscapeJson(planPreview) & """,""message"":""" & EscapeJson(summary) & """}"
End Function

Function CountContentRows(rng)
    Dim r, total
    total = 0
    If rng Is Nothing Then
        CountContentRows = 0
        Exit Function
    End If
    For r = 2 To rng.Rows.Count
        If RowHasContent(rng.Rows(r)) Then total = total + 1
    Next
    CountContentRows = total
End Function

Function MergeWorkbookToSheet(workbook, mergeSheet, ByRef mergedSheets, ByRef mergedRows, ByRef headerWritten, ByRef colCount)
    On Error Resume Next
    Dim ws, rng, r, c, outRow, sourceCols
    MergeWorkbookToSheet = False
    outRow = 1
    mergedSheets = 0
    mergedRows = 0
    headerWritten = False
    colCount = 0
    mergeSheet.Cells(1, 1).Value = "来源工作表"

    For Each ws In workbook.Worksheets
        If StrComp(ws.Name, mergeSheet.Name, vbTextCompare) <> 0 Then
            Set rng = ws.UsedRange
            If Err.Number = 0 And Not rng Is Nothing Then
                If rng.Rows.Count >= 2 And rng.Columns.Count >= 1 Then
                sourceCols = rng.Columns.Count
                If Not headerWritten Then
                    For c = 1 To sourceCols
                        mergeSheet.Cells(1, c + 1).Value = rng.Cells(1, c).Text
                    Next
                    colCount = sourceCols
                    headerWritten = True
                    outRow = 1
                End If
                For r = 2 To rng.Rows.Count
                    If RowHasContent(rng.Rows(r)) Then
                        outRow = outRow + 1
                        mergeSheet.Cells(outRow, 1).Value = ws.Name
                        For c = 1 To sourceCols
                            mergeSheet.Cells(outRow, c + 1).Value = rng.Cells(r, c).Value
                        Next
                        mergedRows = mergedRows + 1
                    End If
                Next
                mergedSheets = mergedSheets + 1
                End If
            Else
                Err.Clear
            End If
        End If
    Next

    If Not headerWritten Or mergedRows = 0 Then
        Exit Function
    End If
    mergeSheet.Rows(1).Font.Bold = True
    mergeSheet.Columns.AutoFit
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    MergeWorkbookToSheet = True
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

