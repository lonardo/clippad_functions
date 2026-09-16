
' 函数名: HostExcelDedupeKeepMovePack
' 描述: 去重留痕场景包：按主键预览保留行与重复行，确认后分别写入去重表与重复项表并生成处理摘要，默认不删源行
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
    Main = HostExcelDedupeKeepMovePack(appObj)
End Function

Function HostExcelDedupeKeepMovePack(appObj)
    On Error Resume Next
    Dim sourceRange, sourceSheet, sourceName
    Set sourceRange = appObj.Selection.CurrentRegion
    If Err.Number <> 0 Or TypeName(sourceRange) = "Empty" Or TypeName(sourceRange) = "Nothing" Then
        Err.Clear
        HostExcelDedupeKeepMovePack = FailureJson("E_NO_RANGE", "请先选中含表头的数据区域")
        Exit Function
    End If
    If sourceRange.Rows.Count < 2 Or sourceRange.Columns.Count < 1 Then
        HostExcelDedupeKeepMovePack = FailureJson("E_RANGE_TOO_SMALL", "当前区域至少需要一行表头和一行数据")
        Exit Function
    End If

    Set sourceSheet = sourceRange.Worksheet
    sourceName = sourceSheet.Name

    Dim outputMode, keyFieldsText, keepPolicy, keyNames, keyCols
    outputMode = NormalizeOutputMode(SafePrompt("输出模式：NewSheet 或 NewWorkbook（默认 NewSheet）", "NewSheet"))
    keyFieldsText = Trim(SafePrompt("去重键字段（逗号分隔，需与表头一致）", CStr(sourceRange.Cells(1, 1).Text)))
    keepPolicy = NormalizeKeepPolicy(SafePrompt("保留策略：first 或 last（默认 first）", "first"))
    keyNames = SplitCsvFields(keyFieldsText)
    keyCols = ResolveHeaderColumns(sourceRange, keyNames)
    If (Not IsArray(keyCols)) Or UBound(keyCols) < 0 Then
        HostExcelDedupeKeepMovePack = FailureJson("E_KEY_NOT_FOUND", "未在表头中找到去重键字段：" & keyFieldsText)
        Exit Function
    End If

    Dim headerMap, sampleJson, dedupeJson, keepRows, dupRows, keepCount, dupCount, truncated
    headerMap = SafeHostText("ExcelGetHeaderMap", "")
    sampleJson = SafeHostText("ExcelPreviewRowSamples", 3, keyFieldsText)
    dedupeJson = SafeHostText("ExcelDedupePlan", keyFieldsText, keepPolicy)
    If Not ExtractJsonBoolean(dedupeJson, "ok") Then
        dedupeJson = LocalDedupePlan(sourceRange, keyCols, keepPolicy)
    End If
    keepRows = ParseJsonIntArray(dedupeJson, "keepRows")
    dupRows = ParseJsonIntArray(dedupeJson, "duplicateRows")
    keepCount = CLng(ExtractJsonNumber(dedupeJson, "keepCount"))
    dupCount = CLng(ExtractJsonNumber(dedupeJson, "duplicateCount"))
    If keepCount <= 0 And IsArray(keepRows) Then
        On Error Resume Next
        keepCount = UBound(keepRows) + 1
        If Err.Number <> 0 Then keepCount = 0
        Err.Clear
    End If
    If dupCount <= 0 And IsArray(dupRows) Then
        On Error Resume Next
        If UBound(dupRows) >= 0 Then dupCount = UBound(dupRows) + 1
        Err.Clear
    End If
    truncated = ExtractJsonBoolean(dedupeJson, "truncated")

    Dim previewText, planId, planPreview
    previewText = "Excel 去重留痕场景包预览（尚未写入）" & vbCrLf & _
        "命令ID：excel.dedupe_keep_move" & vbCrLf & _
        "源表：" & sourceName & " 区域=" & sourceRange.Address & vbCrLf & _
        "去重键=" & keyFieldsText & "；保留策略=" & keepPolicy & "；输出模式=" & outputMode & vbCrLf & _
        "源数据行=" & CStr(sourceRange.Rows.Count - 1) & "；保留行=" & CStr(keepCount) & "；重复行=" & CStr(dupCount) & vbCrLf
    If truncated Then previewText = previewText & "注意：去重计划结果已截断，预览行数可能少于实际。" & vbCrLf
    previewText = previewText & "将生成：去重表 + 重复项表（含_来源行/_重复键/_保留策略）+ 处理摘要；sourceUnchanged=true；不会 RemoveDuplicates 静默删源行"

    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_dedupe_keep_move", "{""source"":""" & EscapeJson(sourceRange.Address) & """,""keys"":""" & EscapeJson(keyFieldsText) & """,""keepPolicy"":""" & EscapeJson(keepPolicy) & """,""keepCount"":" & CStr(keepCount) & ",""duplicateCount"":" & CStr(dupCount) & "}", "office.excel.dedupePack"
    planPreview = SafePreviewWritePlan()
    SafeRollbackWritePlan planId

    If Not SafeConfirmStep(previewText & vbCrLf & vbCrLf & "确认生成去重表与重复项表？", "excel.dedupe_keep_move") Then
        HostExcelDedupeKeepMovePack = FailureJson("E_CONFIRM_REQUIRED", "用户取消或未确认，未修改工作簿")
        Exit Function
    End If

    Dim dedupeSheet, dupSheet, summarySheet, createdSheets(), createdCount
    Dim dedupeOutRows, dupOutRows
    createdCount = 0
    ReDim createdSheets(6)
    Set dedupeSheet = CreateOutputSheet(appObj, sourceName, "去重", outputMode)
    If dedupeSheet Is Nothing Then
        HostExcelDedupeKeepMovePack = FailureJson("E_OUTPUT_SHEET", "无法创建去重输出工作表")
        Exit Function
    End If
    createdCount = createdCount + 1: Set createdSheets(createdCount) = dedupeSheet

    Set dupSheet = CreateOutputSheet(appObj, sourceName, "重复项", outputMode)
    If dupSheet Is Nothing Then
        RollbackCreatedSheets appObj, createdSheets, createdCount
        HostExcelDedupeKeepMovePack = FailureJson("E_OUTPUT_SHEET", "无法创建重复项输出工作表")
        Exit Function
    End If
    createdCount = createdCount + 1: Set createdSheets(createdCount) = dupSheet

    If Not WriteDedupeAndDupSheets(sourceRange, keyCols, keepPolicy, keepRows, dupRows, dedupeSheet, dupSheet, dedupeOutRows, dupOutRows) Then
        RollbackCreatedSheets appObj, createdSheets, createdCount
        HostExcelDedupeKeepMovePack = FailureJson("E_DEDUPE_WRITE", "去重/重复项写入失败，已删除未完成输出")
        Exit Function
    End If

    Set summarySheet = CreateOutputSheet(appObj, sourceName, "处理摘要", outputMode)
    If summarySheet Is Nothing Then
        RollbackCreatedSheets appObj, createdSheets, createdCount
        HostExcelDedupeKeepMovePack = FailureJson("E_OUTPUT_SHEET", "无法创建处理摘要工作表")
        Exit Function
    End If
    createdCount = createdCount + 1: Set createdSheets(createdCount) = summarySheet

    Dim summaryLines, impactJson, summary
    summaryLines = "场景名=去重留痕场景包" & vbCrLf & _
        "命令ID=excel.dedupe_keep_move" & vbCrLf & _
        "源表名称=" & sourceName & vbCrLf & _
        "源区域=" & sourceRange.Address & vbCrLf & _
        "输出模式=" & outputMode & vbCrLf & _
        "输出位置=" & dedupeSheet.Name & "," & dupSheet.Name & vbCrLf & _
        "处理前行数=" & CStr(sourceRange.Rows.Count - 1) & vbCrLf & _
        "处理后行数=" & CStr(dedupeOutRows) & vbCrLf & _
        "参数=keepPolicy=" & keepPolicy & vbCrLf & _
        "删除空行=0" & vbCrLf & _
        "删除空列=0" & vbCrLf & _
        "去重键=" & keyFieldsText & vbCrLf & _
        "重复行数=" & CStr(dupOutRows) & vbCrLf & _
        "规范化单元格=0" & vbCrLf & _
        "执行时间=" & Now & vbCrLf & _
        "结果=成功"
    impactJson = SafeHostText("ExcelWriteImpactSummary", summaryLines, summarySheet.Name)
    If Not ExtractJsonBoolean(impactJson, "ok") Then WriteSummaryFallback summarySheet, summaryLines

    summary = "Excel 去重留痕完成：去重表=" & dedupeSheet.Name & "；重复项=" & dupSheet.Name & _
        "；保留行=" & CStr(dedupeOutRows) & "；重复行=" & CStr(dupOutRows) & "；sourceUnchanged=true"
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostExcelDedupeKeepMovePack = "{""ok"":true,""sourceUnchanged"":true,""commandId"":""excel.dedupe_keep_move"",""outputSheet"":""" & EscapeJson(dedupeSheet.Name) & """,""duplicateSheet"":""" & EscapeJson(dupSheet.Name) & """,""summarySheet"":""" & EscapeJson(summarySheet.Name) & """,""keepRows"":" & CStr(dedupeOutRows) & ",""duplicateRows"":" & CStr(dupOutRows) & ",""sourceRowsDeleted"":false,""writePlanPreview"":""" & EscapeJson(planPreview) & """,""message"":""" & EscapeJson(summary) & """}"
End Function

Function NormalizeOutputMode(text)
    Dim t
    t = LCase(Trim(CStr(text)))
    If t = "newworkbook" Or t = "workbook" Then
        NormalizeOutputMode = "NewWorkbook"
    Else
        NormalizeOutputMode = "NewSheet"
    End If
End Function

Function NormalizeKeepPolicy(text)
    Dim t
    t = LCase(Trim(CStr(text)))
    If t = "last" Or t = "最后" Or t = "末条" Then
        NormalizeKeepPolicy = "last"
    Else
        NormalizeKeepPolicy = "first"
    End If
End Function

Function WriteDedupeAndDupSheets(sourceRange, keyCols, keepPolicy, keepRows, dupRows, dedupeSheet, dupSheet, ByRef dedupeOutRows, ByRef dupOutRows)
    On Error Resume Next
    Dim c, totalCols, r, absRow, outR, i, key, useLocal, localJson
    WriteDedupeAndDupSheets = False
    totalCols = sourceRange.Columns.Count
    For c = 1 To totalCols
        dedupeSheet.Cells(1, c).Value = sourceRange.Cells(1, c).Text
        dupSheet.Cells(1, c).Value = sourceRange.Cells(1, c).Text
    Next
    dupSheet.Cells(1, totalCols + 1).Value = "_来源行"
    dupSheet.Cells(1, totalCols + 2).Value = "_重复键"
    dupSheet.Cells(1, totalCols + 3).Value = "_保留策略"

    useLocal = True
    If IsArray(keepRows) Then
        On Error Resume Next
        If UBound(keepRows) >= 0 Then useLocal = False
        Err.Clear
    End If
    If useLocal Then
        localJson = LocalDedupePlan(sourceRange, keyCols, keepPolicy)
        keepRows = ParseJsonIntArray(localJson, "keepRows")
        dupRows = ParseJsonIntArray(localJson, "duplicateRows")
    End If

    outR = 1
    If IsArray(keepRows) Then
        For i = 0 To UBound(keepRows)
            absRow = CLng(keepRows(i))
            r = absRow - sourceRange.Row + 1
            If r >= 2 And r <= sourceRange.Rows.Count Then
                outR = outR + 1
                For c = 1 To totalCols
                    dedupeSheet.Cells(outR, c).Value = sourceRange.Cells(r, c).Value
                Next
            End If
        Next
    End If
    dedupeOutRows = outR - 1
    dedupeSheet.Rows(1).Font.Bold = True
    dedupeSheet.Columns.AutoFit

    outR = 1
    If IsArray(dupRows) Then
        For i = 0 To UBound(dupRows)
            absRow = CLng(dupRows(i))
            r = absRow - sourceRange.Row + 1
            If r >= 2 And r <= sourceRange.Rows.Count Then
                outR = outR + 1
                For c = 1 To totalCols
                    dupSheet.Cells(outR, c).Value = sourceRange.Cells(r, c).Value
                Next
                key = RowKeyFromColumns(sourceRange, r, keyCols)
                dupSheet.Cells(outR, totalCols + 1).Value = absRow
                dupSheet.Cells(outR, totalCols + 2).Value = key
                dupSheet.Cells(outR, totalCols + 3).Value = keepPolicy
            End If
        Next
    End If
    dupOutRows = outR - 1
    dupSheet.Rows(1).Font.Bold = True
    dupSheet.Columns.AutoFit
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    WriteDedupeAndDupSheets = True
End Function

Sub WriteSummaryFallback(sheetObj, summaryLines)
    Dim lines, i, pair
    lines = Split(CStr(summaryLines), vbCrLf)
    sheetObj.Cells(1, 1).Value = "字段"
    sheetObj.Cells(1, 2).Value = "值"
    For i = 0 To UBound(lines)
        pair = Split(lines(i), "=")
        If UBound(pair) >= 1 Then
            sheetObj.Cells(i + 2, 1).Value = pair(0)
            sheetObj.Cells(i + 2, 2).Value = Mid(lines(i), Len(pair(0)) + 2)
        Else
            sheetObj.Cells(i + 2, 1).Value = lines(i)
        End If
    Next
    sheetObj.Rows(1).Font.Bold = True
    sheetObj.Columns.AutoFit
End Sub

Sub RollbackCreatedSheets(appObj, createdSheets, createdCount)
    Dim i
    For i = createdCount To 1 Step -1
        SafeDeleteSheet appObj, createdSheets(i)
    Next
End Sub


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
        If Len(Trim(CStr(parts(i)))) > 0 Then
            n = n + 1
            values(n) = CLng(Val(Trim(CStr(parts(i)))))
        End If
    Next
    If n < 0 Then Exit Function
    ReDim Preserve values(n)
    ParseJsonIntArray = values
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
    EscapeJson = text
End Function

Function SafePrompt(message, defaultValue)
    On Error Resume Next
    SafePrompt = Host.Prompt(message, defaultValue)
    If Err.Number <> 0 Then SafePrompt = defaultValue
    Err.Clear
End Function

Function SafeConfirmStep(message, stepInfo)
    On Error Resume Next
    Dim confirmed, answer
    confirmed = Host.ConfirmStep(message, stepInfo)
    If Err.Number = 0 Then
        SafeConfirmStep = CBool(confirmed)
        Err.Clear
        Exit Function
    End If
    Err.Clear
    answer = Trim(CStr(SafePrompt(message & vbCrLf & vbCrLf & "确认请输入：确认", "")))
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
        If StrComp(Trim(CStr(sourceRange.Cells(1, colIndex).Text)), Trim(CStr(fieldName)), vbTextCompare) = 0 Then
            FindHeaderColumn = colIndex
            Exit Function
        End If
    Next
End Function

Function SplitCsvFields(text)
    Dim cleaned, parts, i, n, values()
    cleaned = Replace(Replace(Replace(Trim(CStr(text)), "，", ","), "；", ","), ";", ",")
    If Len(cleaned) = 0 Then
        SplitCsvFields = Array()
        Exit Function
    End If
    parts = Split(cleaned, ",")
    n = -1
    ReDim values(UBound(parts))
    For i = 0 To UBound(parts)
        If Len(Trim(CStr(parts(i)))) > 0 Then
            n = n + 1
            values(n) = Trim(CStr(parts(i)))
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
    Dim i, col, cols(), n
    n = -1
    If IsArray(fieldNames) Then
        ReDim cols(UBound(fieldNames))
        For i = 0 To UBound(fieldNames)
            col = FindHeaderColumn(sourceRange, fieldNames(i))
            If col > 0 Then
                n = n + 1
                cols(n) = col
            End If
        Next
    End If
    If n < 0 Then
        ResolveHeaderColumns = Array()
    Else
        ReDim Preserve cols(n)
        ResolveHeaderColumns = cols
    End If
End Function

Function UniqueSheetName(workbook, baseName)
    Dim candidate, index, exists, sheetObj
    candidate = Left(CStr(baseName), 31)
    index = 1
    Do
        exists = False
        On Error Resume Next
        Set sheetObj = workbook.Worksheets(candidate)
        If Err.Number = 0 Then exists = True
        Err.Clear
        On Error GoTo 0
        If Not exists Then
            UniqueSheetName = candidate
            Exit Function
        End If
        index = index + 1
        candidate = Left(CStr(baseName), 27) & "_" & CStr(index)
    Loop
End Function

Function BuildTimestamp()
    BuildTimestamp = Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "_" & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
End Function

Function BuildOutputSheetName(sourceName, functionName)
    Dim baseName
    baseName = CStr(sourceName) & "_" & CStr(functionName) & "_" & BuildTimestamp()
    If Len(baseName) > 31 Then
        baseName = Left(CStr(sourceName), 8) & "_" & CStr(functionName) & "_" & BuildTimestamp()
    End If
    If Len(baseName) > 31 Then baseName = Left(baseName, 31)
    BuildOutputSheetName = baseName
End Function

Function CreateOutputSheet(appObj, sourceSheetName, functionName, outputMode)
    On Error Resume Next
    Dim hostJson, sheetName, workbook, sheetObj, created
    CreateOutputSheet = Nothing
    hostJson = SafeHostText("ExcelCreateOutputSheet", functionName, outputMode, sourceSheetName, True)
    sheetName = ExtractJsonString(hostJson, "sheetName")
    If ExtractJsonBoolean(hostJson, "ok") And Len(sheetName) > 0 Then
        Set workbook = appObj.ActiveWorkbook
        Set sheetObj = workbook.Worksheets(sheetName)
        If Err.Number = 0 And Not sheetObj Is Nothing Then
            Set CreateOutputSheet = sheetObj
            Err.Clear
            Exit Function
        End If
        Err.Clear
    End If

    Set workbook = appObj.ActiveWorkbook
    sheetName = UniqueSheetName(workbook, BuildOutputSheetName(sourceSheetName, functionName))
    Set sheetObj = workbook.Worksheets.Add
    sheetObj.Name = sheetName
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    Set CreateOutputSheet = sheetObj
End Function

Sub SafeDeleteSheet(appObj, sheetObj)
    On Error Resume Next
    Dim oldAlerts
    If sheetObj Is Nothing Then Exit Sub
    oldAlerts = appObj.DisplayAlerts
    appObj.DisplayAlerts = False
    sheetObj.Delete
    appObj.DisplayAlerts = oldAlerts
    Err.Clear
End Sub

Function IsBlankCell(value)
    IsBlankCell = (Len(NormalizeCellText(value)) = 0)
End Function

Function NormalizePhoneDigits(value)
    Dim text, i, ch, digits
    text = CStr(value)
    digits = ""
    For i = 1 To Len(text)
        ch = Mid(text, i, 1)
        If ch >= "0" And ch <= "9" Then digits = digits & ch
    Next
    If Left(digits, 2) = "86" And Len(digits) > 11 Then digits = Mid(digits, 3)
    NormalizePhoneDigits = digits
End Function

Function TryNormalizeDateText(value, ByRef normalized)
    Dim text, parts, y, m, d
    TryNormalizeDateText = False
    normalized = ""
    text = NormalizeCellText(value)
    If Len(text) = 0 Then Exit Function
    If IsDate(text) Then
        normalized = FormatDateTime(CDate(text), vbShortDate)
        TryNormalizeDateText = True
        Exit Function
    End If
    text = Replace(Replace(Replace(text, ".", "-"), "/", "-"), "年", "-")
    text = Replace(Replace(text, "月", "-"), "日", "")
    parts = Split(text, "-")
    If UBound(parts) >= 2 Then
        y = Val(parts(0))
        m = Val(parts(1))
        d = Val(parts(2))
        If y > 1900 And m >= 1 And m <= 12 And d >= 1 And d <= 31 Then
            normalized = CStr(CLng(y)) & "-" & Right("0" & CStr(CLng(m)), 2) & "-" & Right("0" & CStr(CLng(d)), 2)
            TryNormalizeDateText = True
        End If
    End If
End Function

Function RowKeyFromColumns(sourceRange, rowIndex, keyCols)
    Dim i, parts
    parts = ""
    If IsArray(keyCols) Then
        For i = 0 To UBound(keyCols)
            If i > 0 Then parts = parts & Chr(31)
            parts = parts & NormalizeKey(sourceRange.Cells(rowIndex, keyCols(i)).Text)
        Next
    End If
    RowKeyFromColumns = parts
End Function

Function LocalDedupePlan(sourceRange, keyCols, keepPolicy)
    Dim rowCount, i, j, key, keys(), firstRow(), lastRow(), counts(), keyCount
    Dim keepRows(), dupRows(), keepCount, dupCount, keepLast, chosen
    rowCount = sourceRange.Rows.Count
    keyCount = 0
    ReDim keys(rowCount)
    ReDim firstRow(rowCount)
    ReDim lastRow(rowCount)
    ReDim counts(rowCount)
    For i = 2 To rowCount
        key = RowKeyFromColumns(sourceRange, i, keyCols)
        If Len(key) = 0 Then key = "__empty__" & CStr(i)
        chosen = 0
        For j = 1 To keyCount
            If keys(j) = key Then
                chosen = j
                Exit For
            End If
        Next
        If chosen = 0 Then
            keyCount = keyCount + 1
            keys(keyCount) = key
            firstRow(keyCount) = sourceRange.Row + i - 1
            lastRow(keyCount) = sourceRange.Row + i - 1
            counts(keyCount) = 1
        Else
            lastRow(chosen) = sourceRange.Row + i - 1
            counts(chosen) = counts(chosen) + 1
        End If
    Next

    keepLast = (LCase(Trim(CStr(keepPolicy))) = "last")
    keepCount = 0
    dupCount = 0
    ReDim keepRows(rowCount)
    ReDim dupRows(rowCount)
    For i = 2 To rowCount
        key = RowKeyFromColumns(sourceRange, i, keyCols)
        If Len(key) = 0 Then key = "__empty__" & CStr(i)
        chosen = 0
        For j = 1 To keyCount
            If keys(j) = key Then
                chosen = j
                Exit For
            End If
        Next
        Dim absRow, keepAbs
        absRow = sourceRange.Row + i - 1
        If keepLast Then
            keepAbs = lastRow(chosen)
        Else
            keepAbs = firstRow(chosen)
        End If
        If absRow = keepAbs Then
            keepCount = keepCount + 1
            keepRows(keepCount) = absRow
        Else
            dupCount = dupCount + 1
            dupRows(dupCount) = absRow
        End If
    Next

    Dim result
    result = "{""ok"":true,""keepPolicy"":""" & EscapeJson(keepPolicy) & """,""keepCount"":" & CStr(keepCount) & ",""duplicateCount"":" & CStr(dupCount) & ",""keepRows"":["
    For i = 1 To keepCount
        If i > 1 Then result = result & ","
        result = result & CStr(keepRows(i))
    Next
    result = result & "],""duplicateRows"":["
    For i = 1 To dupCount
        If i > 1 Then result = result & ","
        result = result & CStr(dupRows(i))
    Next
    result = result & "],""truncated"":false}"
    LocalDedupePlan = result
End Function
