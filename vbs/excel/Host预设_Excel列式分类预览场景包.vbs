' 函数名: HostExcelClassifyColumnPreviewPack
' 描述: 列式分类预览场景包：按类别:关键词规则先预览约20行，确认后写分类结果新表，不改源表
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
    Main = HostExcelClassifyColumnPreviewPack(appObj)
End Function

Function HostExcelClassifyColumnPreviewPack(appObj)
    On Error Resume Next
    Dim sourceRange, sourceSheet, sourceName, outputMode
    Set sourceRange = appObj.Selection.CurrentRegion
    If Err.Number <> 0 Or TypeName(sourceRange) = "Empty" Or TypeName(sourceRange) = "Nothing" Then
        Err.Clear
        HostExcelClassifyColumnPreviewPack = FailureJson("E_NO_RANGE", "请先选中带表头的数据区域")
        Exit Function
    End If
    If sourceRange.Rows.Count < 2 Then
        HostExcelClassifyColumnPreviewPack = FailureJson("E_RANGE_TOO_SMALL", "当前区域至少需要一行表头和一行数据")
        Exit Function
    End If
    Set sourceSheet = sourceRange.Worksheet
    sourceName = sourceSheet.Name
    outputMode = NormalizeOutputMode(SafePrompt("输出模式：NewSheet 或 NewWorkbook（默认 NewSheet）", "NewSheet"))

    Dim sourceField, sourceCol, ruleText, catNames(), catRules(), catCount, defaultOther
    sourceField = Trim(SafePrompt("待分类文本列（表头名）", CStr(sourceRange.Cells(1, 1).Text)))
    sourceCol = FindHeaderColumn(sourceRange, sourceField)
    If sourceCol <= 0 Then
        HostExcelClassifyColumnPreviewPack = FailureJson("E_FIELD_NOT_FOUND", "未找到待分类列：" & sourceField)
        Exit Function
    End If
    ruleText = Trim(SafePrompt("分类规则（类别:关键词1|关键词2，多条用分号）", "投诉:投诉|差评|不满；咨询:咨询|请问|了解；表扬:表扬|满意|点赞"))
    catCount = ParseClassifyRules(ruleText, catNames, catRules)
    If catCount <= 0 Then
        HostExcelClassifyColumnPreviewPack = FailureJson("E_NO_RULE", "至少需要 1 条分类规则，格式：类别:关键词1|关键词2")
        Exit Function
    End If
    defaultOther = Trim(SafePrompt("未命中时的默认类别", "其他"))
    If Len(defaultOther) = 0 Then defaultOther = "其他"

    Dim dataRows, previewRows, previewText, i, rowIndex, cellText, label, matchedRule
    Dim hitCounts(), otherCount
    dataRows = sourceRange.Rows.Count - 1
    previewRows = dataRows
    If previewRows > 20 Then previewRows = 20
    ReDim hitCounts(catCount)
    For i = 1 To catCount
        hitCounts(i) = 0
    Next
    otherCount = 0

    previewText = "Excel 列式分类预览场景包（先预览约20行，尚未写入）" & vbCrLf & _
        "命令ID：excel.classify_column_preview" & vbCrLf & _
        "源表=" & sourceName & "；区域=" & sourceRange.Address & vbCrLf & _
        "输出模式=" & outputMode & vbCrLf & _
        "文本列=" & sourceField & vbCrLf & _
        "规则=" & BuildRuleSummary(catNames, catRules, catCount) & vbCrLf & _
        "默认类别=" & defaultOther & vbCrLf & _
        "数据行=" & CStr(dataRows) & "；预览行=" & CStr(previewRows) & vbCrLf & _
        "默认写新表，不覆盖原列；sourceUnchanged=true" & vbCrLf & _
        "---- 预览样本 ----"

    For i = 1 To previewRows
        rowIndex = i + 1
        cellText = CStr(sourceRange.Cells(rowIndex, sourceCol).Text)
        label = ClassifyText(cellText, catNames, catRules, catCount, defaultOther, matchedRule)
        If label = defaultOther And matchedRule = "" Then
            otherCount = otherCount + 1
        Else
            Dim ci
            ci = FindCategoryIndex(catNames, catCount, label)
            If ci > 0 Then hitCounts(ci) = hitCounts(ci) + 1
        End If
        If i <= 8 Then
            previewText = previewText & vbCrLf & "R" & CStr(rowIndex) & ": " & Left(cellText, 40) & " => " & label
            If Len(matchedRule) > 0 Then previewText = previewText & "（" & matchedRule & "）"
        End If
    Next
    If previewRows > 8 Then previewText = previewText & vbCrLf & "... 其余预览行确认后写入结果表"
    previewText = previewText & vbCrLf & "预览命中：" & BuildCategoryHitSummary(catNames, hitCounts, catCount, defaultOther, otherCount)
    previewText = previewText & vbCrLf & "AIGEN验收骨架：可用 =AIGEN(""按业务口径分类到指定类别"") 与规则结果对照；不自动改源表。"

    Dim planId, planPreview
    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_classify_column_preview", "{""source"":""" & EscapeJson(sourceRange.Address) & """,""field"":""" & EscapeJson(sourceField) & """,""rules"":""" & EscapeJson(BuildRuleSummary(catNames, catRules, catCount)) & """,""previewRows"":" & CStr(previewRows) & "}", "office.excel.classifyColumnPreview"
    planPreview = SafePreviewWritePlan()
    SafeRollbackWritePlan planId

    If Not SafeConfirmStep(previewText & vbCrLf & vbCrLf & "确认按规则全量分类到新表？", "excel.classify_column_preview") Then
        HostExcelClassifyColumnPreviewPack = FailureJson("E_CONFIRM_REQUIRED", "用户取消或未确认，未修改工作簿")
        Exit Function
    End If

    Dim createdSheets(), createdCount, resultSheet, summarySheet, outRows
    Dim fullHits(), fullOther
    createdCount = 0
    ReDim createdSheets(4)
    ReDim fullHits(catCount)
    For i = 1 To catCount
        fullHits(i) = 0
    Next
    fullOther = 0

    Set resultSheet = CreateOutputSheet(appObj, sourceName, "列式分类结果", outputMode)
    If resultSheet Is Nothing Then
        HostExcelClassifyColumnPreviewPack = FailureJson("E_OUTPUT_SHEET", "无法创建列式分类结果工作表")
        Exit Function
    End If
    createdCount = createdCount + 1: Set createdSheets(createdCount) = resultSheet

    If Not WriteClassifyResultSheet(sourceRange, sourceCol, catNames, catRules, catCount, defaultOther, resultSheet, outRows, fullHits, fullOther) Then
        RollbackCreatedSheets appObj, createdSheets, createdCount
        HostExcelClassifyColumnPreviewPack = FailureJson("E_CLASSIFY_WRITE", "列式分类结果写入失败，已删除未完成输出")
        Exit Function
    End If

    Set summarySheet = CreateOutputSheet(appObj, sourceName, "处理摘要", outputMode)
    If summarySheet Is Nothing Then
        RollbackCreatedSheets appObj, createdSheets, createdCount
        HostExcelClassifyColumnPreviewPack = FailureJson("E_OUTPUT_SHEET", "无法创建处理摘要工作表")
        Exit Function
    End If
    createdCount = createdCount + 1: Set createdSheets(createdCount) = summarySheet

    Dim summaryLines, impactJson, summary, aigenHint
    aigenHint = "AIGEN验收骨架：抽20行对比规则分类与业务口径；差异修正规则词表后再全量。"
    summaryLines = "场景名=列式分类预览场景包" & vbCrLf & _
        "命令ID=excel.classify_column_preview" & vbCrLf & _
        "源表名称=" & sourceName & vbCrLf & _
        "源区域=" & sourceRange.Address & vbCrLf & _
        "输出模式=" & outputMode & vbCrLf & _
        "输出位置=" & resultSheet.Name & "," & summarySheet.Name & vbCrLf & _
        "处理前行数=" & CStr(sourceRange.Rows.Count) & vbCrLf & _
        "处理后行数=" & CStr(outRows) & vbCrLf & _
        "参数=sourceField=" & sourceField & ";rules=" & BuildRuleSummary(catNames, catRules, catCount) & ";default=" & defaultOther & vbCrLf & _
        "预览行=" & CStr(previewRows) & vbCrLf & _
        "全量命中=" & BuildCategoryHitSummary(catNames, fullHits, catCount, defaultOther, fullOther) & vbCrLf & _
        "AIGEN提示=" & aigenHint & vbCrLf & _
        "执行时间=" & Now & vbCrLf & _
        "结果=成功"
    impactJson = SafeHostText("ExcelWriteImpactSummary", summaryLines, summarySheet.Name)
    If Not ExtractJsonBoolean(impactJson, "ok") Then WriteSummaryFallback summarySheet, summaryLines

    summary = "Excel 列式分类完成：结果=" & resultSheet.Name & "；数据行=" & CStr(outRows) & "；" & BuildCategoryHitSummary(catNames, fullHits, catCount, defaultOther, fullOther) & "；sourceUnchanged=true"
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostExcelClassifyColumnPreviewPack = "{""ok"":true,""sourceUnchanged"":true,""commandId"":""excel.classify_column_preview"",""resultSheet"":""" & EscapeJson(resultSheet.Name) & """,""summarySheet"":""" & EscapeJson(summarySheet.Name) & """,""rows"":" & CStr(outRows) & ",""writePlanPreview"":""" & EscapeJson(planPreview) & """,""message"":""" & EscapeJson(summary) & """}"
End Function

Function ParseClassifyRules(ruleText, ByRef catNames(), ByRef catRules())
    Dim text, parts, i, n, item, colonPos, namePart, rulePart
    text = Replace(Replace(CStr(ruleText), "；", ";"), vbLf, ";")
    text = Replace(text, vbCr, ";")
    parts = Split(text, ";")
    n = 0
    ReDim catNames(UBound(parts) + 1)
    ReDim catRules(UBound(parts) + 1)
    For i = 0 To UBound(parts)
        item = Trim(CStr(parts(i)))
        If Len(item) > 0 Then
            colonPos = InStr(item, ":")
            If colonPos = 0 Then colonPos = InStr(item, "：")
            If colonPos > 1 Then
                namePart = Trim(Left(item, colonPos - 1))
                rulePart = Trim(Mid(item, colonPos + 1))
                rulePart = Replace(Replace(rulePart, "|", "|"), "，", "|")
                rulePart = Replace(rulePart, ",", "|")
                If Len(namePart) > 0 And Len(rulePart) > 0 Then
                    n = n + 1
                    catNames(n) = namePart
                    catRules(n) = rulePart
                End If
            End If
        End If
    Next
    ParseClassifyRules = n
End Function

Function BuildRuleSummary(ByRef catNames(), ByRef catRules(), catCount)
    Dim i, buf
    buf = ""
    For i = 1 To catCount
        If Len(buf) > 0 Then buf = buf & "；"
        buf = buf & catNames(i) & ":" & catRules(i)
    Next
    BuildRuleSummary = buf
End Function

Function FindCategoryIndex(ByRef catNames(), catCount, label)
    Dim i
    FindCategoryIndex = 0
    For i = 1 To catCount
        If StrComp(CStr(catNames(i)), CStr(label), vbTextCompare) = 0 Then
            FindCategoryIndex = i
            Exit Function
        End If
    Next
End Function

Function BuildCategoryHitSummary(ByRef catNames(), ByRef hits(), catCount, defaultOther, otherCount)
    Dim i, buf
    buf = ""
    For i = 1 To catCount
        If Len(buf) > 0 Then buf = buf & "；"
        buf = buf & catNames(i) & "=" & CStr(hits(i))
    Next
    If Len(buf) > 0 Then buf = buf & "；"
    buf = buf & defaultOther & "=" & CStr(otherCount)
    BuildCategoryHitSummary = buf
End Function

Function ClassifyText(textValue, ByRef catNames(), ByRef catRules(), catCount, defaultOther, ByRef matchedRule)
    Dim i, kws, k, kw, t
    t = CStr(textValue)
    matchedRule = ""
    ClassifyText = defaultOther
    For i = 1 To catCount
        kws = Split(CStr(catRules(i)), "|")
        For k = 0 To UBound(kws)
            kw = Trim(CStr(kws(k)))
            If Len(kw) > 0 Then
                If InStr(1, t, kw, vbTextCompare) > 0 Then
                    ClassifyText = catNames(i)
                    matchedRule = kw
                    Exit Function
                End If
            End If
        Next
    Next
End Function

Function WriteClassifyResultSheet(sourceRange, sourceCol, ByRef catNames(), ByRef catRules(), catCount, defaultOther, sheetObj, ByRef outRows, ByRef hits(), ByRef otherCount)
    On Error Resume Next
    Dim c, r, srcCols, cellText, label, matchedRule, ci
    WriteClassifyResultSheet = False
    outRows = 0
    otherCount = 0
    srcCols = sourceRange.Columns.Count

    sheetObj.Cells(1, 1).Value = "_来源行"
    For c = 1 To srcCols
        sheetObj.Cells(1, c + 1).Value = sourceRange.Cells(1, c).Text
    Next
    sheetObj.Cells(1, srcCols + 2).Value = "分类结果"
    sheetObj.Cells(1, srcCols + 3).Value = "命中关键词"
    sheetObj.Cells(1, srcCols + 4).Value = "_分类说明"

    For r = 2 To sourceRange.Rows.Count
        outRows = outRows + 1
        sheetObj.Cells(outRows + 1, 1).Value = sourceRange.Row + r - 1
        For c = 1 To srcCols
            sheetObj.Cells(outRows + 1, c + 1).Value = sourceRange.Cells(r, c).Value
        Next
        cellText = CStr(sourceRange.Cells(r, sourceCol).Text)
        label = ClassifyText(cellText, catNames, catRules, catCount, defaultOther, matchedRule)
        sheetObj.Cells(outRows + 1, srcCols + 2).Value = label
        sheetObj.Cells(outRows + 1, srcCols + 3).Value = matchedRule
        If Len(matchedRule) = 0 Then
            sheetObj.Cells(outRows + 1, srcCols + 4).Value = "未命中规则，使用默认类别"
            otherCount = otherCount + 1
        Else
            sheetObj.Cells(outRows + 1, srcCols + 4).Value = "规则命中"
            ci = FindCategoryIndex(catNames, catCount, label)
            If ci > 0 Then hits(ci) = hits(ci) + 1
        End If
        If Err.Number <> 0 Then
            Err.Clear
            Exit Function
        End If
    Next

    sheetObj.Rows(1).Font.Bold = True
    sheetObj.Columns.AutoFit
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    WriteClassifyResultSheet = True
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
