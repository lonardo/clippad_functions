
' 函数名: HostExcelValidateReportPack
' 描述: 数据质检场景包：按必填/主键唯一/电话/日期/错误值规则预览问题清单，确认后只写问题明细与质检报告，不改源表
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
    Main = HostExcelValidateReportPack(appObj)
End Function

Function HostExcelValidateReportPack(appObj)
    On Error Resume Next
    Dim sourceRange, sourceSheet, sourceName
    Set sourceRange = appObj.Selection.CurrentRegion
    If Err.Number <> 0 Or TypeName(sourceRange) = "Empty" Or TypeName(sourceRange) = "Nothing" Then
        Err.Clear
        HostExcelValidateReportPack = FailureJson("E_NO_RANGE", "请先选中含表头的数据区域")
        Exit Function
    End If
    If sourceRange.Rows.Count < 2 Or sourceRange.Columns.Count < 1 Then
        HostExcelValidateReportPack = FailureJson("E_RANGE_TOO_SMALL", "当前区域至少需要一行表头和一行数据")
        Exit Function
    End If

    Set sourceSheet = sourceRange.Worksheet
    sourceName = sourceSheet.Name

    Dim outputMode, requiredText, keyUniqueText, phoneText, dateText, errorText
    outputMode = NormalizeOutputMode(SafePrompt("输出模式：NewSheet 或 NewWorkbook（默认 NewSheet）", "NewSheet"))
    requiredText = Trim(SafePrompt("必填字段（逗号分隔，可空）", CStr(sourceRange.Cells(1, 1).Text)))
    keyUniqueText = Trim(SafePrompt("主键唯一字段（逗号分隔，可空）", ""))
    phoneText = Trim(SafePrompt("电话校验字段（逗号分隔，可空）", ""))
    dateText = Trim(SafePrompt("日期校验字段（逗号分隔，可空）", ""))
    errorText = Trim(SafePrompt("是否检查错误值（是/否，默认是）", "是"))

    Dim rulesJson, validateJson, issueCount, truncated, issueRows(), issueTypes(), issueFields(), issueValues(), issueCountLocal
    rulesJson = BuildValidateRulesJson(requiredText, keyUniqueText, phoneText, dateText, errorText)
    If rulesJson = "{}" Then
        HostExcelValidateReportPack = FailureJson("E_NO_RULES", "请至少指定一类质检规则")
        Exit Function
    End If

    Dim headerMap, sampleJson
    headerMap = SafeHostText("ExcelGetHeaderMap", "")
    sampleJson = SafeHostText("ExcelPreviewRowSamples", 3, "")
    validateJson = SafeHostText("ExcelValidatePlan", rulesJson)
    issueCountLocal = 0
    If ExtractJsonBoolean(validateJson, "ok") Then
        issueCount = CLng(ExtractJsonNumber(validateJson, "issueCount"))
        If issueCount <= 0 Then issueCount = CLng(ExtractJsonNumber(validateJson, "issuesCount"))
        truncated = ExtractJsonBoolean(validateJson, "truncated")
        ParseIssuesFromHost validateJson, issueRows, issueTypes, issueFields, issueValues, issueCountLocal
        If issueCount <= 0 Then issueCount = issueCountLocal
    Else
        BuildLocalValidateIssues sourceRange, requiredText, keyUniqueText, phoneText, dateText, errorText, issueRows, issueTypes, issueFields, issueValues, issueCountLocal
        issueCount = issueCountLocal
        truncated = False
        validateJson = "{""ok"":true,""issueCount"":" & CStr(issueCount) & ",""truncated"":false}"
    End If

    Dim previewText, planId, planPreview
    previewText = "Excel 数据质检场景包预览（尚未写入）" & vbCrLf & _
        "命令ID：excel.validate_report" & vbCrLf & _
        "源表：" & sourceName & " 区域=" & sourceRange.Address & vbCrLf & _
        "输出模式=" & outputMode & "；源数据行=" & CStr(sourceRange.Rows.Count - 1) & vbCrLf & _
        "规则：必填=" & requiredText & "；主键唯一=" & keyUniqueText & "；电话=" & phoneText & "；日期=" & dateText & "；错误值=" & errorText & vbCrLf & _
        "问题数=" & CStr(issueCount)
    If truncated Then previewText = previewText & "（结果已截断）"
    previewText = previewText & vbCrLf & "将生成：问题明细 + 质检报告；sourceUnchanged=true；不修改源表样式或数据"

    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_validate_report", "{""source"":""" & EscapeJson(sourceRange.Address) & """,""issueCount"":" & CStr(issueCount) & ",""rules"":""" & EscapeJson(rulesJson) & """}", "office.excel.validatePack"
    planPreview = SafePreviewWritePlan()
    SafeRollbackWritePlan planId

    If Not SafeConfirmStep(previewText & vbCrLf & vbCrLf & "确认生成问题明细与质检报告？", "excel.validate_report") Then
        HostExcelValidateReportPack = FailureJson("E_CONFIRM_REQUIRED", "用户取消或未确认，未修改工作簿")
        Exit Function
    End If

    Dim detailSheet, reportSheet, createdSheets(), createdCount
    createdCount = 0
    ReDim createdSheets(4)
    Set detailSheet = CreateOutputSheet(appObj, sourceName, "问题明细", outputMode)
    If detailSheet Is Nothing Then
        HostExcelValidateReportPack = FailureJson("E_OUTPUT_SHEET", "无法创建问题明细工作表")
        Exit Function
    End If
    createdCount = createdCount + 1: Set createdSheets(createdCount) = detailSheet

    If Not WriteIssueDetailSheet(detailSheet, issueRows, issueTypes, issueFields, issueValues, issueCountLocal) Then
        RollbackCreatedSheets appObj, createdSheets, createdCount
        HostExcelValidateReportPack = FailureJson("E_VALIDATE_WRITE", "问题明细写入失败，已删除未完成输出")
        Exit Function
    End If

    Set reportSheet = CreateOutputSheet(appObj, sourceName, "质检报告", outputMode)
    If reportSheet Is Nothing Then
        RollbackCreatedSheets appObj, createdSheets, createdCount
        HostExcelValidateReportPack = FailureJson("E_OUTPUT_SHEET", "无法创建质检报告工作表")
        Exit Function
    End If
    createdCount = createdCount + 1: Set createdSheets(createdCount) = reportSheet

    Dim summaryLines, impactJson, summary
    summaryLines = "场景名=数据质检场景包" & vbCrLf & _
        "命令ID=excel.validate_report" & vbCrLf & _
        "源表名称=" & sourceName & vbCrLf & _
        "源区域=" & sourceRange.Address & vbCrLf & _
        "输出模式=" & outputMode & vbCrLf & _
        "输出位置=" & detailSheet.Name & "," & reportSheet.Name & vbCrLf & _
        "处理前行数=" & CStr(sourceRange.Rows.Count - 1) & vbCrLf & _
        "处理后行数=" & CStr(sourceRange.Rows.Count - 1) & vbCrLf & _
        "参数=" & rulesJson & vbCrLf & _
        "异常/问题数=" & CStr(issueCount) & vbCrLf & _
        "执行时间=" & Now & vbCrLf & _
        "结果=成功"
    impactJson = SafeHostText("ExcelWriteImpactSummary", summaryLines, reportSheet.Name)
    If Not ExtractJsonBoolean(impactJson, "ok") Then WriteSummaryFallback reportSheet, summaryLines

    summary = "Excel 数据质检完成：问题明细=" & detailSheet.Name & "；质检报告=" & reportSheet.Name & _
        "；问题数=" & CStr(issueCount) & "；sourceUnchanged=true"
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostExcelValidateReportPack = "{""ok"":true,""sourceUnchanged"":true,""commandId"":""excel.validate_report"",""detailSheet"":""" & EscapeJson(detailSheet.Name) & """,""reportSheet"":""" & EscapeJson(reportSheet.Name) & """,""issueCount"":" & CStr(issueCount) & ",""writePlanPreview"":""" & EscapeJson(planPreview) & """,""message"":""" & EscapeJson(summary) & """}"
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

Function BuildValidateRulesJson(requiredText, keyUniqueText, phoneText, dateText, errorText)
    Dim parts, errFlag, t
    parts = ""
    If Len(Trim(requiredText)) > 0 Then parts = parts & """required"":""" & EscapeJson(requiredText) & """"
    If Len(Trim(keyUniqueText)) > 0 Then
        If Len(parts) > 0 Then parts = parts & ","
        parts = parts & """keyUnique"":""" & EscapeJson(keyUniqueText) & """"
    End If
    If Len(Trim(phoneText)) > 0 Then
        If Len(parts) > 0 Then parts = parts & ","
        parts = parts & """phone"":""" & EscapeJson(phoneText) & """"
    End If
    If Len(Trim(dateText)) > 0 Then
        If Len(parts) > 0 Then parts = parts & ","
        parts = parts & """date"":""" & EscapeJson(dateText) & """"
    End If
    t = LCase(Trim(CStr(errorText)))
    errFlag = True
    If t = "否" Or t = "no" Or t = "false" Or t = "0" Then errFlag = False
    If errFlag Then
        If Len(parts) > 0 Then parts = parts & ","
        parts = parts & """error"":true"
    End If
    If Len(parts) = 0 Then
        BuildValidateRulesJson = "{}"
    Else
        BuildValidateRulesJson = "{" & parts & "}"
    End If
End Function

Sub ParseIssuesFromHost(validateJson, ByRef issueRows, ByRef issueTypes, ByRef issueFields, ByRef issueValues, ByRef issueCountLocal)
    Dim marker, startPos, endPos, body, parts, i, item, rowNo, ruleName, fieldName, msg, valueText
    issueCountLocal = 0
    ReDim issueRows(0)
    ReDim issueTypes(0)
    ReDim issueFields(0)
    ReDim issueValues(0)
    marker = """issues"""
    startPos = InStr(1, CStr(validateJson), marker, vbTextCompare)
    If startPos <= 0 Then Exit Sub
    startPos = InStr(startPos, CStr(validateJson), "[")
    If startPos <= 0 Then Exit Sub
    endPos = InStr(startPos + 1, CStr(validateJson), "]")
    If endPos <= startPos Then Exit Sub
    body = Mid(CStr(validateJson), startPos + 1, endPos - startPos - 1)
    If Len(Trim(body)) = 0 Then Exit Sub
    ' naive split by },{
    body = Replace(body, "}{", "}|{")
    body = Replace(body, "},{", "}|{")
    parts = Split(body, "|")
    ReDim issueRows(UBound(parts))
    ReDim issueTypes(UBound(parts))
    ReDim issueFields(UBound(parts))
    ReDim issueValues(UBound(parts))
    For i = 0 To UBound(parts)
        item = parts(i)
        rowNo = CLng(ExtractJsonNumber(item, "row"))
        ruleName = ExtractJsonString(item, "rule")
        If Len(ruleName) = 0 Then ruleName = ExtractJsonString(item, "type")
        fieldName = ExtractJsonString(item, "column")
        If Len(fieldName) = 0 Then fieldName = ExtractJsonString(item, "field")
        msg = ExtractJsonString(item, "message")
        valueText = ExtractJsonString(item, "value")
        If rowNo > 0 Or Len(ruleName) > 0 Or Len(fieldName) > 0 Then
            issueRows(issueCountLocal) = rowNo
            issueTypes(issueCountLocal) = ruleName
            If Len(msg) > 0 And Len(ruleName) > 0 Then
                issueTypes(issueCountLocal) = ruleName
            ElseIf Len(msg) > 0 Then
                issueTypes(issueCountLocal) = msg
            End If
            issueFields(issueCountLocal) = fieldName
            issueValues(issueCountLocal) = valueText
            issueCountLocal = issueCountLocal + 1
        End If
    Next
    If issueCountLocal = 0 Then Exit Sub
    ReDim Preserve issueRows(issueCountLocal - 1)
    ReDim Preserve issueTypes(issueCountLocal - 1)
    ReDim Preserve issueFields(issueCountLocal - 1)
    ReDim Preserve issueValues(issueCountLocal - 1)
End Sub

Sub BuildLocalValidateIssues(sourceRange, requiredText, keyUniqueText, phoneText, dateText, errorText, ByRef issueRows, ByRef issueTypes, ByRef issueFields, ByRef issueValues, ByRef issueCountLocal)
    Dim requiredCols, keyCols, phoneCols, dateCols, checkError
    Dim r, c, i, j, text, digits, normalized, t
    Dim keyMap(), keyCount, keyValues(), keyFirstRow()
    issueCountLocal = 0
    ReDim issueRows(sourceRange.Rows.Count * 4)
    ReDim issueTypes(sourceRange.Rows.Count * 4)
    ReDim issueFields(sourceRange.Rows.Count * 4)
    ReDim issueValues(sourceRange.Rows.Count * 4)

    requiredCols = ResolveHeaderColumns(sourceRange, SplitCsvFields(requiredText))
    keyCols = ResolveHeaderColumns(sourceRange, SplitCsvFields(keyUniqueText))
    phoneCols = ResolveHeaderColumns(sourceRange, SplitCsvFields(phoneText))
    dateCols = ResolveHeaderColumns(sourceRange, SplitCsvFields(dateText))
    t = LCase(Trim(CStr(errorText)))
    checkError = True
    If t = "否" Or t = "no" Or t = "false" Or t = "0" Then checkError = False

    ' required / phone / date / error
    For r = 2 To sourceRange.Rows.Count
        If IsArray(requiredCols) Then
            On Error Resume Next
            If UBound(requiredCols) >= 0 Then
                For i = 0 To UBound(requiredCols)
                    c = requiredCols(i)
                    text = NormalizeCellText(sourceRange.Cells(r, c).Text)
                    If Len(text) = 0 Then
                        AppendIssue issueRows, issueTypes, issueFields, issueValues, issueCountLocal, sourceRange.Row + r - 1, "required", CStr(sourceRange.Cells(1, c).Text), ""
                    End If
                Next
            End If
            Err.Clear
        End If

        If IsArray(phoneCols) Then
            On Error Resume Next
            If UBound(phoneCols) >= 0 Then
                For i = 0 To UBound(phoneCols)
                    c = phoneCols(i)
                    text = NormalizeCellText(sourceRange.Cells(r, c).Text)
                    If Len(text) > 0 Then
                        digits = NormalizePhoneDigits(text)
                        If Not (Len(digits) = 11 And Left(digits, 1) = "1") And Not (Len(digits) >= 7 And Len(digits) <= 12) Then
                            AppendIssue issueRows, issueTypes, issueFields, issueValues, issueCountLocal, sourceRange.Row + r - 1, "phone", CStr(sourceRange.Cells(1, c).Text), text
                        End If
                    End If
                Next
            End If
            Err.Clear
        End If

        If IsArray(dateCols) Then
            On Error Resume Next
            If UBound(dateCols) >= 0 Then
                For i = 0 To UBound(dateCols)
                    c = dateCols(i)
                    text = NormalizeCellText(sourceRange.Cells(r, c).Text)
                    If Len(text) > 0 Then
                        If Not TryNormalizeDateText(text, normalized) Then
                            AppendIssue issueRows, issueTypes, issueFields, issueValues, issueCountLocal, sourceRange.Row + r - 1, "date", CStr(sourceRange.Cells(1, c).Text), text
                        End If
                    End If
                Next
            End If
            Err.Clear
        End If

        If checkError Then
            For c = 1 To sourceRange.Columns.Count
                text = CStr(sourceRange.Cells(r, c).Text)
                If Left(text, 1) = "#" Then
                    AppendIssue issueRows, issueTypes, issueFields, issueValues, issueCountLocal, sourceRange.Row + r - 1, "error", CStr(sourceRange.Cells(1, c).Text), text
                End If
            Next
        End If
    Next

    ' key unique
    If IsArray(keyCols) Then
        On Error Resume Next
        If UBound(keyCols) >= 0 Then
            keyCount = 0
            ReDim keyValues(sourceRange.Rows.Count)
            ReDim keyFirstRow(sourceRange.Rows.Count)
            For r = 2 To sourceRange.Rows.Count
                text = RowKeyFromColumns(sourceRange, r, keyCols)
                If Len(text) = 0 Then
                    AppendIssue issueRows, issueTypes, issueFields, issueValues, issueCountLocal, sourceRange.Row + r - 1, "keyUnique", JoinFieldNames(sourceRange, keyCols), ""
                Else
                    Dim found
                    found = 0
                    For j = 1 To keyCount
                        If keyValues(j) = text Then
                            found = j
                            Exit For
                        End If
                    Next
                    If found = 0 Then
                        keyCount = keyCount + 1
                        keyValues(keyCount) = text
                        keyFirstRow(keyCount) = sourceRange.Row + r - 1
                    Else
                        AppendIssue issueRows, issueTypes, issueFields, issueValues, issueCountLocal, sourceRange.Row + r - 1, "keyUnique", JoinFieldNames(sourceRange, keyCols), text
                    End If
                End If
            Next
        End If
        Err.Clear
    End If

    If issueCountLocal = 0 Then
        ReDim issueRows(0)
        ReDim issueTypes(0)
        ReDim issueFields(0)
        ReDim issueValues(0)
    Else
        ReDim Preserve issueRows(issueCountLocal - 1)
        ReDim Preserve issueTypes(issueCountLocal - 1)
        ReDim Preserve issueFields(issueCountLocal - 1)
        ReDim Preserve issueValues(issueCountLocal - 1)
    End If
End Sub

Sub AppendIssue(ByRef issueRows, ByRef issueTypes, ByRef issueFields, ByRef issueValues, ByRef issueCountLocal, rowNo, ruleName, fieldName, valueText)
    If issueCountLocal > UBound(issueRows) Then
        ReDim Preserve issueRows(issueCountLocal + 64)
        ReDim Preserve issueTypes(issueCountLocal + 64)
        ReDim Preserve issueFields(issueCountLocal + 64)
        ReDim Preserve issueValues(issueCountLocal + 64)
    End If
    issueRows(issueCountLocal) = rowNo
    issueTypes(issueCountLocal) = ruleName
    issueFields(issueCountLocal) = fieldName
    issueValues(issueCountLocal) = valueText
    issueCountLocal = issueCountLocal + 1
End Sub

Function JoinFieldNames(sourceRange, cols)
    Dim i, parts
    parts = ""
    If Not IsArray(cols) Then
        JoinFieldNames = ""
        Exit Function
    End If
    On Error Resume Next
    For i = 0 To UBound(cols)
        If i > 0 Then parts = parts & ","
        parts = parts & CStr(sourceRange.Cells(1, cols(i)).Text)
    Next
    JoinFieldNames = parts
End Function

Function WriteIssueDetailSheet(sheetObj, issueRows, issueTypes, issueFields, issueValues, issueCountLocal)
    On Error Resume Next
    Dim i
    WriteIssueDetailSheet = False
    sheetObj.Cells(1, 1).Value = "源行号"
    sheetObj.Cells(1, 2).Value = "问题类型"
    sheetObj.Cells(1, 3).Value = "字段"
    sheetObj.Cells(1, 4).Value = "原值"
    sheetObj.Cells(1, 5).Value = "处理状态"
    sheetObj.Cells(1, 6).Value = "备注"
    If issueCountLocal > 0 Then
        For i = 0 To issueCountLocal - 1
            sheetObj.Cells(i + 2, 1).Value = issueRows(i)
            sheetObj.Cells(i + 2, 2).Value = issueTypes(i)
            sheetObj.Cells(i + 2, 3).Value = issueFields(i)
            sheetObj.Cells(i + 2, 4).Value = issueValues(i)
            sheetObj.Cells(i + 2, 5).Value = "待处理"
            sheetObj.Cells(i + 2, 6).Value = ""
        Next
    End If
    sheetObj.Rows(1).Font.Bold = True
    sheetObj.Columns.AutoFit
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    WriteIssueDetailSheet = True
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
