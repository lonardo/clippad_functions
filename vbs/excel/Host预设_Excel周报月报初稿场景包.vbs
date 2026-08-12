' 函数名: HostExcelReportDraftPack
' 描述: 周报/月报初稿场景包：按当前数据区自动汇总事实指标、异常待办候选，并生成含重点说明/下期计划占位的初稿与处理摘要，不改源表
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
    Main = HostExcelReportDraftPack(appObj)
End Function

Function HostExcelReportDraftPack(appObj)
    On Error Resume Next
    Dim sourceRange, sourceSheet, sourceName, outputMode
    Set sourceRange = appObj.Selection.CurrentRegion
    If Err.Number <> 0 Or TypeName(sourceRange) = "Empty" Or TypeName(sourceRange) = "Nothing" Then
        Err.Clear
        HostExcelReportDraftPack = FailureJson("E_NO_RANGE", "请先选中带表头的数据区域")
        Exit Function
    End If
    If sourceRange.Rows.Count < 2 Then
        HostExcelReportDraftPack = FailureJson("E_RANGE_TOO_SMALL", "当前区域至少需要一行表头和一行数据")
        Exit Function
    End If
    Set sourceSheet = sourceRange.Worksheet
    sourceName = sourceSheet.Name

    Dim reportType, reportTitle, amountField, statusField, ownerField, dateField
    reportType = NormalizeReportType(SafePrompt("报表类型：周报 或 月报（默认 周报）", "周报"))
    reportTitle = Trim(SafePrompt("报表标题", sourceName & "_" & reportType & "初稿"))
    If Len(reportTitle) = 0 Then reportTitle = sourceName & "_" & reportType & "初稿"
    amountField = Trim(SafePrompt("金额字段（可空）", GuessHeaderByKeywords(sourceRange, "金额,销售额,回款,应收,费用,数量")))
    statusField = Trim(SafePrompt("状态字段（可空）", GuessHeaderByKeywords(sourceRange, "状态,进度,完成情况")))
    ownerField = Trim(SafePrompt("负责人字段（可空）", GuessHeaderByKeywords(sourceRange, "负责人,处理人,跟进人,责任人,销售")))
    dateField = Trim(SafePrompt("日期字段（可空）", GuessHeaderByKeywords(sourceRange, "日期,截止日期,到期日,业务日期,完成日期")))
    outputMode = NormalizeOutputMode(SafePrompt("输出模式：NewSheet 或 NewWorkbook（默认 NewSheet）", "NewSheet"))

    Dim amountCol, statusCol, ownerCol, dateCol, dataRows
    amountCol = 0: statusCol = 0: ownerCol = 0: dateCol = 0
    If Len(amountField) > 0 Then amountCol = FindHeaderColumn(sourceRange, amountField)
    If Len(statusField) > 0 Then statusCol = FindHeaderColumn(sourceRange, statusField)
    If Len(ownerField) > 0 Then ownerCol = FindHeaderColumn(sourceRange, ownerField)
    If Len(dateField) > 0 Then dateCol = FindHeaderColumn(sourceRange, dateField)
    If Len(amountField) > 0 And amountCol <= 0 Then
        HostExcelReportDraftPack = FailureJson("E_FIELD_NOT_FOUND", "未找到金额字段：" & amountField)
        Exit Function
    End If
    If Len(statusField) > 0 And statusCol <= 0 Then
        HostExcelReportDraftPack = FailureJson("E_FIELD_NOT_FOUND", "未找到状态字段：" & statusField)
        Exit Function
    End If
    If Len(ownerField) > 0 And ownerCol <= 0 Then
        HostExcelReportDraftPack = FailureJson("E_FIELD_NOT_FOUND", "未找到负责人字段：" & ownerField)
        Exit Function
    End If
    If Len(dateField) > 0 And dateCol <= 0 Then
        HostExcelReportDraftPack = FailureJson("E_FIELD_NOT_FOUND", "未找到日期字段：" & dateField)
        Exit Function
    End If

    dataRows = sourceRange.Rows.Count - 1
    Dim amountSum, amountCount, amountAvg, openCount, doneCount, overdueCount
    Dim ownerNames(), ownerCounts(), ownerAmount(), ownerN
    Dim anomalyLines, anomalyCount, sampleOwners
    Dim r, text, val, ownerName, i, found, normalized
    amountSum = 0: amountCount = 0: amountAvg = 0
    openCount = 0: doneCount = 0: overdueCount = 0
    ownerN = 0
    anomalyLines = "": anomalyCount = 0: sampleOwners = ""

    For r = 2 To sourceRange.Rows.Count
        If amountCol > 0 Then
            text = NormalizeCellText(sourceRange.Cells(r, amountCol).Text)
            If Len(text) > 0 And IsNumeric(text) Then
                val = CDbl(text)
                amountSum = amountSum + val
                amountCount = amountCount + 1
            End If
        End If

        Dim statusText, isDone, isOpen
        statusText = ""
        isDone = False
        isOpen = False
        If statusCol > 0 Then
            statusText = NormalizeCellText(sourceRange.Cells(r, statusCol).Text)
            If InStr(1, statusText, "完成", vbTextCompare) > 0 Or InStr(1, statusText, "关闭", vbTextCompare) > 0 Or InStr(1, statusText, "已回款", vbTextCompare) > 0 Then
                isDone = True
                doneCount = doneCount + 1
            Else
                isOpen = True
                openCount = openCount + 1
            End If
        End If

        Dim isOverdue
        isOverdue = False
        If dateCol > 0 Then
            text = NormalizeCellText(sourceRange.Cells(r, dateCol).Text)
            If TryNormalizeDateText(text, normalized) Then
                If CDate(normalized) < Date Then
                    If statusCol <= 0 Or isOpen Then
                        isOverdue = True
                        overdueCount = overdueCount + 1
                    End If
                End If
            End If
        End If

        If ownerCol > 0 Then
            ownerName = NormalizeCellText(sourceRange.Cells(r, ownerCol).Text)
            If Len(ownerName) = 0 Then ownerName = "(空)"
            found = False
            For i = 0 To ownerN - 1
                If StrComp(ownerNames(i), ownerName, vbTextCompare) = 0 Then
                    ownerCounts(i) = ownerCounts(i) + 1
                    If amountCol > 0 Then
                        text = NormalizeCellText(sourceRange.Cells(r, amountCol).Text)
                        If Len(text) > 0 And IsNumeric(text) Then ownerAmount(i) = ownerAmount(i) + CDbl(text)
                    End If
                    found = True
                    Exit For
                End If
            Next
            If Not found Then
                ReDim Preserve ownerNames(ownerN)
                ReDim Preserve ownerCounts(ownerN)
                ReDim Preserve ownerAmount(ownerN)
                ownerNames(ownerN) = ownerName
                ownerCounts(ownerN) = 1
                ownerAmount(ownerN) = 0
                If amountCol > 0 Then
                    text = NormalizeCellText(sourceRange.Cells(r, amountCol).Text)
                    If Len(text) > 0 And IsNumeric(text) Then ownerAmount(ownerN) = CDbl(text)
                End If
                ownerN = ownerN + 1
            End If
        End If

        If anomalyCount < 12 Then
            If isOverdue Or (statusCol > 0 And isOpen And InStr(1, statusText, "高", vbTextCompare) > 0) Or (statusCol > 0 And isOpen And InStr(1, statusText, "风险", vbTextCompare) > 0) Then
                If Len(anomalyLines) > 0 Then anomalyLines = anomalyLines & vbCrLf
                anomalyLines = anomalyLines & "行" & CStr(sourceRange.Row + r - 1)
                If ownerCol > 0 Then anomalyLines = anomalyLines & " / " & NormalizeCellText(sourceRange.Cells(r, ownerCol).Text)
                If statusCol > 0 Then anomalyLines = anomalyLines & " / 状态=" & statusText
                If dateCol > 0 Then anomalyLines = anomalyLines & " / 日期=" & NormalizeCellText(sourceRange.Cells(r, dateCol).Text)
                If amountCol > 0 Then anomalyLines = anomalyLines & " / 金额=" & NormalizeCellText(sourceRange.Cells(r, amountCol).Text)
                anomalyCount = anomalyCount + 1
            End If
        End If
    Next

    If amountCount > 0 Then amountAvg = amountSum / amountCount

    ' top owners sample
    Dim a, b, tmpName, tmpCount, tmpAmt
    If ownerN > 1 Then
        For a = 0 To ownerN - 2
            For b = a + 1 To ownerN - 1
                If ownerCounts(b) > ownerCounts(a) Then
                    tmpName = ownerNames(a): ownerNames(a) = ownerNames(b): ownerNames(b) = tmpName
                    tmpCount = ownerCounts(a): ownerCounts(a) = ownerCounts(b): ownerCounts(b) = tmpCount
                    tmpAmt = ownerAmount(a): ownerAmount(a) = ownerAmount(b): ownerAmount(b) = tmpAmt
                End If
            Next
        Next
    End If
    For i = 0 To ownerN - 1
        If i > 4 Then Exit For
        If Len(sampleOwners) > 0 Then sampleOwners = sampleOwners & "；"
        sampleOwners = sampleOwners & ownerNames(i) & "x" & CStr(ownerCounts(i))
        If amountCol > 0 Then sampleOwners = sampleOwners & "/金额" & CStr(ownerAmount(i))
    Next

    Dim previewText, planId, planPreview
    previewText = "Excel 周报月报初稿预览（尚未写入）" & vbCrLf & _
        "命令ID：excel.report_draft_pack" & vbCrLf & _
        "源表=" & sourceName & "；区域=" & sourceRange.Address & vbCrLf & _
        "类型=" & reportType & "；标题=" & reportTitle & vbCrLf & _
        "数据行=" & CStr(dataRows) & "；输出模式=" & outputMode & vbCrLf & _
        "字段：金额=" & amountField & "；状态=" & statusField & "；负责人=" & ownerField & "；日期=" & dateField & vbCrLf & _
        "指标：金额合计=" & CStr(amountSum) & "；未完成/进行中=" & CStr(openCount) & "；已完成=" & CStr(doneCount) & "；逾期候选=" & CStr(overdueCount) & vbCrLf & _
        "将生成：报表初稿 + 处理摘要；事实自动填入，结论留给人工；sourceUnchanged=true"

    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_report_draft_pack", "{""source"":""" & EscapeJson(sourceRange.Address) & """,""type"":""" & EscapeJson(reportType) & """,""rows"":" & CStr(dataRows) & ",""amountSum"":" & CStr(amountSum) & ",""overdue"":" & CStr(overdueCount) & "}", "office.excel.reportDraftPack"
    planPreview = SafePreviewWritePlan()
    SafeRollbackWritePlan planId

    If Not SafeConfirmStep(previewText & vbCrLf & vbCrLf & "确认生成周报/月报初稿？", "excel.report_draft_pack") Then
        HostExcelReportDraftPack = FailureJson("E_CONFIRM_REQUIRED", "用户取消或未确认，未修改工作簿")
        Exit Function
    End If

    Dim createdSheets(), createdCount, draftSheet, summarySheet
    createdCount = 0
    ReDim createdSheets(4)
    Set draftSheet = CreateOutputSheet(appObj, sourceName, reportType & "初稿", outputMode)
    If draftSheet Is Nothing Then
        HostExcelReportDraftPack = FailureJson("E_OUTPUT_SHEET", "无法创建报表初稿工作表")
        Exit Function
    End If
    createdCount = createdCount + 1: Set createdSheets(createdCount) = draftSheet

    Dim rowOut
    rowOut = 1
    draftSheet.Cells(rowOut, 1).Value = reportTitle
    draftSheet.Cells(rowOut, 1).Font.Bold = True
    draftSheet.Cells(rowOut, 1).Font.Size = 14
    rowOut = rowOut + 1
    draftSheet.Cells(rowOut, 1).Value = "报表类型"
    draftSheet.Cells(rowOut, 2).Value = reportType
    rowOut = rowOut + 1
    draftSheet.Cells(rowOut, 1).Value = "数据来源"
    draftSheet.Cells(rowOut, 2).Value = sourceName & " " & sourceRange.Address
    rowOut = rowOut + 1
    draftSheet.Cells(rowOut, 1).Value = "生成时间"
    draftSheet.Cells(rowOut, 2).Value = Now
    rowOut = rowOut + 2

    draftSheet.Cells(rowOut, 1).Value = "一、本期数据摘要"
    draftSheet.Cells(rowOut, 1).Font.Bold = True
    rowOut = rowOut + 1
    draftSheet.Cells(rowOut, 1).Value = "数据行数"
    draftSheet.Cells(rowOut, 2).Value = dataRows
    rowOut = rowOut + 1
    draftSheet.Cells(rowOut, 1).Value = "金额字段"
    draftSheet.Cells(rowOut, 2).Value = IIf(Len(amountField) > 0, amountField, "（未指定）")
    rowOut = rowOut + 1
    draftSheet.Cells(rowOut, 1).Value = "状态字段"
    draftSheet.Cells(rowOut, 2).Value = IIf(Len(statusField) > 0, statusField, "（未指定）")
    rowOut = rowOut + 1
    draftSheet.Cells(rowOut, 1).Value = "负责人字段"
    draftSheet.Cells(rowOut, 2).Value = IIf(Len(ownerField) > 0, ownerField, "（未指定）")
    rowOut = rowOut + 1
    draftSheet.Cells(rowOut, 1).Value = "日期字段"
    draftSheet.Cells(rowOut, 2).Value = IIf(Len(dateField) > 0, dateField, "（未指定）")
    rowOut = rowOut + 2

    draftSheet.Cells(rowOut, 1).Value = "二、指标"
    draftSheet.Cells(rowOut, 1).Font.Bold = True
    rowOut = rowOut + 1
    draftSheet.Cells(rowOut, 1).Value = "指标"
    draftSheet.Cells(rowOut, 2).Value = "数值"
    draftSheet.Cells(rowOut, 3).Value = "说明"
    draftSheet.Rows(rowOut).Font.Bold = True
    rowOut = rowOut + 1
    draftSheet.Cells(rowOut, 1).Value = "记录数"
    draftSheet.Cells(rowOut, 2).Value = dataRows
    draftSheet.Cells(rowOut, 3).Value = "本期数据行"
    rowOut = rowOut + 1
    draftSheet.Cells(rowOut, 1).Value = "金额合计"
    draftSheet.Cells(rowOut, 2).Value = amountSum
    draftSheet.Cells(rowOut, 3).Value = "有效金额行=" & CStr(amountCount)
    rowOut = rowOut + 1
    draftSheet.Cells(rowOut, 1).Value = "金额均值"
    draftSheet.Cells(rowOut, 2).Value = amountAvg
    draftSheet.Cells(rowOut, 3).Value = "仅统计数值行"
    rowOut = rowOut + 1
    draftSheet.Cells(rowOut, 1).Value = "进行中/未完成"
    draftSheet.Cells(rowOut, 2).Value = openCount
    draftSheet.Cells(rowOut, 3).Value = "按状态字段粗分"
    rowOut = rowOut + 1
    draftSheet.Cells(rowOut, 1).Value = "已完成"
    draftSheet.Cells(rowOut, 2).Value = doneCount
    draftSheet.Cells(rowOut, 3).Value = "按状态字段粗分"
    rowOut = rowOut + 1
    draftSheet.Cells(rowOut, 1).Value = "逾期候选"
    draftSheet.Cells(rowOut, 2).Value = overdueCount
    draftSheet.Cells(rowOut, 3).Value = "日期早于今天且未判完成"
    rowOut = rowOut + 2

    draftSheet.Cells(rowOut, 1).Value = "三、异常/待办候选"
    draftSheet.Cells(rowOut, 1).Font.Bold = True
    rowOut = rowOut + 1
    If anomalyCount = 0 Then
        draftSheet.Cells(rowOut, 1).Value = "未自动识别到明显异常，请人工补充。"
        rowOut = rowOut + 1
    Else
        Dim lines, li
        lines = Split(anomalyLines, vbCrLf)
        For li = 0 To UBound(lines)
            draftSheet.Cells(rowOut, 1).Value = lines(li)
            rowOut = rowOut + 1
        Next
    End If
    rowOut = rowOut + 1

    draftSheet.Cells(rowOut, 1).Value = "四、分组观察（负责人）"
    draftSheet.Cells(rowOut, 1).Font.Bold = True
    rowOut = rowOut + 1
    If ownerN = 0 Then
        draftSheet.Cells(rowOut, 1).Value = "未指定负责人字段。"
        rowOut = rowOut + 1
    Else
        draftSheet.Cells(rowOut, 1).Value = sampleOwners
        rowOut = rowOut + 1
    End If
    rowOut = rowOut + 1

    draftSheet.Cells(rowOut, 1).Value = "五、重点说明（待人工确认）"
    draftSheet.Cells(rowOut, 1).Font.Bold = True
    rowOut = rowOut + 1
    draftSheet.Cells(rowOut, 1).Value = "1. 请补充本期最重要的进展与结论。"
    rowOut = rowOut + 1
    draftSheet.Cells(rowOut, 1).Value = "2. 请核对异常候选是否属实，并标注责任人。"
    rowOut = rowOut + 1
    draftSheet.Cells(rowOut, 1).Value = "3. 请补充需要同步给上级/协作方的风险点。"
    rowOut = rowOut + 2

    draftSheet.Cells(rowOut, 1).Value = "六、下期计划（待人工确认）"
    draftSheet.Cells(rowOut, 1).Font.Bold = True
    rowOut = rowOut + 1
    draftSheet.Cells(rowOut, 1).Value = "1. 目标："
    rowOut = rowOut + 1
    draftSheet.Cells(rowOut, 1).Value = "2. 关键动作："
    rowOut = rowOut + 1
    draftSheet.Cells(rowOut, 1).Value = "3. 需要支持："
    rowOut = rowOut + 2

    draftSheet.Cells(rowOut, 1).Value = "说明：本页仅为事实初稿，不替代人工判断；sourceUnchanged=true"
    draftSheet.Columns.AutoFit

    Set summarySheet = CreateOutputSheet(appObj, sourceName, "处理摘要", outputMode)
    If summarySheet Is Nothing Then
        RollbackCreatedSheets appObj, createdSheets, createdCount
        HostExcelReportDraftPack = FailureJson("E_OUTPUT_SHEET", "无法创建处理摘要工作表")
        Exit Function
    End If
    createdCount = createdCount + 1: Set createdSheets(createdCount) = summarySheet

    Dim summaryLines, impactJson, summary
    summaryLines = "场景名=周报月报初稿场景包" & vbCrLf & _
        "命令ID=excel.report_draft_pack" & vbCrLf & _
        "源表名称=" & sourceName & vbCrLf & _
        "源区域=" & sourceRange.Address & vbCrLf & _
        "输出模式=" & outputMode & vbCrLf & _
        "输出位置=" & draftSheet.Name & "," & summarySheet.Name & vbCrLf & _
        "处理前行数=" & CStr(dataRows) & vbCrLf & _
        "处理后行数=" & CStr(dataRows) & vbCrLf & _
        "参数=type=" & reportType & ";title=" & reportTitle & ";amount=" & amountField & ";status=" & statusField & ";owner=" & ownerField & ";date=" & dateField & vbCrLf & _
        "金额合计=" & CStr(amountSum) & vbCrLf & _
        "进行中=" & CStr(openCount) & vbCrLf & _
        "已完成=" & CStr(doneCount) & vbCrLf & _
        "逾期候选=" & CStr(overdueCount) & vbCrLf & _
        "异常样本数=" & CStr(anomalyCount) & vbCrLf & _
        "执行时间=" & Now & vbCrLf & _
        "结果=成功"
    impactJson = SafeHostText("ExcelWriteImpactSummary", summaryLines, summarySheet.Name)
    If Not ExtractJsonBoolean(impactJson, "ok") Then WriteSummaryFallback summarySheet, summaryLines

    summary = "Excel 报表初稿完成：输出=" & draftSheet.Name & "；类型=" & reportType & "；行=" & CStr(dataRows) & "；逾期候选=" & CStr(overdueCount) & "；sourceUnchanged=true"
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostExcelReportDraftPack = "{""ok"":true,""sourceUnchanged"":true,""commandId"":""excel.report_draft_pack"",""outputSheet"":""" & EscapeJson(draftSheet.Name) & """,""summarySheet"":""" & EscapeJson(summarySheet.Name) & """,""reportType"":""" & EscapeJson(reportType) & """,""dataRows"":" & CStr(dataRows) & ",""overdueCount"":" & CStr(overdueCount) & ",""writePlanPreview"":""" & EscapeJson(planPreview) & """,""message"":""" & EscapeJson(summary) & """}"
End Function

Function NormalizeReportType(text)
    Dim t
    t = LCase(Trim(CStr(text)))
    If t = "month" Or t = "monthly" Or t = "月" Or t = "月报" Then
        NormalizeReportType = "月报"
    Else
        NormalizeReportType = "周报"
    End If
End Function

Function SceneRulesDir()
    Dim basePath
    On Error Resume Next
    basePath = ""
    basePath = CStr(Host.GetKnownFolderPath("localappdata"))
    If Len(Trim(basePath)) = 0 Then basePath = CStr(Host.GetKnownFolderPath("appdata"))
    If Len(Trim(basePath)) = 0 Then basePath = CStr(Host.GetKnownFolderPath("temp"))
    If Len(Trim(basePath)) = 0 Then basePath = "."
    SceneRulesDir = Host.CombinePath(Host.CombinePath(basePath, "OfficeAddin"), "scene_rules")
    Host.CreateFolder SceneRulesDir
    Err.Clear
End Function

Function RuleFilePath(kindName, ruleName)
    Dim safeName, i, ch, buf, fileName
    safeName = LCase(Trim(CStr(ruleName)))
    If Len(safeName) = 0 Then safeName = "last"
    buf = ""
    For i = 1 To Len(safeName)
        ch = Mid(safeName, i, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Or ch = "_" Or ch = "-" Then
            buf = buf & ch
        ElseIf AscW(ch) > 127 Then
            buf = buf & ch
        Else
            buf = buf & "_"
        End If
    Next
    If Len(buf) = 0 Then buf = "last"
    fileName = CStr(kindName) & "_" & buf & ".txt"
    fileName = Host.SanitizeFileName(fileName, "last.txt")
    RuleFilePath = Host.CombinePath(SceneRulesDir(), fileName)
End Function

Function ReadTextFile(pathText)
    On Error Resume Next
    ReadTextFile = ""
    If Host.PathExists(CStr(pathText)) Then
        ReadTextFile = CStr(Host.ReadTextFile(CStr(pathText)))
        If Err.Number <> 0 Then
            ReadTextFile = ""
            Err.Clear
        End If
    End If
    Err.Clear
End Function

Function WriteTextFile(pathText, contentText)
    On Error Resume Next
    WriteTextFile = False
    WriteTextFile = Host.WriteTextFile(CStr(pathText), CStr(contentText), True)
    If Err.Number <> 0 Then
        WriteTextFile = False
        Err.Clear
    End If
    Err.Clear
End Function
Function ExtractRuleValue(text, keyName)
    Dim lines, i, lineText, prefix
    ExtractRuleValue = ""
    prefix = LCase(Trim(CStr(keyName))) & "="
    lines = Split(Replace(CStr(text), vbCrLf, vbLf), vbLf)
    For i = 0 To UBound(lines)
        lineText = Trim(Replace(lines(i), vbCr, ""))
        If LCase(Left(lineText, Len(prefix))) = prefix Then
            ExtractRuleValue = Mid(lineText, Len(prefix) + 1)
            Exit Function
        End If
    Next
End Function

Function GuessHeaderByKeywords(sourceRange, keywordsCsv)
    Dim keys, c, headerText, k
    GuessHeaderByKeywords = ""
    keys = SplitCsvFields(keywordsCsv)
    If Not IsArray(keys) Then Exit Function
    On Error Resume Next
    If UBound(keys) < 0 Then Exit Function
    For c = 1 To sourceRange.Columns.Count
        headerText = NormalizeCellText(sourceRange.Cells(1, c).Text)
        For k = 0 To UBound(keys)
            If Len(keys(k)) > 0 Then
                If InStr(1, headerText, keys(k), vbTextCompare) > 0 Or StrComp(headerText, keys(k), vbTextCompare) = 0 Then
                    GuessHeaderByKeywords = CStr(sourceRange.Cells(1, c).Text)
                    Exit Function
                End If
            End If
        Next
    Next
    Err.Clear
End Function

Sub MaybeSaveRule(kindName, defaultName, contentText)
    Dim ans, pathText, ok
    ans = Trim(SafePrompt("保存规则：否 / 上次 / 输入名称（默认 " & defaultName & "）", "上次"))
    If Len(ans) = 0 Then ans = "否"
    If StrComp(ans, "否", vbTextCompare) = 0 Or LCase(ans) = "no" Or LCase(ans) = "n" Then Exit Sub
    If StrComp(ans, "上次", vbTextCompare) = 0 Or LCase(ans) = "last" Then
        pathText = RuleFilePath(kindName, "last")
    Else
        pathText = RuleFilePath(kindName, ans)
    End If
    ok = WriteTextFile(pathText, contentText)
    WriteTextFile RuleFilePath(kindName, "last"), contentText
    If ok Then SafeWriteLog "规则已保存: " & pathText
End Sub

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
    ElseIf t = "ne" Or t = "!=" Or t = "<>" Or t = "不等于" Or t = "not" Or t = "not_equals" Then
        NormalizeMatchMode = "ne"
    ElseIf t = "gt" Or t = ">" Or t = "大于" Then
        NormalizeMatchMode = "gt"
    ElseIf t = "lt" Or t = "<" Or t = "小于" Then
        NormalizeMatchMode = "lt"
    ElseIf t = "gte" Or t = ">=" Or t = "大于等于" Then
        NormalizeMatchMode = "gte"
    ElseIf t = "lte" Or t = "<=" Or t = "小于等于" Then
        NormalizeMatchMode = "lte"
    ElseIf t = "date_lt" Or t = "日期早于" Or t = "before" Or t = "逾期" Then
        NormalizeMatchMode = "date_lt"
    ElseIf t = "date_gt" Or t = "日期晚于" Or t = "after" Then
        NormalizeMatchMode = "date_gt"
    Else
        NormalizeMatchMode = "exact"
    End If
End Function

Function MatchModeLabel(value)
    Dim m
    m = NormalizeMatchMode(value)
    If m = "contains" Then
        MatchModeLabel = "包含"
    ElseIf m = "ne" Then
        MatchModeLabel = "不等于"
    ElseIf m = "gt" Then
        MatchModeLabel = "大于"
    ElseIf m = "lt" Then
        MatchModeLabel = "小于"
    ElseIf m = "gte" Then
        MatchModeLabel = "大于等于"
    ElseIf m = "lte" Then
        MatchModeLabel = "小于等于"
    ElseIf m = "date_lt" Then
        MatchModeLabel = "日期早于"
    ElseIf m = "date_gt" Then
        MatchModeLabel = "日期晚于"
    Else
        MatchModeLabel = "精确"
    End If
End Function

Function ResolveCompareTarget(keyword)

    Dim t, normalized

    t = Trim(CStr(keyword))

    If StrComp(t, "TODAY", vbTextCompare) = 0 Or t = "今天" Or t = "今日" Then

        ResolveCompareTarget = Date

        Exit Function

    End If

    If TryNormalizeDateText(t, normalized) Then

        ResolveCompareTarget = CDate(normalized)

        Exit Function

    End If

    If IsNumeric(t) Then

        ResolveCompareTarget = CDbl(t)

        Exit Function

    End If

    ResolveCompareTarget = t

End Function

Function IsMatch(cellText, keyword, matchMode)
    Dim leftText, rightText, mode, leftVal, rightVal, leftDate, rightDate, normalized
    leftText = NormalizeCellText(cellText)
    rightText = NormalizeCellText(keyword)
    mode = NormalizeMatchMode(matchMode)

    If mode = "exact" Then
        If Len(rightText) = 0 Then
            IsMatch = (Len(leftText) = 0)
        Else
            IsMatch = (StrComp(leftText, rightText, vbTextCompare) = 0)
        End If
        Exit Function
    End If

    If mode = "contains" Then
        If Len(rightText) = 0 Then
            IsMatch = (Len(leftText) = 0)
        Else
            IsMatch = (InStr(1, leftText, rightText, vbTextCompare) > 0)
        End If
        Exit Function
    End If

    If mode = "ne" Then
        If Len(rightText) = 0 Then
            IsMatch = (Len(leftText) > 0)
        Else
            IsMatch = (StrComp(leftText, rightText, vbTextCompare) <> 0)
        End If
        Exit Function
    End If

    If mode = "date_lt" Or mode = "date_gt" Then
        If Not TryNormalizeDateText(leftText, normalized) Then
            IsMatch = False
            Exit Function
        End If
        leftDate = CDate(normalized)
        rightDate = ResolveCompareTarget(keyword)
        If Not IsDate(rightDate) Then
            IsMatch = False
            Exit Function
        End If
        If mode = "date_lt" Then
            IsMatch = (leftDate < CDate(rightDate))
        Else
            IsMatch = (leftDate > CDate(rightDate))
        End If
        Exit Function
    End If

    ' numeric compares
    If Len(leftText) = 0 Or Not IsNumeric(leftText) Then
        IsMatch = False
        Exit Function
    End If
    leftVal = CDbl(leftText)
    rightVal = ResolveCompareTarget(keyword)
    If Not IsNumeric(rightVal) Then
        IsMatch = False
        Exit Function
    End If
    rightVal = CDbl(rightVal)
    If mode = "gt" Then
        IsMatch = (leftVal > rightVal)
    ElseIf mode = "lt" Then
        IsMatch = (leftVal < rightVal)
    ElseIf mode = "gte" Then
        IsMatch = (leftVal >= rightVal)
    ElseIf mode = "lte" Then
        IsMatch = (leftVal <= rightVal)
    Else
        IsMatch = False
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
    FailureJson = "{""ok"":false,""code"":""" & EscapeJson(code) & """,""message"":""" & EscapeJson(message) & """,""sourceUnchanged"":true}"
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
