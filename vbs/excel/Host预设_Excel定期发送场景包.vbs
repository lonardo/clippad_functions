' 函数名: HostExcelScheduledSendPack
' 描述: 定期发送场景包。按日报/周报/未完成清单规则生成发送清单与发送记录，可选导出 CSV；预览确认后只写输出，不改源表。
' 适用应用: Excel
' 搜索范围: 当前范围
' 搜索对象: 无
' 作用范围: 当前范围
' 需要选区: 是
' 破坏范围: 当前范围
' 需要备份: 否

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
    Main = HostExcelScheduledSendPack(appObj)
End Function

Function HostExcelScheduledSendPack(appObj)
    On Error Resume Next
    Dim sourceRange, sourceSheet, sourceName, outputMode
    Set sourceRange = appObj.Selection.CurrentRegion
    If Err.Number <> 0 Or TypeName(sourceRange) = "Empty" Or TypeName(sourceRange) = "Nothing" Then
        Err.Clear
        HostExcelScheduledSendPack = FailureJson("E_NO_RANGE", "请先选中含表头的数据区域")
        Exit Function
    End If
    If sourceRange.Rows.Count < 2 Then
        HostExcelScheduledSendPack = FailureJson("E_RANGE_TOO_SMALL", "当前区域至少需要一行表头和一行数据")
        Exit Function
    End If
    Set sourceSheet = sourceRange.Worksheet
    sourceName = sourceSheet.Name

    Dim sceneName, statusField, ownerField, dateField, recipientText, deliveryMode
    sceneName = NormalizeSendScene(SafePrompt("发送场景：日报 / 周报 / 未完成清单（默认 日报）", "日报"))
    statusField = Trim(SafePrompt("状态字段（可空）", GuessHeaderByKeywords(sourceRange, "状态,进度,完成情况")))
    ownerField = Trim(SafePrompt("负责人字段（可空）", GuessHeaderByKeywords(sourceRange, "负责人,跟进人,处理人,责任人,owner")))
    dateField = Trim(SafePrompt("日期字段（可空）", GuessHeaderByKeywords(sourceRange, "日期,截止日期,到期日,业务日期,完成日期")))
    recipientText = Trim(SafePrompt("接收人备注（可空，仅写入发送记录）", "相关同事"))
    If Len(recipientText) = 0 Then recipientText = "相关同事"
    deliveryMode = NormalizeDeliveryMode(SafePrompt("交付方式：sheet / export / both（默认 both）", "both"))
    outputMode = NormalizeOutputMode(SafePrompt("输出模式：NewSheet 或 NewWorkbook（默认 NewSheet）", "NewSheet"))

    Dim statusCol, ownerCol, dateCol, dataRows
    statusCol = 0: ownerCol = 0: dateCol = 0
    If Len(statusField) > 0 Then statusCol = FindHeaderColumn(sourceRange, statusField)
    If Len(ownerField) > 0 Then ownerCol = FindHeaderColumn(sourceRange, ownerField)
    If Len(dateField) > 0 Then dateCol = FindHeaderColumn(sourceRange, dateField)
    If Len(statusField) > 0 And statusCol <= 0 Then
        HostExcelScheduledSendPack = FailureJson("E_FIELD_NOT_FOUND", "未找到状态字段：" & statusField)
        Exit Function
    End If
    If Len(ownerField) > 0 And ownerCol <= 0 Then
        HostExcelScheduledSendPack = FailureJson("E_FIELD_NOT_FOUND", "未找到负责人字段：" & ownerField)
        Exit Function
    End If
    If Len(dateField) > 0 And dateCol <= 0 Then
        HostExcelScheduledSendPack = FailureJson("E_FIELD_NOT_FOUND", "未找到日期字段：" & dateField)
        Exit Function
    End If

    dataRows = sourceRange.Rows.Count - 1
    Dim matchFlags(), matchCount, openCount, doneCount, overdueCount
    Dim r, statusText, isDone, isOpen, isOverdue, text, normalized, includeRow
    Dim hasDate, rowDate, weekStart, weekEnd, windowText
    ReDim matchFlags(sourceRange.Rows.Count)
    matchCount = 0: openCount = 0: doneCount = 0: overdueCount = 0
    ' week window starts Monday (Weekday(..., vbMonday)).
    weekStart = Date - (Weekday(Date, vbMonday) - 1)
    weekEnd = weekStart + 6
    If sceneName = "周报" Then
        windowText = "本周 " & FormatDateTime(weekStart, vbShortDate) & " ~ " & FormatDateTime(weekEnd, vbShortDate)
    ElseIf sceneName = "日报" Then
        windowText = "今天 " & FormatDateTime(Date, vbShortDate)
    Else
        windowText = "未完成/开口项"
    End If

    For r = 2 To sourceRange.Rows.Count
        statusText = ""
        isDone = False
        isOpen = False
        isOverdue = False
        hasDate = False
        rowDate = Date
        If statusCol > 0 Then
            statusText = NormalizeCellText(sourceRange.Cells(r, statusCol).Text)
            If IsDoneStatus(statusText) Then
                isDone = True
                doneCount = doneCount + 1
            Else
                isOpen = True
                openCount = openCount + 1
            End If
        End If
        If dateCol > 0 Then
            ' Prefer Value/DateValue so narrow-column Text (###) does not drop real dates.
            On Error Resume Next
            If IsDate(sourceRange.Cells(r, dateCol).Value) Then
                hasDate = True
                rowDate = DateValue(CDate(sourceRange.Cells(r, dateCol).Value))
            End If
            If Err.Number <> 0 Then
                Err.Clear
                hasDate = False
            End If
            On Error Resume Next
            If Not hasDate Then
                text = NormalizeCellText(sourceRange.Cells(r, dateCol).Text)
                If TryNormalizeDateText(text, normalized) Then
                    hasDate = True
                    rowDate = DateValue(CDate(normalized))
                End If
            End If
            If hasDate Then
                If rowDate < Date Then
                    If statusCol <= 0 Or isOpen Then
                        isOverdue = True
                        overdueCount = overdueCount + 1
                    End If
                End If
            End If
        End If

        includeRow = False
        If sceneName = "未完成清单" Then
            If statusCol > 0 Then
                includeRow = isOpen
            ElseIf dateCol > 0 Then
                includeRow = isOverdue
            Else
                includeRow = True
            End If
        ElseIf sceneName = "周报" Then
            ' weekly: date in current Mon-Sun window; no date col => full snapshot.
            If dateCol <= 0 Then
                includeRow = True
            ElseIf hasDate Then
                includeRow = (rowDate >= weekStart And rowDate <= weekEnd)
            Else
                ' unparsable date keeps open items to avoid silent drop.
                includeRow = (statusCol <= 0 Or isOpen)
            End If
        Else
            ' daily: only today; no date col => full snapshot.
            If dateCol <= 0 Then
                includeRow = True
            ElseIf hasDate Then
                includeRow = (rowDate = DateValue(Date))
            Else
                includeRow = (statusCol <= 0 Or isOpen)
            End If
        End If
        matchFlags(r) = includeRow
        If includeRow Then matchCount = matchCount + 1
    Next

    Dim needExport, needSheet, formatName, delimiter, extension, outputPlan, outputPath, defaultName
    needExport = (deliveryMode = "export" Or deliveryMode = "both")
    needSheet = (deliveryMode = "sheet" Or deliveryMode = "both")
    formatName = "csv"
    delimiter = ","
    extension = ".csv"
    outputPath = ""
    If needExport Then
        formatName = NormalizeExportFormat(SafePrompt("导出格式：CSV 或 TSV", "CSV"))
        If formatName = "tsv" Then
            delimiter = vbTab
            extension = ".tsv"
        End If
        defaultName = "excel_scheduled_send_" & FormatFileStamp(Now()) & extension
        On Error Resume Next
        outputPlan = Host.ResolveOutputPlan(defaultName, "avoid")
        If Err.Number <> 0 Then
            Err.Clear
            outputPlan = ""
        End If
        outputPath = ExtractJsonString(outputPlan, "path")
        If Len(outputPath) = 0 Then
            HostExcelScheduledSendPack = FailureJson("E_OUTPUT_PATH", "无法生成避免覆盖的导出文件路径，未改表")
            Exit Function
        End If
    End If

    Dim previewText, planId, planPreview
    previewText = "Excel 定期发送预览（尚未写入）" & vbCrLf & _
        "场景ID：excel.scheduled_send_pack" & vbCrLf & _
        "源表=" & sourceName & "，范围=" & sourceRange.Address & vbCrLf & _
        "发送场景=" & sceneName & "，交付方式=" & deliveryMode & "，输出模式=" & outputMode & vbCrLf & _
        "接收人=" & recipientText & vbCrLf & _
        "字段：状态=" & statusField & "，负责人=" & ownerField & "，日期=" & dateField & vbCrLf & _
        "日期窗口=" & windowText & vbCrLf & _
        "源数据行=" & CStr(dataRows) & "，待发送行=" & CStr(matchCount) & _
        "，进行中/未完成=" & CStr(openCount) & "，已完成=" & CStr(doneCount) & "，逾期候选=" & CStr(overdueCount) & vbCrLf
    If needExport Then previewText = previewText & "导出=" & UCase(formatName) & "，路径=" & outputPath & vbCrLf
    previewText = previewText & "将生成：发送清单" & IIf(needSheet, "工作表", "") & IIf(needExport, "/导出文件", "") & " + 发送记录 + 处理摘要；sourceUnchanged=true，不改源表"

    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_scheduled_send_pack", "{""source"":""" & EscapeJson(sourceRange.Address) & """,""scene"":""" & EscapeJson(sceneName) & """,""matchCount"":" & CStr(matchCount) & ",""deliveryMode"":""" & EscapeJson(deliveryMode) & """}", "office.excel.scheduledSendPack"
    planPreview = SafePreviewWritePlan()
    SafeRollbackWritePlan planId

    If Not SafeConfirmStep(previewText & vbCrLf & vbCrLf & "确认生成定期发送清单与记录？", "excel.scheduled_send_pack") Then
        HostExcelScheduledSendPack = FailureJson("E_CONFIRM_REQUIRED", "用户取消或未确认，未修改工作簿与文件")
        Exit Function
    End If

    Dim createdSheets(), createdCount, listSheet, recordSheet, summarySheet
    createdCount = 0
    ReDim createdSheets(6)
    Set listSheet = Nothing
    Set recordSheet = Nothing

    If needSheet Then
        Set listSheet = CreateOutputSheet(appObj, sourceName, "发送清单", outputMode)
        If listSheet Is Nothing Then
            HostExcelScheduledSendPack = FailureJson("E_OUTPUT_SHEET", "无法创建发送清单工作表")
            Exit Function
        End If
        createdCount = createdCount + 1: Set createdSheets(createdCount) = listSheet
        If Not WriteSendListSheet(sourceRange, listSheet, matchFlags, sceneName, recipientText, statusField, ownerField, dateField, matchCount) Then
            RollbackCreatedSheets appObj, createdSheets, createdCount
            HostExcelScheduledSendPack = FailureJson("E_LIST_WRITE", "发送清单写入失败，已删除未完成输出")
            Exit Function
        End If
    End If

    Dim exportText, written
    written = False
    If needExport Then
        exportText = BuildSendListExportText(sourceRange, matchFlags, formatName, delimiter)
        On Error Resume Next
        written = Host.WriteTextFile(outputPath, exportText, False)
        If Err.Number <> 0 Then
            Err.Clear
            written = False
        End If
        If Not CBool(written) Then
            RollbackCreatedSheets appObj, createdSheets, createdCount
            HostExcelScheduledSendPack = FailureJson("E_EXPORT_WRITE", "导出文件写入失败，未修改源表：" & outputPath)
            Exit Function
        End If
    End If

    Set recordSheet = CreateOutputSheet(appObj, sourceName, "发送记录", outputMode)
    If recordSheet Is Nothing Then
        RollbackCreatedSheets appObj, createdSheets, createdCount
        HostExcelScheduledSendPack = FailureJson("E_OUTPUT_SHEET", "无法创建发送记录工作表")
        Exit Function
    End If
    createdCount = createdCount + 1: Set createdSheets(createdCount) = recordSheet
    WriteSendRecordSheet recordSheet, sceneName, recipientText, sourceName, sourceRange.Address, matchCount, dataRows, openCount, doneCount, overdueCount, deliveryMode, outputPath, statusField, ownerField, dateField

    Set summarySheet = CreateOutputSheet(appObj, sourceName, "处理摘要", outputMode)
    If summarySheet Is Nothing Then
        RollbackCreatedSheets appObj, createdSheets, createdCount
        HostExcelScheduledSendPack = FailureJson("E_OUTPUT_SHEET", "无法创建处理摘要工作表")
        Exit Function
    End If
    createdCount = createdCount + 1: Set createdSheets(createdCount) = summarySheet

    Dim summaryLines, impactJson, summary, outputNames, listName
    listName = ""
    If needSheet Then listName = listSheet.Name
    outputNames = ""
    If needSheet Then outputNames = listSheet.Name
    If Len(outputNames) > 0 Then outputNames = outputNames & ","
    outputNames = outputNames & recordSheet.Name & "," & summarySheet.Name
    If needExport Then outputNames = outputNames & "," & outputPath

    summaryLines = "场景名=定期发送场景包" & vbCrLf & _
        "场景ID=excel.scheduled_send_pack" & vbCrLf & _
        "源表名称=" & sourceName & vbCrLf & _
        "源范围=" & sourceRange.Address & vbCrLf & _
        "输出模式=" & outputMode & vbCrLf & _
        "交付方式=" & deliveryMode & vbCrLf & _
        "发送场景=" & sceneName & vbCrLf & _
        "接收人=" & recipientText & vbCrLf & _
        "输出位置=" & outputNames & vbCrLf & _
        "处理前行数=" & CStr(dataRows) & vbCrLf & _
        "处理后行数=" & CStr(matchCount) & vbCrLf & _
        "规则=scene=" & sceneName & ";status=" & statusField & ";owner=" & ownerField & ";date=" & dateField & vbCrLf & _
        "日期窗口=" & windowText & vbCrLf & _
        "进行中=" & CStr(openCount) & vbCrLf & _
        "已完成=" & CStr(doneCount) & vbCrLf & _
        "逾期候选=" & CStr(overdueCount) & vbCrLf & _
        "导出路径=" & outputPath & vbCrLf & _
        "执行时间=" & Now & vbCrLf & _
        "结果=成功（已生成发送清单/记录，实际推送需外部通道）"
    impactJson = SafeHostText("ExcelWriteImpactSummary", summaryLines, summarySheet.Name)
    If Not ExtractJsonBoolean(impactJson, "ok") Then WriteSummaryFallback summarySheet, summaryLines

    summary = "Excel 定期发送准备完成：场景=" & sceneName & "，待发送行=" & CStr(matchCount)
    If needSheet Then summary = summary & "，清单=" & listName
    summary = summary & "，记录=" & recordSheet.Name
    If needExport Then summary = summary & "，文件=" & outputPath
    summary = summary & "，sourceUnchanged=true"
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostExcelScheduledSendPack = "{""ok"":true,""sourceUnchanged"":true,""commandId"":""excel.scheduled_send_pack"",""scene"":""" & EscapeJson(sceneName) & """,""matchCount"":" & CStr(matchCount) & ",""deliveryMode"":""" & EscapeJson(deliveryMode) & """,""listSheet"":""" & EscapeJson(listName) & """,""recordSheet"":""" & EscapeJson(recordSheet.Name) & """,""summarySheet"":""" & EscapeJson(summarySheet.Name) & """,""exportPath"":""" & EscapeJson(outputPath) & """,""recipient"":""" & EscapeJson(recipientText) & """,""writePlanPreview"":""" & EscapeJson(planPreview) & """,""message"":""" & EscapeJson(summary) & """}"
End Function

Function NormalizeSendScene(text)
    Dim t
    t = LCase(Trim(CStr(text)))
    If t = "week" Or t = "weekly" Or t = "周" Or t = "周报" Then
        NormalizeSendScene = "周报"
    ElseIf t = "open" Or t = "todo" Or t = "未完成" Or t = "未完成清单" Or t = "3" Then
        NormalizeSendScene = "未完成清单"
    Else
        NormalizeSendScene = "日报"
    End If
End Function

Function NormalizeDeliveryMode(value)
    Dim t
    t = LCase(Trim(CStr(value)))
    If t = "sheet" Or t = "工作表" Or t = "表" Then
        NormalizeDeliveryMode = "sheet"
    ElseIf t = "export" Or t = "file" Or t = "导出" Or t = "csv" Or t = "tsv" Then
        NormalizeDeliveryMode = "export"
    Else
        NormalizeDeliveryMode = "both"
    End If
End Function

Function IsDoneStatus(statusText)
    Dim t
    t = NormalizeCellText(statusText)
    If Len(t) = 0 Then
        IsDoneStatus = False
        Exit Function
    End If
    ' Open/in-progress statuses must win over generic completion substring matches.
    If InStr(1, t, "未完成", vbTextCompare) > 0 Or InStr(1, t, "未完", vbTextCompare) > 0 Or InStr(1, t, "进行中", vbTextCompare) > 0 Or InStr(1, t, "处理中", vbTextCompare) > 0 Or InStr(1, t, "待处理", vbTextCompare) > 0 Or InStr(1, t, "todo", vbTextCompare) > 0 Or InStr(1, t, "open", vbTextCompare) > 0 Or InStr(1, t, "pending", vbTextCompare) > 0 Then
        IsDoneStatus = False
        Exit Function
    End If
    If InStr(1, t, "已完成", vbTextCompare) > 0 Or InStr(1, t, "完成", vbTextCompare) > 0 Or InStr(1, t, "关闭", vbTextCompare) > 0 Or InStr(1, t, "已回款", vbTextCompare) > 0 Or InStr(1, t, "done", vbTextCompare) > 0 Or InStr(1, t, "closed", vbTextCompare) > 0 Or InStr(1, t, "complete", vbTextCompare) > 0 Then
        IsDoneStatus = True
    Else
        IsDoneStatus = False
    End If
End Function

Function WriteSendListSheet(sourceRange, listSheet, matchFlags, sceneName, recipientText, statusField, ownerField, dateField, matchCount)
    On Error Resume Next
    Dim c, r, outRow, colCount
    WriteSendListSheet = False
    colCount = sourceRange.Columns.Count
    listSheet.Cells(1, 1).Value = "定期发送清单"
    listSheet.Cells(1, 1).Font.Bold = True
    listSheet.Cells(1, 1).Font.Size = 13
    listSheet.Cells(2, 1).Value = "发送场景"
    listSheet.Cells(2, 2).Value = sceneName
    listSheet.Cells(3, 1).Value = "接收人"
    listSheet.Cells(3, 2).Value = recipientText
    listSheet.Cells(4, 1).Value = "生成时间"
    listSheet.Cells(4, 2).Value = Now
    listSheet.Cells(5, 1).Value = "待发送行数"
    listSheet.Cells(5, 2).Value = matchCount
    listSheet.Cells(6, 1).Value = "状态字段"
    listSheet.Cells(6, 2).Value = IIf(Len(statusField) > 0, statusField, "（未指定）")
    listSheet.Cells(7, 1).Value = "负责人字段"
    listSheet.Cells(7, 2).Value = IIf(Len(ownerField) > 0, ownerField, "（未指定）")
    listSheet.Cells(8, 1).Value = "日期字段"
    listSheet.Cells(8, 2).Value = IIf(Len(dateField) > 0, dateField, "（未指定）")
    listSheet.Cells(9, 1).Value = "说明"
    listSheet.Cells(9, 2).Value = "本页为待发送事实清单；实际邮件/IM 推送由外部通道完成"

    outRow = 11
    For c = 1 To colCount
        listSheet.Cells(outRow, c).Value = sourceRange.Cells(1, c).Text
        listSheet.Cells(outRow, c).Font.Bold = True
    Next
    outRow = outRow + 1
    For r = 2 To sourceRange.Rows.Count
        If matchFlags(r) Then
            For c = 1 To colCount
                listSheet.Cells(outRow, c).Value = sourceRange.Cells(r, c).Text
            Next
            outRow = outRow + 1
        End If
    Next
    listSheet.Columns.AutoFit
    If Err.Number <> 0 Then
        Err.Clear
        WriteSendListSheet = False
        Exit Function
    End If
    WriteSendListSheet = True
End Function

Sub WriteSendRecordSheet(recordSheet, sceneName, recipientText, sourceName, sourceAddress, matchCount, dataRows, openCount, doneCount, overdueCount, deliveryMode, outputPath, statusField, ownerField, dateField)
    On Error Resume Next
    recordSheet.Cells(1, 1).Value = "发送记录"
    recordSheet.Cells(1, 1).Font.Bold = True
    recordSheet.Cells(1, 1).Font.Size = 13
    recordSheet.Cells(3, 1).Value = "字段"
    recordSheet.Cells(3, 2).Value = "值"
    recordSheet.Rows(3).Font.Bold = True
    recordSheet.Cells(4, 1).Value = "记录时间"
    recordSheet.Cells(4, 2).Value = Now
    recordSheet.Cells(5, 1).Value = "发送场景"
    recordSheet.Cells(5, 2).Value = sceneName
    recordSheet.Cells(6, 1).Value = "接收人"
    recordSheet.Cells(6, 2).Value = recipientText
    recordSheet.Cells(7, 1).Value = "源表"
    recordSheet.Cells(7, 2).Value = sourceName
    recordSheet.Cells(8, 1).Value = "源范围"
    recordSheet.Cells(8, 2).Value = sourceAddress
    recordSheet.Cells(9, 1).Value = "交付方式"
    recordSheet.Cells(9, 2).Value = deliveryMode
    recordSheet.Cells(10, 1).Value = "待发送行数"
    recordSheet.Cells(10, 2).Value = matchCount
    recordSheet.Cells(11, 1).Value = "源数据行数"
    recordSheet.Cells(11, 2).Value = dataRows
    recordSheet.Cells(12, 1).Value = "进行中/未完成"
    recordSheet.Cells(12, 2).Value = openCount
    recordSheet.Cells(13, 1).Value = "已完成"
    recordSheet.Cells(13, 2).Value = doneCount
    recordSheet.Cells(14, 1).Value = "逾期候选"
    recordSheet.Cells(14, 2).Value = overdueCount
    recordSheet.Cells(15, 1).Value = "状态字段"
    recordSheet.Cells(15, 2).Value = statusField
    recordSheet.Cells(16, 1).Value = "负责人字段"
    recordSheet.Cells(16, 2).Value = ownerField
    recordSheet.Cells(17, 1).Value = "日期字段"
    recordSheet.Cells(17, 2).Value = dateField
    recordSheet.Cells(18, 1).Value = "导出路径"
    recordSheet.Cells(18, 2).Value = outputPath
    recordSheet.Cells(19, 1).Value = "发送状态"
    recordSheet.Cells(19, 2).Value = "已生成待发送"
    recordSheet.Cells(20, 1).Value = "说明"
    recordSheet.Cells(20, 2).Value = "本记录只证明清单已生成；真正定时推送可对接邮件/IM/机器人"
    recordSheet.Columns.AutoFit
End Sub

Function BuildSendListExportText(sourceRange, matchFlags, formatName, delimiter)
    On Error Resume Next
    Dim lines, r, c, rowText, colCount
    lines = ""
    colCount = sourceRange.Columns.Count
    rowText = ""
    For c = 1 To colCount
        If c > 1 Then rowText = rowText & delimiter
        rowText = rowText & EscapeExportField(CStr(sourceRange.Cells(1, c).Text), formatName)
    Next
    lines = rowText
    For r = 2 To sourceRange.Rows.Count
        If matchFlags(r) Then
            rowText = ""
            For c = 1 To colCount
                If c > 1 Then rowText = rowText & delimiter
                rowText = rowText & EscapeExportField(CStr(sourceRange.Cells(r, c).Text), formatName)
            Next
            lines = lines & vbCrLf & rowText
        End If
    Next
    BuildSendListExportText = lines
End Function

Function EscapeExportField(value, formatName)
    Dim text
    On Error Resume Next
    text = CStr(value)
    text = Replace(text, vbCrLf, " ")
    text = Replace(text, vbCr, " ")
    text = Replace(text, vbLf, " ")
    If LCase(CStr(formatName)) = "tsv" Then
        text = Replace(text, vbTab, " ")
        EscapeExportField = text
        Exit Function
    End If
    EscapeExportField = Host.CsvEscape(text)
    If Err.Number <> 0 Or (Len(CStr(EscapeExportField)) = 0 And Len(text) > 0) Then
        Err.Clear
        text = Replace(text, """", """""")
        If InStr(1, text, ",", vbBinaryCompare) > 0 Or InStr(1, text, """", vbBinaryCompare) > 0 Then
            EscapeExportField = """" & text & """"
        Else
            EscapeExportField = text
        End If
    End If
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
