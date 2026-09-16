' 函数名: HostExcelTwoTableReconcilePack
' 描述: 两表对账场景包：按主键预览匹配/缺失/重复，可选字段差异；确认后只写对账明细、可选字段差异与处理摘要，不改源表
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
    Main = HostExcelTwoTableReconcilePack(appObj)
End Function

Function HostExcelTwoTableReconcilePack(appObj)
    On Error Resume Next
    Dim sourceRange, sourceSheet, workbook
    Set sourceRange = appObj.Selection.CurrentRegion
    Set sourceSheet = sourceRange.Worksheet
    Set workbook = sourceSheet.Parent
    If Err.Number <> 0 Or TypeName(sourceRange) = "Empty" Or TypeName(sourceRange) = "Nothing" Then
        Err.Clear
        HostExcelTwoTableReconcilePack = FailureJson("E_NO_SOURCE_RANGE", "请先在来源表选中带表头的数据区域")
        Exit Function
    End If
    If sourceRange.Rows.Count < 2 Then
        HostExcelTwoTableReconcilePack = FailureJson("E_RANGE_TOO_SMALL", "来源区域至少需要一行表头和一行数据")
        Exit Function
    End If

    Dim targetSheetName, targetSheet, targetRange
    targetSheetName = Trim(SafePrompt("目标工作表名称", FindOtherSheetName(workbook, sourceSheet.Name)))
    Set targetSheet = FindWorksheet(workbook, targetSheetName)
    If IsNothing(targetSheet) Then
        HostExcelTwoTableReconcilePack = FailureJson("E_TARGET_SHEET", "未找到目标工作表：" & targetSheetName)
        Exit Function
    End If
    If StrComp(sourceSheet.Name, targetSheet.Name, vbTextCompare) = 0 Then
        HostExcelTwoTableReconcilePack = FailureJson("E_TARGET_SHEET", "请选择与来源表不同的目标工作表")
        Exit Function
    End If
    Set targetRange = targetSheet.UsedRange
    If targetRange.Rows.Count < 2 Then
        HostExcelTwoTableReconcilePack = FailureJson("E_TARGET_RANGE", "目标工作表没有可核对的数据")
        Exit Function
    End If

    Dim sourceKeyName, targetKeyName, sourceKeyCol, targetKeyCol, outputMode
    sourceKeyName = Trim(SafePrompt("来源表主键字段", CStr(sourceRange.Cells(1, 1).Text)))
    targetKeyName = Trim(SafePrompt("目标表主键字段", sourceKeyName))
    sourceKeyCol = FindHeaderColumn(sourceRange, sourceKeyName)
    targetKeyCol = FindHeaderColumn(targetRange, targetKeyName)
    If sourceKeyCol <= 0 Then
        HostExcelTwoTableReconcilePack = FailureJson("E_SOURCE_KEY", "来源表未找到主键字段：" & sourceKeyName)
        Exit Function
    End If
    If targetKeyCol <= 0 Then
        HostExcelTwoTableReconcilePack = FailureJson("E_TARGET_KEY", "目标表未找到主键字段：" & targetKeyName)
        Exit Function
    End If
    outputMode = NormalizeOutputMode(SafePrompt("输出模式：NewSheet 或 NewWorkbook（默认 NewSheet）", "NewSheet"))

    Dim compareEnabled, fieldText, sourceFields, targetFields, fieldCount, fieldError
    compareEnabled = False
    fieldCount = 0
    fieldText = Trim(SafePrompt("可选：要比对的字段（逗号分隔；留空=不做字段差异；输 all=两表同名非主键字段）", ""))
    If Len(fieldText) > 0 Then
        compareEnabled = True
        fieldError = BuildCompareFields(sourceRange, targetRange, sourceKeyCol, targetKeyCol, fieldText, sourceFields, targetFields, fieldCount)
        If Len(fieldError) > 0 Then
            HostExcelTwoTableReconcilePack = FailureJson("E_COMPARE_FIELDS", fieldError)
            Exit Function
        End If
    End If

    Dim sourceRow, targetRow, keyText, matchedCount, targetMissingCount, sourceMissingCount, duplicateKeyCount
    Dim differenceRows, differenceCells, rowDifferences, sourceDup, targetDup
    matchedCount = 0
    targetMissingCount = 0
    sourceMissingCount = 0
    duplicateKeyCount = 0
    differenceRows = 0
    differenceCells = 0
    For sourceRow = 2 To sourceRange.Rows.Count
        keyText = NormalizeKey(sourceRange.Cells(sourceRow, sourceKeyCol).Text)
        If Len(keyText) > 0 Then
            targetRow = FindKeyRow(targetRange, targetKeyCol, keyText)
            sourceDup = CountKeyOccurrences(sourceRange, sourceKeyCol, keyText)
            If targetRow > 0 Then matchedCount = matchedCount + 1 Else targetMissingCount = targetMissingCount + 1
            If sourceDup > 1 Then duplicateKeyCount = duplicateKeyCount + 1
            If compareEnabled And targetRow > 0 And sourceDup = 1 Then
                If CountKeyOccurrences(targetRange, targetKeyCol, keyText) = 1 Then
                    rowDifferences = CountFieldDifferences(sourceRange, targetRange, sourceRow, targetRow, sourceFields, targetFields, fieldCount)
                    If rowDifferences > 0 Then
                        differenceRows = differenceRows + 1
                        differenceCells = differenceCells + rowDifferences
                    End If
                End If
            End If
        End If
    Next
    For targetRow = 2 To targetRange.Rows.Count
        keyText = NormalizeKey(targetRange.Cells(targetRow, targetKeyCol).Text)
        If Len(keyText) > 0 And FindKeyRow(sourceRange, sourceKeyCol, keyText) <= 0 Then sourceMissingCount = sourceMissingCount + 1
    Next

    Dim previewText, planId, planPreview
    previewText = "Excel 两表对账场景包预览（尚未写入）" & vbCrLf & _
        "命令ID：excel.two_table_reconcile_pack" & vbCrLf & _
        "来源=" & sourceSheet.Name & sourceRange.Address & "；主键=" & sourceKeyName & vbCrLf & _
        "目标=" & targetSheet.Name & targetRange.Address & "；主键=" & targetKeyName & vbCrLf & _
        "输出模式=" & outputMode & vbCrLf & _
        "匹配=" & CStr(matchedCount) & "；目标缺失=" & CStr(targetMissingCount) & "；来源缺失=" & CStr(sourceMissingCount) & "；来源重复键行=" & CStr(duplicateKeyCount) & vbCrLf
    If compareEnabled Then
        previewText = previewText & "字段差异行=" & CStr(differenceRows) & "；差异单元格=" & CStr(differenceCells) & "；对比字段=" & BuildFieldLabels(sourceRange, sourceFields, fieldCount) & vbCrLf
    Else
        previewText = previewText & "字段差异=关闭" & vbCrLf
    End If
    previewText = previewText & "将生成：对账明细"
    If compareEnabled Then previewText = previewText & " + 字段差异"
    previewText = previewText & " + 处理摘要；sourceUnchanged=true；不修改两张源表"

    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_two_table_reconcile_pack", "{""sourceSheet"":""" & EscapeJson(sourceSheet.Name) & """,""targetSheet"":""" & EscapeJson(targetSheet.Name) & """,""matched"":" & CStr(matchedCount) & ",""targetMissing"":" & CStr(targetMissingCount) & ",""sourceMissing"":" & CStr(sourceMissingCount) & ",""duplicateKeyRows"":" & CStr(duplicateKeyCount) & ",""differenceRows"":" & CStr(differenceRows) & ",""compareEnabled"":" & LCase(CStr(compareEnabled)) & "}", "office.excel.twoTableReconcilePack"
    planPreview = SafePreviewWritePlan()
    SafeRollbackWritePlan planId

    If Not SafeConfirmStep(previewText & vbCrLf & vbCrLf & "确认生成对账明细与摘要？", "excel.two_table_reconcile_pack") Then
        HostExcelTwoTableReconcilePack = FailureJson("E_CONFIRM_REQUIRED", "用户取消或未确认，未修改工作簿")
        Exit Function
    End If

    Dim createdSheets(), createdCount, detailSheet, diffSheet, summarySheet
    Dim outputRow, statusText, detailRows, diffOutRows
    createdCount = 0
    ReDim createdSheets(6)
    detailRows = 0
    diffOutRows = 0

    Set detailSheet = CreateOutputSheet(appObj, sourceSheet.Name, "对账明细", outputMode)
    If detailSheet Is Nothing Then
        HostExcelTwoTableReconcilePack = FailureJson("E_OUTPUT_SHEET", "无法创建对账明细工作表")
        Exit Function
    End If
    createdCount = createdCount + 1: Set createdSheets(createdCount) = detailSheet

    detailSheet.Cells(1, 1).Value = "主键"
    detailSheet.Cells(1, 2).Value = "状态"
    detailSheet.Cells(1, 3).Value = "来源行"
    detailSheet.Cells(1, 4).Value = "目标行"
    detailSheet.Cells(1, 5).Value = "来源重复次"
    detailSheet.Cells(1, 6).Value = "目标重复次"
    outputRow = 1

    For sourceRow = 2 To sourceRange.Rows.Count
        keyText = NormalizeKey(sourceRange.Cells(sourceRow, sourceKeyCol).Text)
        outputRow = outputRow + 1
        targetRow = 0
        sourceDup = 0
        targetDup = 0
        If Len(keyText) = 0 Then
            statusText = "空主键"
        Else
            targetRow = FindKeyRow(targetRange, targetKeyCol, keyText)
            sourceDup = CountKeyOccurrences(sourceRange, sourceKeyCol, keyText)
            targetDup = CountKeyOccurrences(targetRange, targetKeyCol, keyText)
            If targetRow > 0 Then statusText = "已匹配" Else statusText = "目标缺失"
            If sourceDup > 1 Then statusText = statusText & ";来源重复"
            If targetDup > 1 Then statusText = statusText & ";目标重复"
        End If
        detailSheet.Cells(outputRow, 1).Value = keyText
        detailSheet.Cells(outputRow, 2).Value = statusText
        detailSheet.Cells(outputRow, 3).Value = sourceRange.Row + sourceRow - 1
        detailSheet.Cells(outputRow, 4).Value = TargetAbsoluteRow(targetRange, targetRow)
        detailSheet.Cells(outputRow, 5).Value = sourceDup
        detailSheet.Cells(outputRow, 6).Value = targetDup
    Next
    For targetRow = 2 To targetRange.Rows.Count
        keyText = NormalizeKey(targetRange.Cells(targetRow, targetKeyCol).Text)
        If Len(keyText) > 0 And FindKeyRow(sourceRange, sourceKeyCol, keyText) <= 0 Then
            outputRow = outputRow + 1
            detailSheet.Cells(outputRow, 1).Value = keyText
            detailSheet.Cells(outputRow, 2).Value = "来源缺失"
            detailSheet.Cells(outputRow, 3).Value = 0
            detailSheet.Cells(outputRow, 4).Value = TargetAbsoluteRow(targetRange, targetRow)
            detailSheet.Cells(outputRow, 5).Value = 0
            detailSheet.Cells(outputRow, 6).Value = CountKeyOccurrences(targetRange, targetKeyCol, keyText)
        End If
    Next
    detailRows = outputRow - 1
    detailSheet.Rows(1).Font.Bold = True
    detailSheet.Columns.AutoFit
    If Err.Number <> 0 Then
        Err.Clear
        RollbackCreatedSheets appObj, createdSheets, createdCount
        HostExcelTwoTableReconcilePack = FailureJson("E_RECONCILE_WRITE", "对账明细写入失败，已删除未完成输出")
        Exit Function
    End If

    If compareEnabled Then
        Set diffSheet = CreateOutputSheet(appObj, sourceSheet.Name, "字段差异", outputMode)
        If diffSheet Is Nothing Then
            RollbackCreatedSheets appObj, createdSheets, createdCount
            HostExcelTwoTableReconcilePack = FailureJson("E_OUTPUT_SHEET", "无法创建字段差异工作表")
            Exit Function
        End If
        createdCount = createdCount + 1: Set createdSheets(createdCount) = diffSheet
        diffSheet.Cells(1, 1).Value = "主键"
        diffSheet.Cells(1, 2).Value = "状态"
        diffSheet.Cells(1, 3).Value = "字段"
        diffSheet.Cells(1, 4).Value = "来源值"
        diffSheet.Cells(1, 5).Value = "目标值"
        diffSheet.Cells(1, 6).Value = "来源行"
        diffSheet.Cells(1, 7).Value = "目标行"
        outputRow = 1
        For sourceRow = 2 To sourceRange.Rows.Count
            keyText = NormalizeKey(sourceRange.Cells(sourceRow, sourceKeyCol).Text)
            targetRow = FindKeyRow(targetRange, targetKeyCol, keyText)
            If Len(keyText) = 0 Then
                outputRow = outputRow + 1
                WriteDifferenceRow diffSheet, outputRow, "", "空主键未比较", "", "", "", sourceRange.Row + sourceRow - 1, 0
            ElseIf CountKeyOccurrences(sourceRange, sourceKeyCol, keyText) > 1 Or CountKeyOccurrences(targetRange, targetKeyCol, keyText) > 1 Then
                outputRow = outputRow + 1
                WriteDifferenceRow diffSheet, outputRow, keyText, "重复主键未比较", "", "", "", sourceRange.Row + sourceRow - 1, TargetAbsoluteRow(targetRange, targetRow)
            ElseIf targetRow <= 0 Then
                outputRow = outputRow + 1
                WriteDifferenceRow diffSheet, outputRow, keyText, "目标缺失", "", "", "", sourceRange.Row + sourceRow - 1, 0
            Else
                WriteFieldDifferences diffSheet, outputRow, keyText, sourceRange, targetRange, sourceRow, targetRow, sourceFields, targetFields, fieldCount
            End If
        Next
        For targetRow = 2 To targetRange.Rows.Count
            keyText = NormalizeKey(targetRange.Cells(targetRow, targetKeyCol).Text)
            If Len(keyText) > 0 And CountKeyOccurrences(targetRange, targetKeyCol, keyText) = 1 And FindKeyRow(sourceRange, sourceKeyCol, keyText) <= 0 Then
                outputRow = outputRow + 1
                WriteDifferenceRow diffSheet, outputRow, keyText, "来源缺失", "", "", "", 0, TargetAbsoluteRow(targetRange, targetRow)
            End If
        Next
        diffOutRows = outputRow - 1
        diffSheet.Rows(1).Font.Bold = True
        diffSheet.Columns.AutoFit
        If Err.Number <> 0 Then
            Err.Clear
            RollbackCreatedSheets appObj, createdSheets, createdCount
            HostExcelTwoTableReconcilePack = FailureJson("E_DIFF_WRITE", "字段差异写入失败，已删除未完成输出")
            Exit Function
        End If
    End If

    Set summarySheet = CreateOutputSheet(appObj, sourceSheet.Name, "处理摘要", outputMode)
    If summarySheet Is Nothing Then
        RollbackCreatedSheets appObj, createdSheets, createdCount
        HostExcelTwoTableReconcilePack = FailureJson("E_OUTPUT_SHEET", "无法创建处理摘要工作表")
        Exit Function
    End If
    createdCount = createdCount + 1: Set createdSheets(createdCount) = summarySheet

    Dim summaryLines, impactJson, summary, outputNames
    outputNames = detailSheet.Name
    If compareEnabled Then outputNames = outputNames & "," & diffSheet.Name
    outputNames = outputNames & "," & summarySheet.Name
    summaryLines = "场景名=两表对账场景包" & vbCrLf & _
        "命令ID=excel.two_table_reconcile_pack" & vbCrLf & _
        "源表名称=" & sourceSheet.Name & vbCrLf & _
        "目标表名称=" & targetSheet.Name & vbCrLf & _
        "源区域=" & sourceRange.Address & vbCrLf & _
        "输出模式=" & outputMode & vbCrLf & _
        "输出位置=" & outputNames & vbCrLf & _
        "处理前行数=" & CStr(sourceRange.Rows.Count - 1) & vbCrLf & _
        "处理后行数=" & CStr(detailRows) & vbCrLf & _
        "参数=sourceKey=" & sourceKeyName & ";targetKey=" & targetKeyName & ";compareEnabled=" & CStr(compareEnabled) & vbCrLf & _
        "匹配行数=" & CStr(matchedCount) & vbCrLf & _
        "目标缺失=" & CStr(targetMissingCount) & vbCrLf & _
        "来源缺失=" & CStr(sourceMissingCount) & vbCrLf & _
        "重复键行=" & CStr(duplicateKeyCount) & vbCrLf & _
        "字段差异行=" & CStr(differenceRows) & vbCrLf & _
        "差异单元格=" & CStr(differenceCells) & vbCrLf & _
        "执行时间=" & Now & vbCrLf & _
        "结果=成功"
    impactJson = SafeHostText("ExcelWriteImpactSummary", summaryLines, summarySheet.Name)
    If Not ExtractJsonBoolean(impactJson, "ok") Then WriteSummaryFallback summarySheet, summaryLines

    summary = "Excel 两表对账完成：明细=" & detailSheet.Name
    If compareEnabled Then summary = summary & "；字段差异=" & diffSheet.Name
    summary = summary & "；匹配=" & CStr(matchedCount) & "；目标缺失=" & CStr(targetMissingCount) & "；来源缺失=" & CStr(sourceMissingCount) & "；sourceUnchanged=true"
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostExcelTwoTableReconcilePack = "{""ok"":true,""sourceUnchanged"":true,""commandId"":""excel.two_table_reconcile_pack"",""outputSheet"":""" & EscapeJson(detailSheet.Name) & """,""summarySheet"":""" & EscapeJson(summarySheet.Name) & """,""matched"":" & CStr(matchedCount) & ",""targetMissing"":" & CStr(targetMissingCount) & ",""sourceMissing"":" & CStr(sourceMissingCount) & ",""duplicateKeyRows"":" & CStr(duplicateKeyCount) & ",""differenceRows"":" & CStr(differenceRows) & ",""writePlanPreview"":""" & EscapeJson(planPreview) & """,""message"":""" & EscapeJson(summary) & """}"
End Function

Function BuildCompareFields(sourceRange, targetRange, sourceKeyCol, targetKeyCol, fieldText, ByRef sourceFields, ByRef targetFields, ByRef fieldCount)
    Dim names, i, sourceCol, targetCol, n, label
    BuildCompareFields = ""
    fieldCount = 0
    If StrComp(Trim(CStr(fieldText)), "all", vbTextCompare) = 0 Or StrComp(Trim(CStr(fieldText)), "全部", vbTextCompare) = 0 Or Len(Trim(CStr(fieldText))) = 0 Then
        n = 0
        ReDim sourceFields(sourceRange.Columns.Count)
        ReDim targetFields(sourceRange.Columns.Count)
        For i = 1 To sourceRange.Columns.Count
            If i <> sourceKeyCol Then
                label = NormalizeCellText(sourceRange.Cells(1, i).Text)
                targetCol = FindHeaderColumn(targetRange, label)
                If targetCol > 0 And targetCol <> targetKeyCol Then
                    n = n + 1
                    sourceFields(n) = i
                    targetFields(n) = targetCol
                End If
            End If
        Next
        If n = 0 Then
            BuildCompareFields = "未找到可对比的同名非主键字段"
            Exit Function
        End If
        ReDim Preserve sourceFields(n)
        ReDim Preserve targetFields(n)
        fieldCount = n
        Exit Function
    End If

    names = SplitCsvFields(fieldText)
    If (Not IsArray(names)) Then
        BuildCompareFields = "对比字段为空"
        Exit Function
    End If
    On Error Resume Next
    If UBound(names) < 0 Then
        BuildCompareFields = "对比字段为空"
        Exit Function
    End If
    Err.Clear
    n = 0
    ReDim sourceFields(UBound(names) + 1)
    ReDim targetFields(UBound(names) + 1)
    For i = 0 To UBound(names)
        sourceCol = FindHeaderColumn(sourceRange, names(i))
        targetCol = FindHeaderColumn(targetRange, names(i))
        If sourceCol <= 0 Or targetCol <= 0 Then
            BuildCompareFields = "未在两表找到对比字段：" & names(i)
            Exit Function
        End If
        n = n + 1
        sourceFields(n) = sourceCol
        targetFields(n) = targetCol
    Next
    ReDim Preserve sourceFields(n)
    ReDim Preserve targetFields(n)
    fieldCount = n
End Function

Function CountFieldDifferences(sourceRange, targetRange, sourceRow, targetRow, sourceFields, targetFields, fieldCount)
    Dim i, total
    total = 0
    For i = 1 To fieldCount
        If Not ValuesEqual(sourceRange.Cells(sourceRow, sourceFields(i)).Text, targetRange.Cells(targetRow, targetFields(i)).Text) Then total = total + 1
    Next
    CountFieldDifferences = total
End Function

Sub WriteFieldDifferences(outputSheet, ByRef outputRow, keyText, sourceRange, targetRange, sourceRow, targetRow, sourceFields, targetFields, fieldCount)
    Dim i, fieldName, sourceValue, targetValue
    For i = 1 To fieldCount
        sourceValue = CStr(sourceRange.Cells(sourceRow, sourceFields(i)).Text)
        targetValue = CStr(targetRange.Cells(targetRow, targetFields(i)).Text)
        If Not ValuesEqual(sourceValue, targetValue) Then
            fieldName = NormalizeCellText(sourceRange.Cells(1, sourceFields(i)).Text)
            outputRow = outputRow + 1
            WriteDifferenceRow outputSheet, outputRow, keyText, "字段不一致", fieldName, sourceValue, targetValue, sourceRange.Row + sourceRow - 1, TargetAbsoluteRow(targetRange, targetRow)
        End If
    Next
End Sub

Sub WriteDifferenceRow(sheetObj, rowIndex, keyText, statusText, fieldName, sourceValue, targetValue, sourceRow, targetRow)
    sheetObj.Cells(rowIndex, 1).Value = keyText
    sheetObj.Cells(rowIndex, 2).Value = statusText
    sheetObj.Cells(rowIndex, 3).Value = fieldName
    sheetObj.Cells(rowIndex, 4).Value = sourceValue
    sheetObj.Cells(rowIndex, 5).Value = targetValue
    sheetObj.Cells(rowIndex, 6).Value = sourceRow
    sheetObj.Cells(rowIndex, 7).Value = targetRow
End Sub

Function BuildFieldLabels(sourceRange, sourceFields, fieldCount)
    Dim i, buf
    buf = ""
    For i = 1 To fieldCount
        If Len(buf) > 0 Then buf = buf & ","
        buf = buf & NormalizeCellText(sourceRange.Cells(1, sourceFields(i)).Text)
    Next
    BuildFieldLabels = buf
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

