' 函数名: HostExcelReportCleanPipelinePreflight
' 描述: 可汇报清洗流水线：预检空白/重复/首尾空格/公式问题，可选生成去空白副本，并给出下一步推荐预设清单；源表默认不变
' 适用应用: Excel
' 搜索范围: 当前范围
' 搜索对象: 无

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
    Main = HostExcelReportCleanPipelinePreflight(appObj)
End Function

Function HostExcelReportCleanPipelinePreflight(appObj)
    On Error Resume Next

    Dim sourceRange
    Set sourceRange = appObj.Selection.CurrentRegion
    If Err.Number <> 0 Or TypeName(sourceRange) = "Empty" Or TypeName(sourceRange) = "Nothing" Then
        Err.Clear
        HostExcelReportCleanPipelinePreflight = FailureJson("E_NO_RANGE", "请先选中带表头的数据区域")
        Exit Function
    End If
    If sourceRange.Rows.Count < 2 Then
        HostExcelReportCleanPipelinePreflight = FailureJson("E_RANGE_TOO_SMALL", "当前区域至少需要一行表头和一行数据")
        Exit Function
    End If

    Dim blankCells, trimCells, multiSpaceCells, duplicateRows, errorCells, totalRows, totalCols
    blankCells = 0
    trimCells = 0
    multiSpaceCells = 0
    duplicateRows = 0
    errorCells = 0
    totalRows = sourceRange.Rows.Count - 1
    totalCols = sourceRange.Columns.Count
    ScanCleanIssues sourceRange, blankCells, trimCells, multiSpaceCells, duplicateRows, errorCells

    Dim headersJson, rangeSummary, formulaSummary, hasBlank, hasFormula
    headersJson = SafeHostTextExcel("GetHeaders")
    rangeSummary = SafeHostTextExcel("GetRangeSummary")
    formulaSummary = SafeHostTextExcel("GetExcelFormulaIssueSummary")
    hasBlank = SafeHostTextExcel("HasBlank")
    hasFormula = SafeHostTextExcel("HasFormula")

    Dim nextSteps
    nextSteps = "下一步推荐（请在指令列表继续运行）：" & vbCrLf & _
        "1) Host预设_Excel重复值处理包.vbs" & vbCrLf & _
        "2) Host预设_Excel多条件筛选生成副本.vbs" & vbCrLf & _
        "3) Host预设_Excel高级合并行生成副本.vbs" & vbCrLf & _
        "4) Host预设_Excel字段映射导出交付.vbs"

    Dim previewText
    previewText = "Excel 可汇报清洗流水线预检（尚未写入）" & vbCrLf & _
        "区域=" & sourceRange.Address & "；数据行=" & CStr(totalRows) & "；列=" & CStr(totalCols) & vbCrLf & _
        "空白单元格=" & CStr(blankCells) & "；需去首尾空格=" & CStr(trimCells) & _
        "；多余空格单元格=" & CStr(multiSpaceCells) & vbCrLf & _
        "疑似整行重复=" & CStr(duplicateRows) & "；错误值单元格=" & CStr(errorCells) & vbCrLf & _
        "Host.HasBlank=" & hasBlank & "；Host.HasFormula=" & hasFormula & vbCrLf & _
        "Host.GetExcelFormulaIssueSummary: " & formulaSummary & vbCrLf & _
        "Host.GetHeaders: " & headersJson & vbCrLf & _
        "Host.GetRangeSummary: " & rangeSummary & vbCrLf & _
        nextSteps

    Dim actionText, doCopy
    actionText = Trim(SafePrompt(previewText & vbCrLf & vbCrLf & "本步可选：仅摘要 / 生成去空白副本", "生成去空白副本"))
    doCopy = (InStr(1, actionText, "生成", vbTextCompare) > 0 Or InStr(1, actionText, "副本", vbTextCompare) > 0)

    If Not doCopy Then
        Host.WriteClipboard previewText
        SafeWriteLog previewText
        HostExcelReportCleanPipelinePreflight = "{""ok"":true,""sourceUnchanged"":true,""changed"":false,""message"":""" & EscapeJson(previewText) & """,""blankCells"":" & CStr(blankCells) & ",""duplicateRows"":" & CStr(duplicateRows) & "}"
        Exit Function
    End If

    Dim planId, planPreview
    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_report_clean_pipeline", "{""blankCells"":" & CStr(blankCells) & ",""trimCells"":" & CStr(trimCells) & "}", "office.excel.pipeline.clean"
    planPreview = SafePreviewWritePlan()
    SafeRollbackWritePlan planId

    If StrComp(Trim(SafePrompt("确认生成去空白清洗副本（源表不变），请输入：生成", "")), "生成", vbTextCompare) <> 0 Then
        HostExcelReportCleanPipelinePreflight = FailureJson("E_CONFIRM_REQUIRED", "未输入“生成”，已取消且未修改工作簿")
        Exit Function
    End If

    Dim workbook, outputSheet, outputName, r, c, outputRow, rawText, cleanText, changedCells
    Set workbook = sourceRange.Worksheet.Parent
    outputName = UniqueSheetName(workbook, "清洗流水线")
    Set outputSheet = workbook.Worksheets.Add
    outputSheet.Name = outputName
    If Err.Number <> 0 Then
        Err.Clear
        SafeDeleteSheet appObj, outputSheet
        HostExcelReportCleanPipelinePreflight = FailureJson("E_OUTPUT_SHEET", "无法创建清洗副本工作表")
        Exit Function
    End If

    For c = 1 To totalCols
        outputSheet.Cells(1, c).Value = sourceRange.Cells(1, c).Value
    Next
    outputSheet.Cells(1, totalCols + 1).Value = "_来源行"
    outputSheet.Cells(1, totalCols + 2).Value = "_本行清洗单元格数"

    outputRow = 1
    changedCells = 0
    For r = 2 To sourceRange.Rows.Count
        outputRow = outputRow + 1
        Dim rowChanged
        rowChanged = 0
        For c = 1 To totalCols
            rawText = CStr(sourceRange.Cells(r, c).Text)
            cleanText = CollapseSpaces(NormalizeCellText(rawText))
            If IsErrorValue(sourceRange.Cells(r, c)) Then
                outputSheet.Cells(outputRow, c).Value = sourceRange.Cells(r, c).Value
            ElseIf IsNumericCell(sourceRange.Cells(r, c)) And StrComp(rawText, cleanText, vbBinaryCompare) = 0 Then
                outputSheet.Cells(outputRow, c).Value = sourceRange.Cells(r, c).Value
                outputSheet.Cells(outputRow, c).NumberFormat = sourceRange.Cells(r, c).NumberFormat
            Else
                outputSheet.Cells(outputRow, c).Value = cleanText
                If StrComp(CStr(rawText), CStr(cleanText), vbBinaryCompare) <> 0 Then
                    rowChanged = rowChanged + 1
                    changedCells = changedCells + 1
                End If
            End If
        Next
        outputSheet.Cells(outputRow, totalCols + 1).Value = sourceRange.Row + r - 1
        outputSheet.Cells(outputRow, totalCols + 2).Value = rowChanged
    Next
    outputSheet.Rows(1).Font.Bold = True
    outputSheet.Columns.AutoFit

    Dim summary
    summary = "Excel 可汇报清洗流水线第1步完成；输出表=" & outputName & "；清洗单元格=" & CStr(changedCells) & "；源数据未修改" & vbCrLf & nextSteps & vbCrLf & "Preview: " & planPreview
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostExcelReportCleanPipelinePreflight = "{""ok"":true,""sourceUnchanged"":true,""outputSheet"":""" & EscapeJson(outputName) & """,""changedCells"":" & CStr(changedCells) & ",""message"":""" & EscapeJson(summary) & """}"
End Function

Sub ScanCleanIssues(sourceRange, ByRef blankCells, ByRef trimCells, ByRef multiSpaceCells, ByRef duplicateRows, ByRef errorCells)
    On Error Resume Next
    Dim r, c, textValue, seen(), seenCount, rowKey, i, isDup
    blankCells = 0
    trimCells = 0
    multiSpaceCells = 0
    duplicateRows = 0
    errorCells = 0
    seenCount = 0
    For r = 2 To sourceRange.Rows.Count
        rowKey = ""
        For c = 1 To sourceRange.Columns.Count
            If IsErrorValue(sourceRange.Cells(r, c)) Then
                errorCells = errorCells + 1
                textValue = "#ERR#"
            Else
                textValue = CStr(sourceRange.Cells(r, c).Text)
                If Len(NormalizeCellText(textValue)) = 0 Then
                    blankCells = blankCells + 1
                Else
                    If Len(textValue) <> Len(Trim(textValue)) Then trimCells = trimCells + 1
                    If InStr(textValue, "  ") > 0 Then multiSpaceCells = multiSpaceCells + 1
                End If
            End If
            rowKey = rowKey & vbVerticalTab & NormalizeCellText(textValue)
        Next
        isDup = False
        For i = 0 To seenCount - 1
            If StrComp(seen(i), rowKey, vbBinaryCompare) = 0 Then
                isDup = True
                Exit For
            End If
        Next
        If isDup Then
            duplicateRows = duplicateRows + 1
        Else
            ReDim Preserve seen(seenCount)
            seen(seenCount) = rowKey
            seenCount = seenCount + 1
        End If
    Next
End Sub

Function CollapseSpaces(text)
    Dim t, prev
    t = CStr(text)
    Do
        prev = t
        t = Replace(t, "  ", " ")
    Loop While t <> prev
    CollapseSpaces = Trim(t)
End Function

Function IsErrorValue(cell)
    On Error Resume Next
    IsErrorValue = False
    If IsError(cell.Value) Then IsErrorValue = True
    Err.Clear
End Function

Function IsNumericCell(cell)
    On Error Resume Next
    IsNumericCell = False
    If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) And VarType(cell.Value) <> vbString Then IsNumericCell = True
    Err.Clear
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

Function UniqueSheetName(workbook, baseName)
    Dim candidate, suffix
    candidate = Left(baseName, 31)
    suffix = 1
    Do While SheetExists(workbook, candidate)
        suffix = suffix + 1
        candidate = Left(baseName, 27) & "_" & CStr(suffix)
    Loop
    UniqueSheetName = candidate
End Function

Function SheetExists(workbook, sheetName)
    Dim sheetObj
    SheetExists = False
    For Each sheetObj In workbook.Worksheets
        If StrComp(CStr(sheetObj.Name), CStr(sheetName), vbTextCompare) = 0 Then
            SheetExists = True
            Exit Function
        End If
    Next
End Function

Sub SafeDeleteSheet(appObj, sheetObj)
    On Error Resume Next
    Dim oldAlerts
    oldAlerts = appObj.DisplayAlerts
    appObj.DisplayAlerts = False
    sheetObj.Delete
    appObj.DisplayAlerts = oldAlerts
    Err.Clear
End Sub

Function NormalizeCellText(value)
    Dim text
    text = CStr(value)
    text = Replace(text, ChrW(160), " ")
    text = Replace(text, ChrW(12288), " ")
    NormalizeCellText = Trim(text)
End Function

Function SafeHostTextExcel(methodName)
    On Error Resume Next
    If methodName = "GetHeaders" Then
        SafeHostTextExcel = Host.GetHeaders()
    ElseIf methodName = "GetRangeSummary" Then
        SafeHostTextExcel = Host.GetRangeSummary()
    ElseIf methodName = "GetExcelContextInfo" Then
        SafeHostTextExcel = Host.GetExcelContextInfo()
    ElseIf methodName = "GetExcelFormulaIssueSummary" Then
        SafeHostTextExcel = Host.GetExcelFormulaIssueSummary()
    ElseIf methodName = "HasBlank" Then
        SafeHostTextExcel = CStr(Host.HasBlank())
    ElseIf methodName = "HasFormula" Then
        SafeHostTextExcel = CStr(Host.HasFormula())
    Else
        SafeHostTextExcel = ""
    End If
    If Err.Number <> 0 Then SafeHostTextExcel = ""
    Err.Clear
End Function
