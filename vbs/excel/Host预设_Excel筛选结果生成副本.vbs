' 函数名: HostExcelCreateFilteredCopy
' 描述: 按表头和关键词筛选当前区域，先预览命中行数，确认后把结果写入新工作表；不隐藏、删除或改写源数据
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
    Main = HostExcelCreateFilteredCopy(appObj)
End Function

Function HostExcelCreateFilteredCopy(appObj)
    On Error Resume Next
    Dim sourceRange
    Set sourceRange = appObj.Selection.CurrentRegion
    If Err.Number <> 0 Or TypeName(sourceRange) = "Empty" Or TypeName(sourceRange) = "Nothing" Then
        Err.Clear
        HostExcelCreateFilteredCopy = FailureJson("E_NO_RANGE", "请先选中带表头的数据区域")
        Exit Function
    End If
    If sourceRange.Rows.Count < 2 Or sourceRange.Columns.Count < 1 Then
        HostExcelCreateFilteredCopy = FailureJson("E_RANGE_TOO_SMALL", "当前区域至少需要一行表头和一行数据")
        Exit Function
    End If

    Dim fieldName, keyword, matchMode, fieldColumn
    fieldName = Trim(SafePrompt("筛选字段名称（必须与表头一致）", CStr(sourceRange.Cells(1, 1).Text)))
    keyword = SafePrompt("筛选关键词", "")
    matchMode = NormalizeMatchMode(SafePrompt("匹配方式：精确 或 包含", "精确"))
    fieldColumn = FindHeaderColumn(sourceRange, fieldName)
    If fieldColumn <= 0 Then
        HostExcelCreateFilteredCopy = FailureJson("E_FIELD_NOT_FOUND", "未找到筛选字段：" & fieldName)
        Exit Function
    End If

    Dim rowIndex, matchCount, cellText
    matchCount = 0
    For rowIndex = 2 To sourceRange.Rows.Count
        cellText = CStr(sourceRange.Cells(rowIndex, fieldColumn).Text)
        If IsMatch(cellText, keyword, matchMode) Then matchCount = matchCount + 1
    Next

    Dim previewText, planId, planPreview
    previewText = "Excel 筛选交付预览（尚未写入）" & vbCrLf & _
        "源工作表：" & sourceRange.Worksheet.Name & "；区域=" & sourceRange.Address & vbCrLf & _
        "字段：" & fieldName & "；关键词：" & keyword & "；方式：" & MatchModeLabel(matchMode) & vbCrLf & _
        "数据行=" & CStr(sourceRange.Rows.Count - 1) & "；命中=" & CStr(matchCount) & "；输出=新工作表；源数据不变"
    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_filtered_copy", "{""source"":""" & EscapeJson(sourceRange.Address) & """,""field"":""" & EscapeJson(fieldName) & """,""keyword"":""" & EscapeJson(keyword) & """,""matches"":" & CStr(matchCount) & "}", "office.excel.filteredCopy"
    planPreview = SafePreviewWritePlan()
    SafeRollbackWritePlan planId

    Dim confirmation
    confirmation = Trim(SafePrompt(previewText & vbCrLf & vbCrLf & "确认生成筛选副本，请输入：生成", ""))
    If StrComp(confirmation, "生成", vbTextCompare) <> 0 Then
        HostExcelCreateFilteredCopy = FailureJson("E_CONFIRM_REQUIRED", "未输入“生成”，已取消且未修改工作簿")
        Exit Function
    End If

    Dim workbook, outputSheet, outputName
    Set workbook = sourceRange.Worksheet.Parent
    outputName = UniqueSheetName(workbook, "筛选结果")
    Set outputSheet = workbook.Worksheets.Add
    outputSheet.Name = outputName
    If Err.Number <> 0 Then
        Err.Clear
        SafeDeleteSheet appObj, outputSheet
        HostExcelCreateFilteredCopy = FailureJson("E_OUTPUT_SHEET", "无法创建筛选结果工作表")
        Exit Function
    End If

    Dim colIndex, outputRow
    outputRow = 1
    For colIndex = 1 To sourceRange.Columns.Count
        outputSheet.Cells(outputRow, colIndex).Value = sourceRange.Cells(1, colIndex).Value
        outputSheet.Cells(outputRow, colIndex).NumberFormat = sourceRange.Cells(1, colIndex).NumberFormat
    Next
    outputSheet.Cells(outputRow, sourceRange.Columns.Count + 1).Value = "_来源行"
    For rowIndex = 2 To sourceRange.Rows.Count
        cellText = CStr(sourceRange.Cells(rowIndex, fieldColumn).Text)
        If IsMatch(cellText, keyword, matchMode) Then
            outputRow = outputRow + 1
            For colIndex = 1 To sourceRange.Columns.Count
                outputSheet.Cells(outputRow, colIndex).Value = sourceRange.Cells(rowIndex, colIndex).Value
                outputSheet.Cells(outputRow, colIndex).NumberFormat = sourceRange.Cells(rowIndex, colIndex).NumberFormat
            Next
            outputSheet.Cells(outputRow, sourceRange.Columns.Count + 1).Value = sourceRange.Row + rowIndex - 1
        End If
    Next
    outputSheet.Rows(1).Font.Bold = True
    outputSheet.Columns.AutoFit

    If Err.Number <> 0 Then
        Dim writeError
        writeError = Err.Description
        Err.Clear
        SafeDeleteSheet appObj, outputSheet
        HostExcelCreateFilteredCopy = FailureJson("E_COPY_FAILED", "筛选副本写入失败，已删除未完成工作表：" & writeError)
        Exit Function
    End If

    Dim summary
    summary = "Excel 筛选副本已生成；工作表=" & outputName & "；命中行=" & CStr(matchCount) & "；源数据未修改"
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostExcelCreateFilteredCopy = "{""ok"":true,""sourceUnchanged"":true,""outputSheet"":""" & EscapeJson(outputName) & _
        """,""matchedRows"":" & CStr(matchCount) & ",""writePlanPreview"":""" & EscapeJson(planPreview) & _
        """,""message"":""" & EscapeJson(summary) & """}"
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

Function NormalizeMatchMode(value)
    Dim text
    text = LCase(Trim(CStr(value)))
    If text = "包含" Or text = "contains" Then NormalizeMatchMode = "contains" Else NormalizeMatchMode = "exact"
End Function

Function MatchModeLabel(value)
    If value = "contains" Then MatchModeLabel = "包含" Else MatchModeLabel = "精确"
End Function

Function IsMatch(cellText, keyword, matchMode)
    If matchMode = "contains" Then
        IsMatch = (InStr(1, CStr(cellText), CStr(keyword), vbTextCompare) > 0)
    Else
        IsMatch = (StrComp(Trim(CStr(cellText)), Trim(CStr(keyword)), vbTextCompare) = 0)
    End If
End Function

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
        If StrComp(CStr(sheetObj.Name), CStr(sheetName), vbTextCompare) = 0 Then SheetExists = True: Exit Function
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
