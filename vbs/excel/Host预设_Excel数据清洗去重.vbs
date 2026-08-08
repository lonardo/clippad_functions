' 函数名: HostExcelCleanCurrentRegion
' 描述: 表级清洗去重：规范化文本/去掉空行/按整行去重，结果写入新工作表副本（不改源表）
' 适用应用: Excel
' 搜索范围: 当前区域
' 搜索对象: 低
'
' 实现约定（供后续维护/生成使用）：
' - 只读取 Selection.CurrentRegion；按显示文本规范化、跳过全空行，并以规范化后的整行作为去重键保留首行。
' - 只有明确输入“生成”才创建结果表；源表不改，结果表保留“_来源行”以便追溯。
' - 写入计划只用于预览元数据，不能回滚直接 Excel COM 写入；发生写入错误时仅删除本次未完成的结果表。

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
    Main = HostExcelCleanCurrentRegion(appObj)
End Function

Function HostExcelCleanCurrentRegion(appObj)
    On Error Resume Next

    Dim sourceRange
    Set sourceRange = appObj.Selection.CurrentRegion
    If Err.Number <> 0 Or TypeName(sourceRange) = "Empty" Or TypeName(sourceRange) = "Nothing" Then
        Err.Clear
        HostExcelCleanCurrentRegion = FailureJson("E_NO_RANGE", "请先选中带表头的数据区域")
        Exit Function
    End If
    If sourceRange.Rows.Count < 2 Or sourceRange.Columns.Count < 1 Then
        HostExcelCleanCurrentRegion = FailureJson("E_RANGE_TOO_SMALL", "当前区域至少需要一行表头和一行数据")
        Exit Function
    End If

    Dim beforeRows, emptyRows, uniqueRows, duplicateRows
    Dim rowIndex, colIndex, key, cellText, isEmpty
    Dim keys(), keyCount
    beforeRows = sourceRange.Rows.Count - 1
    emptyRows = 0
    uniqueRows = 0
    duplicateRows = 0
    keyCount = 0

    For rowIndex = 2 To sourceRange.Rows.Count
        isEmpty = True
        key = ""
        For colIndex = 1 To sourceRange.Columns.Count
            cellText = NormalizeCellText(sourceRange.Cells(rowIndex, colIndex).Text)
            If Len(cellText) > 0 Then isEmpty = False
            If colIndex > 1 Then key = key & "|"
            key = key & LCase(cellText)
        Next
        If isEmpty Then
            emptyRows = emptyRows + 1
        ElseIf KeySeen(keys, keyCount, key) Then
            duplicateRows = duplicateRows + 1
        Else
            ReDim Preserve keys(keyCount)
            keys(keyCount) = key
            keyCount = keyCount + 1
            uniqueRows = uniqueRows + 1
        End If
    Next

    Dim preflightJson, statsJson, previewText, planId, planPreview
    preflightJson = SafeHostText("GetRunPreflightPlan")
    statsJson = SafeHostText("GetTypeStats")
    previewText = "Excel 数据清洗去重预检（尚未写入）" & vbCrLf & _
        "源表=" & sourceRange.Worksheet.Name & "，区域=" & sourceRange.Address & vbCrLf & _
        "数据行=" & CStr(beforeRows) & "，空行=" & CStr(emptyRows) & "，重复行=" & CStr(duplicateRows) & "，保留行=" & CStr(uniqueRows) & vbCrLf & _
        "步骤=NormalizeText -> DropEmptyRows -> Deduplicate(first)" & vbCrLf & _
        "输出=新工作表；源数据不变" & vbCrLf & _
        "Host.GetTypeStats: " & statsJson & vbCrLf & _
        "Host.GetRunPreflightPlan: " & preflightJson

    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_clean_dedupe_copy", "{""beforeRows"":" & CStr(beforeRows) & ",""keepRows"":" & CStr(uniqueRows) & "}", "office.excel.cleanDedupe"
    planPreview = SafePreviewWritePlan()
    SafeRollbackWritePlan planId

    Dim confirmation
    confirmation = Trim(SafePrompt(previewText & vbCrLf & vbCrLf & "确认生成清洗去重副本，请输入：生成", ""))
    If StrComp(confirmation, "生成", vbTextCompare) <> 0 Then
        HostExcelCleanCurrentRegion = FailureJson("E_CONFIRM_REQUIRED", "未输入“生成”，已取消且未修改工作簿")
        Exit Function
    End If

    Dim workbook, outputSheet, outputName, outputRow, written
    Dim seenKeys(), seenCount
    Set workbook = sourceRange.Worksheet.Parent
    outputName = UniqueSheetName(workbook, "清洗去重")
    Set outputSheet = workbook.Worksheets.Add
    outputSheet.Name = outputName
    If Err.Number <> 0 Then
        Err.Clear
        SafeDeleteSheet appObj, outputSheet
        HostExcelCleanCurrentRegion = FailureJson("E_OUTPUT_SHEET", "无法创建清洗去重结果表")
        Exit Function
    End If

    For colIndex = 1 To sourceRange.Columns.Count
        outputSheet.Cells(1, colIndex).Value = NormalizeCellText(sourceRange.Cells(1, colIndex).Text)
    Next
    outputSheet.Cells(1, sourceRange.Columns.Count + 1).Value = "_来源行"

    outputRow = 1
    written = 0
    seenCount = 0
    For rowIndex = 2 To sourceRange.Rows.Count
        isEmpty = True
        key = ""
        For colIndex = 1 To sourceRange.Columns.Count
            cellText = NormalizeCellText(sourceRange.Cells(rowIndex, colIndex).Text)
            If Len(cellText) > 0 Then isEmpty = False
            If colIndex > 1 Then key = key & "|"
            key = key & LCase(cellText)
        Next
        If isEmpty Then
            ' drop empty
        ElseIf KeySeen(seenKeys, seenCount, key) Then
            ' drop duplicate
        Else
            ReDim Preserve seenKeys(seenCount)
            seenKeys(seenCount) = key
            seenCount = seenCount + 1
            outputRow = outputRow + 1
            For colIndex = 1 To sourceRange.Columns.Count
                outputSheet.Cells(outputRow, colIndex).Value = NormalizeCellText(sourceRange.Cells(rowIndex, colIndex).Text)
            Next
            outputSheet.Cells(outputRow, sourceRange.Columns.Count + 1).Value = sourceRange.Row + rowIndex - 1
            written = written + 1
        End If
    Next

    outputSheet.Rows(1).Font.Bold = True
    outputSheet.Columns.AutoFit

    If Err.Number <> 0 Then
        Dim writeError
        writeError = Err.Description
        Err.Clear
        SafeDeleteSheet appObj, outputSheet
        HostExcelCleanCurrentRegion = FailureJson("E_COPY_FAILED", "清洗去重写入失败，已删除未完成结果表：" & writeError)
        Exit Function
    End If

    Dim summary
    summary = "Excel 数据清洗去重完成：结果表=" & outputName & "，清洗前=" & CStr(beforeRows) & "，清洗后=" & CStr(written) & "，源表未修改"
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostExcelCleanCurrentRegion = "{""ok"":true,""sourceUnchanged"":true,""outputSheet"":""" & EscapeJson(outputName) & _
        """,""beforeRows"":" & CStr(beforeRows) & ",""afterRows"":" & CStr(written) & ",""writePlanPreview"":""" & EscapeJson(planPreview) & _
        """,""message"":""" & EscapeJson(summary) & """}"
End Function

Function KeySeen(ByRef seenKeys, seenCount, keyText)
    Dim i
    KeySeen = False
    For i = 0 To seenCount - 1
        If StrComp(seenKeys(i), keyText, vbBinaryCompare) = 0 Then
            KeySeen = True
            Exit Function
        End If
    Next
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
    Do While InStr(1, text, "  ", vbBinaryCompare) > 0
        text = Replace(text, "  ", " ")
    Loop
    NormalizeCellText = Trim(text)
End Function

Function SafeHostText(methodName)
    On Error Resume Next
    If methodName = "GetRunPreflightPlan" Then
        SafeHostText = Host.GetRunPreflightPlan("Excel", True, True, False)
    ElseIf methodName = "GetTypeStats" Then
        SafeHostText = Host.GetTypeStats()
    Else
        SafeHostText = ""
    End If
    If Err.Number <> 0 Then
        SafeHostText = ""
        Err.Clear
    End If
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
