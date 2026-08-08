' 函数名: HostExcelExtractUniqueCopy
' 描述: 智能识别键字段，预检唯一值后一键确认，将首现唯一行写入新工作表副本并保留来源行；源表不变
' 适用应用: Excel
' 搜索范围: 当前区域
' 搜索对象: 低
'
' 实现约定（供后续维护/生成使用）：
' - 默认使用活动单元格所在列（否则首列）作为唯一键；用户可用“生成|字段名”或“生成|*”明确覆盖。
' - 只有明确输入“生成”才创建结果表；保留每个键的首现行、键值、重复次数和来源行，源表不改。
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
    Main = HostExcelExtractUniqueCopy(appObj)
End Function

Function HostExcelExtractUniqueCopy(appObj)
    On Error Resume Next

    Dim sourceRange
    Set sourceRange = appObj.Selection.CurrentRegion
    If Err.Number <> 0 Or TypeName(sourceRange) = "Empty" Or TypeName(sourceRange) = "Nothing" Then
        Err.Clear
        HostExcelExtractUniqueCopy = FailureJson("E_NO_RANGE", "请先选中带表头的数据区域")
        Exit Function
    End If
    If sourceRange.Rows.Count < 2 Or sourceRange.Columns.Count < 1 Then
        HostExcelExtractUniqueCopy = FailureJson("E_RANGE_TOO_SMALL", "当前区域至少需要一行表头和一行数据")
        Exit Function
    End If

    Dim defaultField, fieldName, fieldColumn, useWholeRow
    defaultField = InferDefaultFieldName(appObj, sourceRange)
    fieldName = defaultField
    useWholeRow = False
    fieldColumn = 0
    If Len(Trim(CStr(fieldName))) = 0 Then
        useWholeRow = True
        fieldName = "[整行]"
    Else
        fieldColumn = FindHeaderColumn(sourceRange, fieldName)
        If fieldColumn <= 0 Then
            HostExcelExtractUniqueCopy = FailureJson("E_FIELD_NOT_FOUND", "未找到默认键字段：" & fieldName)
            Exit Function
        End If
    End If

    Dim totalRows, uniqueRows, duplicateRows, duplicateGroups
    Dim keys(), counts(), keyCount
    Dim rowIndex, key, i, sampleText, sampleCount
    totalRows = sourceRange.Rows.Count - 1
    keyCount = 0
    For rowIndex = 2 To sourceRange.Rows.Count
        key = BuildRowKey(sourceRange, rowIndex, fieldColumn, useWholeRow)
        AddKeyCount keys, counts, keyCount, key
    Next

    uniqueRows = 0
    duplicateRows = 0
    duplicateGroups = 0
    sampleText = ""
    sampleCount = 0
    For i = 0 To keyCount - 1
        If counts(i) = 1 Then
            uniqueRows = uniqueRows + 1
        Else
            duplicateGroups = duplicateGroups + 1
            duplicateRows = duplicateRows + counts(i)
        End If
        If sampleCount < 8 Then
            If sampleCount > 0 Then sampleText = sampleText & "; "
            sampleText = sampleText & keys(i) & " x" & CStr(counts(i))
            sampleCount = sampleCount + 1
        End If
    Next

    Dim headersJson, rangeSummary, previewText, planId, planPreview
    headersJson = SafeHostText("GetHeaders")
    rangeSummary = SafeHostText("GetRangeSummary")
    previewText = "Excel 唯一值提取副本预检（尚未写入）" & vbCrLf & _
        "源表=" & sourceRange.Worksheet.Name & "，区域=" & sourceRange.Address & vbCrLf & _
        "键字段=" & fieldName & "（智能默认；当前列优先，否则首列；整行键可用 *）" & vbCrLf & _
        "数据行=" & CStr(totalRows) & "，唯一键=" & CStr(keyCount) & "，仅出现一次=" & CStr(uniqueRows) & _
        "，重复组=" & CStr(duplicateGroups) & "，重复行=" & CStr(duplicateRows) & vbCrLf & _
        "规则=保留每个键的首现行；输出=新工作表；源数据不变" & vbCrLf & _
        "样例=" & sampleText & vbCrLf & _
        "Host.GetHeaders: " & headersJson & vbCrLf & _
        "Host.GetRangeSummary: " & rangeSummary & vbCrLf & vbCrLf & _
        "确认生成请输入“生成”。" & vbCrLf & _
        "改字段：生成|字段名；整行键：生成|*；取消请清空或输入其他内容"

    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_unique_extract_copy", "{""field"":""" & EscapeJson(fieldName) & """,""uniqueKeys"":" & CStr(keyCount) & "}", "office.excel.uniqueExtract"
    planPreview = SafePreviewWritePlan()
    SafeRollbackWritePlan planId

    Dim confirmRaw, overrideField
    confirmRaw = SafePrompt(previewText, "")
    If Not ParseGenerateConfirm(confirmRaw, overrideField) Then
        HostExcelExtractUniqueCopy = FailureJson("E_CONFIRM_REQUIRED", "未确认生成，已取消且未修改工作簿")
        Exit Function
    End If

    If Len(overrideField) > 0 Then
        If StrComp(overrideField, "*", vbTextCompare) = 0 Or StrComp(overrideField, "整行", vbTextCompare) = 0 Then
            useWholeRow = True
            fieldName = "[整行]"
            fieldColumn = 0
        Else
            fieldColumn = FindHeaderColumn(sourceRange, overrideField)
            If fieldColumn <= 0 Then
                HostExcelExtractUniqueCopy = FailureJson("E_FIELD_NOT_FOUND", "未找到键字段：" & overrideField)
                Exit Function
            End If
            useWholeRow = False
            fieldName = Trim(CStr(overrideField))
        End If

        keyCount = 0
        Erase keys
        Erase counts
        For rowIndex = 2 To sourceRange.Rows.Count
            key = BuildRowKey(sourceRange, rowIndex, fieldColumn, useWholeRow)
            AddKeyCount keys, counts, keyCount, key
        Next
    End If

    Dim workbook, outputSheet, outputName
    Set workbook = sourceRange.Worksheet.Parent
    outputName = UniqueSheetName(workbook, "唯一值提取")
    Set outputSheet = workbook.Worksheets.Add
    outputSheet.Name = outputName
    If Err.Number <> 0 Then
        Err.Clear
        SafeDeleteSheet appObj, outputSheet
        HostExcelExtractUniqueCopy = FailureJson("E_OUTPUT_SHEET", "无法创建唯一值提取结果表")
        Exit Function
    End If

    Dim colIndex, outputRow, written, seenKeys(), seenCount
    For colIndex = 1 To sourceRange.Columns.Count
        outputSheet.Cells(1, colIndex).Value = sourceRange.Cells(1, colIndex).Value
        outputSheet.Cells(1, colIndex).NumberFormat = sourceRange.Cells(1, colIndex).NumberFormat
    Next
    outputSheet.Cells(1, sourceRange.Columns.Count + 1).Value = "_来源行"
    outputSheet.Cells(1, sourceRange.Columns.Count + 2).Value = "_键值"
    outputSheet.Cells(1, sourceRange.Columns.Count + 3).Value = "_重复次数"

    outputRow = 1
    written = 0
    seenCount = 0
    For rowIndex = 2 To sourceRange.Rows.Count
        key = BuildRowKey(sourceRange, rowIndex, fieldColumn, useWholeRow)
        If Not KeySeen(seenKeys, seenCount, key) Then
            ReDim Preserve seenKeys(seenCount)
            seenKeys(seenCount) = key
            seenCount = seenCount + 1

            outputRow = outputRow + 1
            For colIndex = 1 To sourceRange.Columns.Count
                outputSheet.Cells(outputRow, colIndex).Value = sourceRange.Cells(rowIndex, colIndex).Value
                outputSheet.Cells(outputRow, colIndex).NumberFormat = sourceRange.Cells(rowIndex, colIndex).NumberFormat
            Next
            outputSheet.Cells(outputRow, sourceRange.Columns.Count + 1).Value = sourceRange.Row + rowIndex - 1
            outputSheet.Cells(outputRow, sourceRange.Columns.Count + 2).Value = DisplayKey(sourceRange, rowIndex, fieldColumn, useWholeRow)
            outputSheet.Cells(outputRow, sourceRange.Columns.Count + 3).Value = GetKeyCount(keys, counts, keyCount, key)
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
        HostExcelExtractUniqueCopy = FailureJson("E_COPY_FAILED", "唯一值提取写入失败，已删除未完成结果表：" & writeError)
        Exit Function
    End If

    Dim summary
    summary = "Excel 唯一值提取副本完成：结果表=" & outputName & "，输出行=" & CStr(written) & "，键字段=" & fieldName & "，源数据未修改"
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostExcelExtractUniqueCopy = "{""ok"":true,""sourceUnchanged"":true,""outputSheet"":""" & EscapeJson(outputName) & _
        """,""outputRows"":" & CStr(written) & ",""uniqueKeys"":" & CStr(keyCount) & ",""keyField"":""" & EscapeJson(fieldName) & _
        """,""writePlanPreview"":""" & EscapeJson(planPreview) & """,""message"":""" & EscapeJson(summary) & """}"
End Function

Function InferDefaultFieldName(appObj, sourceRange)
    On Error Resume Next
    Dim activeCell, relCol, headerText
    InferDefaultFieldName = Trim(CStr(sourceRange.Cells(1, 1).Text))
    Set activeCell = appObj.ActiveCell
    If Err.Number <> 0 Or TypeName(activeCell) = "Empty" Or TypeName(activeCell) = "Nothing" Then
        Err.Clear
        Exit Function
    End If
    If activeCell.Row >= sourceRange.Row And activeCell.Row <= sourceRange.Row + sourceRange.Rows.Count - 1 Then
        If activeCell.Column >= sourceRange.Column And activeCell.Column <= sourceRange.Column + sourceRange.Columns.Count - 1 Then
            relCol = activeCell.Column - sourceRange.Column + 1
            headerText = Trim(CStr(sourceRange.Cells(1, relCol).Text))
            If Len(headerText) > 0 Then InferDefaultFieldName = headerText
        End If
    End If
    Err.Clear
End Function

Function ParseGenerateConfirm(rawText, ByRef overrideField)
    Dim text, parts, head, tail
    overrideField = ""
    text = Trim(CStr(rawText))
    If Len(text) = 0 Then
        ParseGenerateConfirm = False
        Exit Function
    End If
    If InStr(1, text, "|", vbBinaryCompare) > 0 Then
        parts = Split(text, "|")
        head = Trim(CStr(parts(0)))
        If UBound(parts) >= 1 Then tail = Trim(CStr(parts(1))) Else tail = ""
        If StrComp(head, "生成", vbTextCompare) = 0 Then
            overrideField = tail
            ParseGenerateConfirm = True
        Else
            ParseGenerateConfirm = False
        End If
    Else
        ParseGenerateConfirm = (StrComp(text, "生成", vbTextCompare) = 0)
    End If
End Function

Function BuildRowKey(sourceRange, rowIndex, fieldColumn, useWholeRow)
    Dim colIndex, parts, cellText
    If Not useWholeRow And fieldColumn > 0 Then
        BuildRowKey = NormalizeKey(sourceRange.Cells(rowIndex, fieldColumn).Text)
        Exit Function
    End If
    parts = ""
    For colIndex = 1 To sourceRange.Columns.Count
        cellText = NormalizeKey(sourceRange.Cells(rowIndex, colIndex).Text)
        If colIndex > 1 Then parts = parts & "|"
        parts = parts & cellText
    Next
    BuildRowKey = parts
End Function

Function DisplayKey(sourceRange, rowIndex, fieldColumn, useWholeRow)
    Dim colIndex, parts, cellText
    If Not useWholeRow And fieldColumn > 0 Then
        DisplayKey = NormalizeCellText(sourceRange.Cells(rowIndex, fieldColumn).Text)
        Exit Function
    End If
    parts = ""
    For colIndex = 1 To sourceRange.Columns.Count
        cellText = NormalizeCellText(sourceRange.Cells(rowIndex, colIndex).Text)
        If colIndex > 1 Then parts = parts & "|"
        parts = parts & cellText
    Next
    DisplayKey = parts
End Function

Sub AddKeyCount(ByRef keys, ByRef counts, ByRef keyCount, keyText)
    Dim i
    For i = 0 To keyCount - 1
        If StrComp(keys(i), keyText, vbBinaryCompare) = 0 Then
            counts(i) = counts(i) + 1
            Exit Sub
        End If
    Next
    ReDim Preserve keys(keyCount)
    ReDim Preserve counts(keyCount)
    keys(keyCount) = keyText
    counts(keyCount) = 1
    keyCount = keyCount + 1
End Sub

Function GetKeyCount(ByRef keys, ByRef counts, keyCount, keyText)
    Dim i
    For i = 0 To keyCount - 1
        If StrComp(keys(i), keyText, vbBinaryCompare) = 0 Then
            GetKeyCount = counts(i)
            Exit Function
        End If
    Next
    GetKeyCount = 0
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

Function SafeHostText(methodName)
    On Error Resume Next
    If methodName = "GetHeaders" Then
        SafeHostText = Host.GetHeaders()
    ElseIf methodName = "GetRangeSummary" Then
        SafeHostText = Host.GetRangeSummary()
    Else
        SafeHostText = ""
    End If
    If Err.Number <> 0 Then SafeHostText = ""
    Err.Clear
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
