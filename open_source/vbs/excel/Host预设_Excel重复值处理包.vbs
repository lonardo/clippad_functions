' 函数名: HostExcelDuplicateValueToolkit

' 描述: 按主键字段预览重复组与重复行，确认后可选高亮、提取重复副本、标注列或生成去重副本；默认不删除源行

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

    Main = HostExcelDuplicateValueToolkit(appObj)

End Function

Function HostExcelDuplicateValueToolkit(appObj)

    On Error Resume Next

    Dim sourceRange

    Set sourceRange = appObj.Selection.CurrentRegion

    If Err.Number <> 0 Or TypeName(sourceRange) = "Empty" Or TypeName(sourceRange) = "Nothing" Then

        Err.Clear

        HostExcelDuplicateValueToolkit = FailureJson("E_NO_RANGE", "请先选中带表头的数据区域")

        Exit Function

    End If

    If sourceRange.Rows.Count < 2 Or sourceRange.Columns.Count < 1 Then

        HostExcelDuplicateValueToolkit = FailureJson("E_RANGE_TOO_SMALL", "当前区域至少需要一行表头和一行数据")

        Exit Function

    End If

    Dim fieldName, actionText, fieldColumn, actionCode

    fieldName = Trim(SafePrompt("主键字段名称（必须与表头一致；留空=整行比较）", CStr(sourceRange.Cells(1, 1).Text)))

    actionText = Trim(SafePrompt("处理动作：高亮 / 提取重复 / 标注 / 去重副本", "提取重复"))

    actionCode = NormalizeDuplicateAction(actionText)

    If Len(actionCode) = 0 Then

        HostExcelDuplicateValueToolkit = FailureJson("E_BAD_ACTION", "无法识别动作，请输入：高亮 / 提取重复 / 标注 / 去重副本")

        Exit Function

    End If

    fieldColumn = 0

    If Len(fieldName) > 0 Then

        fieldColumn = FindHeaderColumn(sourceRange, fieldName)

        If fieldColumn <= 0 Then

            HostExcelDuplicateValueToolkit = FailureJson("E_FIELD_NOT_FOUND", "未找到主键字段：" & fieldName)

            Exit Function

        End If

    Else

        fieldName = "[整行]"

    End If

    Dim rowIndex, key, keys(), counts(), firstRows(), keyCount, i

    Dim totalRows, duplicateRows, duplicateGroups, uniqueRows, sampleText, sampleCount

    keyCount = 0

    totalRows = sourceRange.Rows.Count - 1

    For rowIndex = 2 To sourceRange.Rows.Count

        key = BuildRowKey(sourceRange, rowIndex, fieldColumn)

        If Len(key) > 0 Then AddKeyCount keys, counts, firstRows, keyCount, key, sourceRange.Row + rowIndex - 1

    Next

    duplicateRows = 0

    duplicateGroups = 0

    uniqueRows = 0

    sampleText = ""

    sampleCount = 0

    For i = 0 To keyCount - 1

        If counts(i) > 1 Then

            duplicateGroups = duplicateGroups + 1

            duplicateRows = duplicateRows + counts(i)

            If sampleCount < 8 Then

                sampleText = sampleText & keys(i) & " x" & CStr(counts(i)) & "; "

                sampleCount = sampleCount + 1

            End If

        Else

            uniqueRows = uniqueRows + 1

        End If

    Next

    Dim headersJson, rangeSummary, previewText, planId, planPreview, confirmWord, confirmPrompt

    headersJson = SafeHostText("GetHeaders")

    rangeSummary = SafeHostText("GetRangeSummary")

    previewText = "Excel 重复值处理预检（尚未写入）" & vbCrLf & _
        "区域=" & sourceRange.Address & "；主键=" & fieldName & "；动作=" & ActionLabel(actionCode) & vbCrLf & _
        "数据行=" & CStr(totalRows) & "；唯一键=" & CStr(uniqueRows) & "；重复组=" & CStr(duplicateGroups) & "；重复行=" & CStr(duplicateRows) & vbCrLf & _
        "样例=" & sampleText & vbCrLf & _
        "默认不删除源行；高亮/标注会改源表格式或加列，提取/去重输出到新表" & vbCrLf & _
        "Host.GetHeaders: " & headersJson & vbCrLf & _
        "Host.GetRangeSummary: " & rangeSummary

    confirmWord = ConfirmWordForAction(actionCode)

    confirmPrompt = "确认执行重复值处理，请输入：" & confirmWord

    planId = SafeBeginWritePlan()

    SafeRecordWrite "excel_duplicate_toolkit", "{""field"":""" & EscapeJson(fieldName) & """,""action"":""" & EscapeJson(actionCode) & """,""duplicateRows"":" & CStr(duplicateRows) & "}", "office.excel.quality"

    planPreview = SafePreviewWritePlan()

    SafeRollbackWritePlan planId

    If StrComp(Trim(SafePrompt(previewText & vbCrLf & vbCrLf & confirmPrompt, "")), confirmWord, vbTextCompare) <> 0 Then

        HostExcelDuplicateValueToolkit = FailureJson("E_CONFIRM_REQUIRED", "未输入“" & confirmWord & "”，已取消且未修改工作簿")

        Exit Function

    End If

    Dim resultJson

    If actionCode = "highlight" Then

        resultJson = ApplyHighlightDuplicates(sourceRange, fieldColumn, keys, counts, keyCount, duplicateRows, planPreview)

    ElseIf actionCode = "extract" Then

        resultJson = ApplyExtractOrDedup(appObj, sourceRange, fieldColumn, keys, counts, keyCount, True, duplicateRows, planPreview)

    ElseIf actionCode = "mark" Then

        resultJson = ApplyMarkDuplicates(sourceRange, fieldColumn, keys, counts, keyCount, duplicateRows, planPreview)

    Else

        resultJson = ApplyExtractOrDedup(appObj, sourceRange, fieldColumn, keys, counts, keyCount, False, duplicateRows, planPreview)

    End If

    HostExcelDuplicateValueToolkit = resultJson

End Function

Function ApplyHighlightDuplicates(sourceRange, fieldColumn, ByRef keys, ByRef counts, keyCount, duplicateRows, planPreview)

    On Error Resume Next

    Dim rowIndex, key, highlighted

    highlighted = 0

    For rowIndex = 2 To sourceRange.Rows.Count

        key = BuildRowKey(sourceRange, rowIndex, fieldColumn)

        If Len(key) > 0 And GetKeyCount(keys, counts, keyCount, key) > 1 Then

            If fieldColumn > 0 Then

                sourceRange.Cells(rowIndex, fieldColumn).Interior.Color = RGB(255, 199, 206)

                sourceRange.Cells(rowIndex, fieldColumn).Font.Color = RGB(156, 0, 6)

            Else

                sourceRange.Rows(rowIndex).Interior.Color = RGB(255, 199, 206)

            End If

            highlighted = highlighted + 1

        End If

    Next

    Dim summary

    summary = "Excel 重复值高亮完成；高亮行=" & CStr(highlighted) & "；重复行=" & CStr(duplicateRows) & "；未删除源行"

    Host.WriteClipboard summary

    SafeWriteLog summary

    ApplyHighlightDuplicates = "{""ok"":true,""action"":""highlight"",""highlightedRows"":" & CStr(highlighted) & _
        ",""duplicateRows"":" & CStr(duplicateRows) & ",""sourceRowsDeleted"":false,""writePlanPreview"":""" & EscapeJson(planPreview) & _
        """,""message"":""" & EscapeJson(summary) & """}"

End Function

Function ApplyMarkDuplicates(sourceRange, fieldColumn, ByRef keys, ByRef counts, keyCount, duplicateRows, planPreview)

    On Error Resume Next

    Dim markCol, rowIndex, key, marked, ws

    Set ws = sourceRange.Worksheet

    markCol = sourceRange.Column + sourceRange.Columns.Count

    ws.Cells(sourceRange.Row, markCol).Value = "_重复标记"

    ws.Cells(sourceRange.Row, markCol).Font.Bold = True

    marked = 0

    For rowIndex = 2 To sourceRange.Rows.Count

        key = BuildRowKey(sourceRange, rowIndex, fieldColumn)

        If Len(key) > 0 And GetKeyCount(keys, counts, keyCount, key) > 1 Then

            ws.Cells(sourceRange.Row + rowIndex - 1, markCol).Value = "重复 x" & CStr(GetKeyCount(keys, counts, keyCount, key))

            marked = marked + 1

        Else

            ws.Cells(sourceRange.Row + rowIndex - 1, markCol).Value = "唯一"

        End If

    Next

    Dim summary

    summary = "Excel 重复值标注完成；标注列=_重复标记；重复行=" & CStr(marked) & "；未删除源行"

    Host.WriteClipboard summary

    SafeWriteLog summary

    ApplyMarkDuplicates = "{""ok"":true,""action"":""mark"",""markedRows"":" & CStr(marked) & _
        ",""duplicateRows"":" & CStr(duplicateRows) & ",""sourceRowsDeleted"":false,""writePlanPreview"":""" & EscapeJson(planPreview) & _
        """,""message"":""" & EscapeJson(summary) & """}"

End Function

Function ApplyExtractOrDedup(appObj, sourceRange, fieldColumn, ByRef keys, ByRef counts, keyCount, extractDupOnly, duplicateRows, planPreview)

    On Error Resume Next

    Dim workbook, outputSheet, outputName, baseName, rowIndex, colIndex, outputRow, key, written, keep

    Set workbook = sourceRange.Worksheet.Parent

    If extractDupOnly Then baseName = "重复提取" Else baseName = "去重副本"

    outputName = UniqueSheetName(workbook, baseName)

    Set outputSheet = workbook.Worksheets.Add

    outputSheet.Name = outputName

    If Err.Number <> 0 Then

        Err.Clear

        SafeDeleteSheet appObj, outputSheet

        ApplyExtractOrDedup = FailureJson("E_OUTPUT_SHEET", "无法创建结果工作表")

        Exit Function

    End If

    For colIndex = 1 To sourceRange.Columns.Count

        outputSheet.Cells(1, colIndex).Value = sourceRange.Cells(1, colIndex).Value

    Next

    outputSheet.Cells(1, sourceRange.Columns.Count + 1).Value = "_来源行"

    outputSheet.Cells(1, sourceRange.Columns.Count + 2).Value = "_重复次数"

    outputRow = 1

    written = 0

    Dim seenKeys(), seenCount, seen

    seenCount = 0

    For rowIndex = 2 To sourceRange.Rows.Count

        key = BuildRowKey(sourceRange, rowIndex, fieldColumn)

        keep = False

        If extractDupOnly Then

            keep = (Len(key) > 0 And GetKeyCount(keys, counts, keyCount, key) > 1)

        Else

            If Len(key) = 0 Then

                keep = True

            ElseIf Not KeySeen(seenKeys, seenCount, key) Then

                keep = True

                ReDim Preserve seenKeys(seenCount)

                seenKeys(seenCount) = key

                seenCount = seenCount + 1

            End If

        End If

        If keep Then

            outputRow = outputRow + 1

            For colIndex = 1 To sourceRange.Columns.Count

                outputSheet.Cells(outputRow, colIndex).Value = sourceRange.Cells(rowIndex, colIndex).Value

                outputSheet.Cells(outputRow, colIndex).NumberFormat = sourceRange.Cells(rowIndex, colIndex).NumberFormat

            Next

            outputSheet.Cells(outputRow, sourceRange.Columns.Count + 1).Value = sourceRange.Row + rowIndex - 1

            If Len(key) > 0 Then

                outputSheet.Cells(outputRow, sourceRange.Columns.Count + 2).Value = GetKeyCount(keys, counts, keyCount, key)

            Else

                outputSheet.Cells(outputRow, sourceRange.Columns.Count + 2).Value = 1

            End If

            written = written + 1

        End If

    Next

    outputSheet.Rows(1).Font.Bold = True

    outputSheet.Columns.AutoFit

    Dim summary, actionName

    If extractDupOnly Then actionName = "extract" Else actionName = "dedup"

    summary = "Excel 重复值" & ActionLabel(actionName) & "完成；输出表=" & outputName & "；输出行=" & CStr(written) & "；源数据未删除"

    Host.WriteClipboard summary

    SafeWriteLog summary

    ApplyExtractOrDedup = "{""ok"":true,""action"":""" & EscapeJson(actionName) & """,""sourceUnchanged"":true,""outputSheet"":""" & EscapeJson(outputName) & _
        """,""outputRows"":" & CStr(written) & ",""duplicateRows"":" & CStr(duplicateRows) & ",""sourceRowsDeleted"":false,""writePlanPreview"":""" & EscapeJson(planPreview) & _
        """,""message"":""" & EscapeJson(summary) & """}"

End Function

Function BuildRowKey(sourceRange, rowIndex, fieldColumn)

    Dim colIndex, parts, cellText

    If fieldColumn > 0 Then

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

Sub AddKeyCount(ByRef keys, ByRef counts, ByRef firstRows, ByRef keyCount, keyText, rowNumber)

    Dim i

    For i = 0 To keyCount - 1

        If StrComp(keys(i), keyText, vbBinaryCompare) = 0 Then

            counts(i) = counts(i) + 1

            Exit Sub

        End If

    Next

    ReDim Preserve keys(keyCount)

    ReDim Preserve counts(keyCount)

    ReDim Preserve firstRows(keyCount)

    keys(keyCount) = keyText

    counts(keyCount) = 1

    firstRows(keyCount) = rowNumber

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

Function NormalizeDuplicateAction(value)

    Dim text

    text = LCase(Trim(CStr(value)))

    If text = "高亮" Or text = "highlight" Then

        NormalizeDuplicateAction = "highlight"

    ElseIf text = "提取重复" Or text = "提取" Or text = "extract" Then

        NormalizeDuplicateAction = "extract"

    ElseIf text = "标注" Or text = "mark" Then

        NormalizeDuplicateAction = "mark"

    ElseIf text = "去重副本" Or text = "去重" Or text = "dedup" Then

        NormalizeDuplicateAction = "dedup"

    Else

        NormalizeDuplicateAction = ""

    End If

End Function

Function ActionLabel(code)

    If code = "highlight" Then

        ActionLabel = "高亮"

    ElseIf code = "extract" Then

        ActionLabel = "提取重复"

    ElseIf code = "mark" Then

        ActionLabel = "标注"

    Else

        ActionLabel = "去重副本"

    End If

End Function

Function ConfirmWordForAction(code)

    If code = "highlight" Then

        ConfirmWordForAction = "高亮"

    ElseIf code = "mark" Then

        ConfirmWordForAction = "标注"

    Else

        ConfirmWordForAction = "生成"

    End If

End Function

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
