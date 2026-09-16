' 函数名: HostExcelAdvancedMergeRowsCopy
' 描述: 按主键字段分组，将其余字段去重后用分隔符合并到新工作表；预览分组与输出行数，确认后生成副本，源表不变
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
    Main = HostExcelAdvancedMergeRowsCopy(appObj)
End Function

Function HostExcelAdvancedMergeRowsCopy(appObj)
    On Error Resume Next

    Dim sourceRange
    Set sourceRange = appObj.Selection.CurrentRegion
    If Err.Number <> 0 Or TypeName(sourceRange) = "Empty" Or TypeName(sourceRange) = "Nothing" Then
        Err.Clear
        HostExcelAdvancedMergeRowsCopy = FailureJson("E_NO_RANGE", "请先选中带表头的数据区域")
        Exit Function
    End If
    If sourceRange.Rows.Count < 2 Or sourceRange.Columns.Count < 2 Then
        HostExcelAdvancedMergeRowsCopy = FailureJson("E_RANGE_TOO_SMALL", "当前区域至少需要一行表头、一行数据和两列")
        Exit Function
    End If

    Dim keyField, keyCol, delim, keepFirstText, keepFirst
    keyField = Trim(SafePrompt("主键字段名称（按该字段合并行）", CStr(sourceRange.Cells(1, 1).Text)))
    If Len(keyField) = 0 Then
        HostExcelAdvancedMergeRowsCopy = FailureJson("E_KEY_REQUIRED", "必须指定主键字段")
        Exit Function
    End If
    keyCol = FindHeaderColumn(sourceRange, keyField)
    If keyCol <= 0 Then
        HostExcelAdvancedMergeRowsCopy = FailureJson("E_FIELD_NOT_FOUND", "未找到主键字段：" & keyField)
        Exit Function
    End If

    delim = SafePrompt("非主键字段合并分隔符", "；")
    If Len(delim) = 0 Then delim = "；"
    keepFirstText = Trim(SafePrompt("非主键字段若多值不同：合并全部 / 仅保留首个", "合并全部"))
    keepFirst = (InStr(1, keepFirstText, "首", vbTextCompare) > 0)

    Dim rowIndex, key, keys(), firstRows(), members(), keyCount, i
    Dim totalRows, emptyKeyRows, maxGroup, sampleText, sampleCount
    keyCount = 0
    totalRows = sourceRange.Rows.Count - 1
    emptyKeyRows = 0
    maxGroup = 0
    sampleText = ""
    sampleCount = 0

    For rowIndex = 2 To sourceRange.Rows.Count
        key = NormalizeCellText(sourceRange.Cells(rowIndex, keyCol).Text)
        If Len(key) = 0 Then
            emptyKeyRows = emptyKeyRows + 1
        Else
            AddMergeKey keys, firstRows, members, keyCount, key, rowIndex
        End If
    Next

    For i = 0 To keyCount - 1
        If members(i) > maxGroup Then maxGroup = members(i)
        If sampleCount < 6 Then
            sampleText = sampleText & keys(i) & " x" & CStr(members(i)) & "; "
            sampleCount = sampleCount + 1
        End If
    Next

    Dim headersJson, rangeSummary, previewText, planId, planPreview
    headersJson = SafeHostTextExcel("GetHeaders")
    rangeSummary = SafeHostTextExcel("GetRangeSummary")
    previewText = "Excel 高级合并行预检（尚未写入）" & vbCrLf & _
        "区域=" & sourceRange.Address & "；主键=" & keyField & "；分隔符=" & delim & vbCrLf & _
        "数据行=" & CStr(totalRows) & "；有效主键组=" & CStr(keyCount) & "；空主键行跳过=" & CStr(emptyKeyRows) & vbCrLf & _
        "最大组内行数=" & CStr(maxGroup) & "；多值策略=" & KeepFirstLabel(keepFirst) & vbCrLf & _
        "输出=新工作表；源数据不变" & vbCrLf & _
        "样例组=" & sampleText & vbCrLf & _
        "Host.GetHeaders: " & headersJson & vbCrLf & _
        "Host.GetRangeSummary: " & rangeSummary

    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_advanced_merge_rows_copy", "{""keyField"":""" & EscapeJson(keyField) & """,""groups"":" & CStr(keyCount) & ",""keepFirst"":" & LCase(CStr(keepFirst)) & "}", "office.excel.mergeRows"
    planPreview = SafePreviewWritePlan()
    SafeRollbackWritePlan planId

    If StrComp(Trim(SafePrompt(previewText & vbCrLf & vbCrLf & "确认生成高级合并行副本，请输入：生成", "")), "生成", vbTextCompare) <> 0 Then
        HostExcelAdvancedMergeRowsCopy = FailureJson("E_CONFIRM_REQUIRED", "未输入“生成”，已取消且未修改工作簿")
        Exit Function
    End If

    Dim workbook, outputSheet, outputName, colIndex, outputRow, groupIndex, memberList, valueText
    Set workbook = sourceRange.Worksheet.Parent
    outputName = UniqueSheetName(workbook, "高级合并行")
    Set outputSheet = workbook.Worksheets.Add
    outputSheet.Name = outputName
    If Err.Number <> 0 Then
        Err.Clear
        SafeDeleteSheet appObj, outputSheet
        HostExcelAdvancedMergeRowsCopy = FailureJson("E_OUTPUT_SHEET", "无法创建合并结果工作表")
        Exit Function
    End If

    For colIndex = 1 To sourceRange.Columns.Count
        outputSheet.Cells(1, colIndex).Value = sourceRange.Cells(1, colIndex).Value
    Next
    outputSheet.Cells(1, sourceRange.Columns.Count + 1).Value = "_合并行数"
    outputSheet.Cells(1, sourceRange.Columns.Count + 2).Value = "_来源首行"

    outputRow = 1
    For groupIndex = 0 To keyCount - 1
        outputRow = outputRow + 1
        For colIndex = 1 To sourceRange.Columns.Count
            If colIndex = keyCol Then
                outputSheet.Cells(outputRow, colIndex).Value = keys(groupIndex)
            Else
                valueText = BuildMergedColumnValues(sourceRange, keys, firstRows, members, groupIndex, keyCol, colIndex, delim, keepFirst)
                outputSheet.Cells(outputRow, colIndex).Value = valueText
            End If
        Next
        outputSheet.Cells(outputRow, sourceRange.Columns.Count + 1).Value = members(groupIndex)
        outputSheet.Cells(outputRow, sourceRange.Columns.Count + 2).Value = sourceRange.Row + firstRows(groupIndex) - 1
    Next

    outputSheet.Rows(1).Font.Bold = True
    outputSheet.Columns.AutoFit
    If Err.Number <> 0 Then
        Dim writeError
        writeError = Err.Description
        Err.Clear
        SafeDeleteSheet appObj, outputSheet
        HostExcelAdvancedMergeRowsCopy = FailureJson("E_COPY_FAILED", "合并副本写入失败，已删除未完成工作表：" & writeError)
        Exit Function
    End If

    Dim summary
    summary = "Excel 高级合并行副本已生成；工作表=" & outputName & "；输出行=" & CStr(keyCount) & "；源数据行=" & CStr(totalRows) & "；主键=" & keyField & "；源数据未修改"
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostExcelAdvancedMergeRowsCopy = "{""ok"":true,""sourceUnchanged"":true,""outputSheet"":""" & EscapeJson(outputName) & """,""groupCount"":" & CStr(keyCount) & ",""sourceRows"":" & CStr(totalRows) & ",""keyField"":""" & EscapeJson(keyField) & """,""writePlanPreview"":""" & EscapeJson(planPreview) & """,""message"":""" & EscapeJson(summary) & """}"
End Function

Sub AddMergeKey(ByRef keys, ByRef firstRows, ByRef members, ByRef keyCount, key, rowIndex)
    Dim i
    For i = 0 To keyCount - 1
        If StrComp(keys(i), key, vbBinaryCompare) = 0 Then
            members(i) = members(i) + 1
            Exit Sub
        End If
    Next
    ReDim Preserve keys(keyCount)
    ReDim Preserve firstRows(keyCount)
    ReDim Preserve members(keyCount)
    keys(keyCount) = key
    firstRows(keyCount) = rowIndex
    members(keyCount) = 1
    keyCount = keyCount + 1
End Sub

Function BuildMergedColumnValues(sourceRange, ByRef keys, ByRef firstRows, ByRef members, groupIndex, keyCol, colIndex, delim, keepFirst)
    On Error Resume Next
    Dim rowIndex, key, values(), valueCount, cellText, j, exists
    key = keys(groupIndex)
    valueCount = 0
    For rowIndex = 2 To sourceRange.Rows.Count
        If StrComp(NormalizeCellText(sourceRange.Cells(rowIndex, keyCol).Text), key, vbBinaryCompare) = 0 Then
            cellText = NormalizeCellText(sourceRange.Cells(rowIndex, colIndex).Text)
            If Len(cellText) > 0 Then
                exists = False
                For j = 0 To valueCount - 1
                    If StrComp(values(j), cellText, vbBinaryCompare) = 0 Then
                        exists = True
                        Exit For
                    End If
                Next
                If Not exists Then
                    ReDim Preserve values(valueCount)
                    values(valueCount) = cellText
                    valueCount = valueCount + 1
                    If keepFirst Then Exit For
                End If
            End If
        End If
    Next
    If valueCount = 0 Then
        BuildMergedColumnValues = ""
    ElseIf valueCount = 1 Then
        BuildMergedColumnValues = values(0)
    Else
        Dim outText, k
        outText = values(0)
        For k = 1 To valueCount - 1
            outText = outText & delim & values(k)
        Next
        BuildMergedColumnValues = outText
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

Function SafeHostTextExcel(methodName)
    On Error Resume Next
    If methodName = "GetHeaders" Then
        SafeHostTextExcel = Host.GetHeaders()
    ElseIf methodName = "GetRangeSummary" Then
        SafeHostTextExcel = Host.GetRangeSummary()
    ElseIf methodName = "GetExcelContextInfo" Then
        SafeHostTextExcel = Host.GetExcelContextInfo()
    ElseIf methodName = "GetWorkbookName" Then
        SafeHostTextExcel = Host.GetWorkbookName()
    ElseIf methodName = "GetSheetName" Then
        SafeHostTextExcel = Host.GetSheetName()
    Else
        SafeHostTextExcel = ""
    End If
    If Err.Number <> 0 Then SafeHostTextExcel = ""
    Err.Clear
End Function

Function KeepFirstLabel(keepFirst)
    If keepFirst Then
        KeepFirstLabel = "仅保留首个"
    Else
        KeepFirstLabel = "合并全部"
    End If
End Function
