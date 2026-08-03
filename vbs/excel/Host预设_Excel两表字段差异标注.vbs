' 函数名: HostExcelCompareFieldDifferences
' 描述: 按用户指定主键和核对字段预览两表差异，确认后生成逐字段差异明细；不修改两张源表
' 适用应用: Excel
' 搜索范围: 当前范围
' 搜索对象: 文本

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
    Main = HostExcelCompareFieldDifferences(appObj)
End Function

Function HostExcelCompareFieldDifferences(appObj)
    On Error Resume Next
    Dim sourceRange, sourceSheet, workbook
    Set sourceRange = appObj.Selection.CurrentRegion
    Set sourceSheet = sourceRange.Worksheet
    Set workbook = sourceSheet.Parent
    If Err.Number <> 0 Or sourceRange.Rows.Count < 2 Then
        Err.Clear
        HostExcelCompareFieldDifferences = FailureJson("E_NO_SOURCE_RANGE", "请先在来源表选中带表头的数据区域")
        Exit Function
    End If

    Dim targetSheetName, targetSheet, targetRange
    targetSheetName = Trim(SafePrompt("目标工作表名称", FindOtherSheetName(workbook, sourceSheet.Name)))
    Set targetSheet = FindWorksheet(workbook, targetSheetName)
    If IsNothing(targetSheet) Then
        HostExcelCompareFieldDifferences = FailureJson("E_TARGET_SHEET", "未找到目标工作表：" & targetSheetName)
        Exit Function
    End If
    If StrComp(sourceSheet.Name, targetSheet.Name, vbTextCompare) = 0 Then
        HostExcelCompareFieldDifferences = FailureJson("E_TARGET_SHEET", "请选择与来源表不同的目标工作表")
        Exit Function
    End If
    Set targetRange = targetSheet.UsedRange
    If targetRange.Rows.Count < 2 Then
        HostExcelCompareFieldDifferences = FailureJson("E_TARGET_RANGE", "目标工作表没有可核对的数据")
        Exit Function
    End If

    Dim sourceKeyName, targetKeyName, sourceKeyCol, targetKeyCol
    sourceKeyName = Trim(SafePrompt("来源表主键字段", CStr(sourceRange.Cells(1, 1).Text)))
    targetKeyName = Trim(SafePrompt("目标表主键字段", sourceKeyName))
    sourceKeyCol = FindHeaderColumn(sourceRange, sourceKeyName)
    targetKeyCol = FindHeaderColumn(targetRange, targetKeyName)
    If sourceKeyCol <= 0 Or targetKeyCol <= 0 Then
        HostExcelCompareFieldDifferences = FailureJson("E_KEY_FIELD", "来源表或目标表未找到指定主键字段")
        Exit Function
    End If

    Dim fieldText, sourceFields, targetFields, fieldCount, fieldError
    fieldText = SafePrompt("要核对的字段（逗号分隔；留空=两表同名非主键字段）", "")
    fieldError = BuildCompareFields(sourceRange, targetRange, sourceKeyCol, targetKeyCol, fieldText, sourceFields, targetFields, fieldCount)
    If Len(fieldError) > 0 Then
        HostExcelCompareFieldDifferences = FailureJson("E_COMPARE_FIELDS", fieldError)
        Exit Function
    End If

    Dim sourceRow, targetRow, keyText, sourceOnlyCount, targetOnlyCount, duplicateCount, differenceRows, differenceCells
    sourceOnlyCount = 0
    targetOnlyCount = 0
    duplicateCount = 0
    differenceRows = 0
    differenceCells = 0
    For sourceRow = 2 To sourceRange.Rows.Count
        keyText = NormalizeKey(sourceRange.Cells(sourceRow, sourceKeyCol).Text)
        If Len(keyText) = 0 Then
            duplicateCount = duplicateCount + 1
        ElseIf CountKeyOccurrences(sourceRange, sourceKeyCol, keyText) > 1 Or CountKeyOccurrences(targetRange, targetKeyCol, keyText) > 1 Then
            duplicateCount = duplicateCount + 1
        Else
            targetRow = FindKeyRow(targetRange, targetKeyCol, keyText)
            If targetRow <= 0 Then
                sourceOnlyCount = sourceOnlyCount + 1
            Else
                Dim rowDifferences
                rowDifferences = CountFieldDifferences(sourceRange, targetRange, sourceRow, targetRow, sourceFields, targetFields, fieldCount)
                If rowDifferences > 0 Then
                    differenceRows = differenceRows + 1
                    differenceCells = differenceCells + rowDifferences
                End If
            End If
        End If
    Next
    For targetRow = 2 To targetRange.Rows.Count
        keyText = NormalizeKey(targetRange.Cells(targetRow, targetKeyCol).Text)
        If Len(keyText) > 0 And CountKeyOccurrences(targetRange, targetKeyCol, keyText) = 1 And FindKeyRow(sourceRange, sourceKeyCol, keyText) <= 0 Then targetOnlyCount = targetOnlyCount + 1
    Next

    Dim previewText, planId, planPreview
    previewText = "Excel 两表字段差异预览（尚未写入）" & vbCrLf & _
        "来源=" & sourceSheet.Name & sourceRange.Address & "；目标=" & targetSheet.Name & targetRange.Address & vbCrLf & _
        "主键=" & sourceKeyName & "→" & targetKeyName & "；核对字段=" & BuildFieldLabels(sourceRange, sourceFields, fieldCount) & vbCrLf & _
        "字段差异行=" & CStr(differenceRows) & "；差异单元格=" & CStr(differenceCells) & "；来源缺失=" & CStr(targetOnlyCount) & _
        "；目标缺失=" & CStr(sourceOnlyCount) & "；空/重复主键风险=" & CStr(duplicateCount) & vbCrLf & _
        "确认后只创建差异明细新工作表；两张源表不变"
    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_field_difference_report", BuildPlanParams(sourceSheet.Name, targetSheet.Name, sourceKeyName, targetKeyName, fieldCount, differenceCells), "office.excel.fieldDifferenceReport"
    planPreview = SafePreviewWritePlan()
    SafeRollbackWritePlan planId
    If StrComp(Trim(SafePrompt(previewText & vbCrLf & vbCrLf & "确认生成差异明细，请输入：生成", "")), "生成", vbTextCompare) <> 0 Then
        HostExcelCompareFieldDifferences = FailureJson("E_CONFIRM_REQUIRED", "未输入“生成”，已取消且未修改工作簿")
        Exit Function
    End If

    Dim outputSheet, outputName, outputRow
    outputName = UniqueSheetName(workbook, "字段差异")
    Set outputSheet = workbook.Worksheets.Add
    outputSheet.Name = outputName
    outputSheet.Cells(1, 1).Value = "主键"
    outputSheet.Cells(1, 2).Value = "状态"
    outputSheet.Cells(1, 3).Value = "字段"
    outputSheet.Cells(1, 4).Value = "来源值"
    outputSheet.Cells(1, 5).Value = "目标值"
    outputSheet.Cells(1, 6).Value = "来源行"
    outputSheet.Cells(1, 7).Value = "目标行"
    outputRow = 1
    For sourceRow = 2 To sourceRange.Rows.Count
        keyText = NormalizeKey(sourceRange.Cells(sourceRow, sourceKeyCol).Text)
        targetRow = FindKeyRow(targetRange, targetKeyCol, keyText)
        If Len(keyText) = 0 Then
            outputRow = outputRow + 1
            WriteDifferenceRow outputSheet, outputRow, "", "空主键，未比较", "", "", "", sourceRange.Row + sourceRow - 1, 0
        ElseIf CountKeyOccurrences(sourceRange, sourceKeyCol, keyText) > 1 Or CountKeyOccurrences(targetRange, targetKeyCol, keyText) > 1 Then
            outputRow = outputRow + 1
            WriteDifferenceRow outputSheet, outputRow, keyText, "重复主键，未比较", "", "", "", sourceRange.Row + sourceRow - 1, TargetAbsoluteRow(targetRange, targetRow)
        ElseIf targetRow <= 0 Then
            outputRow = outputRow + 1
            WriteDifferenceRow outputSheet, outputRow, keyText, "目标缺失", "", "", "", sourceRange.Row + sourceRow - 1, 0
        Else
            WriteFieldDifferences outputSheet, outputRow, keyText, sourceRange, targetRange, sourceRow, targetRow, sourceFields, targetFields, fieldCount
        End If
    Next
    For targetRow = 2 To targetRange.Rows.Count
        keyText = NormalizeKey(targetRange.Cells(targetRow, targetKeyCol).Text)
        If Len(keyText) > 0 And CountKeyOccurrences(targetRange, targetKeyCol, keyText) = 1 And FindKeyRow(sourceRange, sourceKeyCol, keyText) <= 0 Then
            outputRow = outputRow + 1
            WriteDifferenceRow outputSheet, outputRow, keyText, "来源缺失", "", "", "", 0, targetRange.Row + targetRow - 1
        End If
    Next
    outputSheet.Rows(1).Font.Bold = True
    outputSheet.Columns.AutoFit
    If Err.Number <> 0 Then
        Dim writeError
        writeError = Err.Description
        Err.Clear
        SafeDeleteSheet appObj, outputSheet
        HostExcelCompareFieldDifferences = FailureJson("E_DIFFERENCE_WRITE", "差异明细写入失败，已删除未完成工作表：" & writeError)
        Exit Function
    End If

    Dim summary
    summary = "Excel 两表字段差异明细已生成；工作表=" & outputName & "；差异单元格=" & CStr(differenceCells) & "；两张源表未修改"
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostExcelCompareFieldDifferences = "{""ok"":true,""sourceUnchanged"":true,""outputSheet"":""" & EscapeJson(outputName) & _
        """,""differenceRows"":" & CStr(differenceRows) & ",""differenceCells"":" & CStr(differenceCells) & _
        ",""writePlanPreview"":""" & EscapeJson(planPreview) & """,""message"":""" & EscapeJson(summary) & """}"
End Function

Function BuildCompareFields(sourceRange, targetRange, sourceKeyCol, targetKeyCol, fieldText, ByRef sourceFields, ByRef targetFields, ByRef fieldCount)
    Dim parts, part, sourceCol, targetCol, index
    BuildCompareFields = ""
    fieldCount = 0
    ReDim sourceFields(sourceRange.Columns.Count)
    ReDim targetFields(sourceRange.Columns.Count)
    fieldText = Replace(Trim(CStr(fieldText)), "，", ",")
    If Len(fieldText) = 0 Then
        For index = 1 To sourceRange.Columns.Count
            If index <> sourceKeyCol Then
                targetCol = FindHeaderColumn(targetRange, CStr(sourceRange.Cells(1, index).Text))
                If targetCol > 0 And targetCol <> targetKeyCol Then
                    fieldCount = fieldCount + 1
                    sourceFields(fieldCount) = index
                    targetFields(fieldCount) = targetCol
                End If
            End If
        Next
    Else
        parts = Split(fieldText, ",")
        For Each part In parts
            sourceCol = FindHeaderColumn(sourceRange, Trim(CStr(part)))
            targetCol = FindHeaderColumn(targetRange, Trim(CStr(part)))
            If sourceCol <= 0 Or targetCol <= 0 Then
                BuildCompareFields = "两表中未同时找到核对字段：" & CStr(part)
                Exit Function
            End If
            If sourceCol = sourceKeyCol Or targetCol = targetKeyCol Then
                BuildCompareFields = "主键字段不能同时作为差异核对字段：" & CStr(part)
                Exit Function
            End If
            fieldCount = fieldCount + 1
            sourceFields(fieldCount) = sourceCol
            targetFields(fieldCount) = targetCol
        Next
    End If
    If fieldCount = 0 Then BuildCompareFields = "没有可核对的同名非主键字段"
End Function

Function CountFieldDifferences(sourceRange, targetRange, sourceRow, targetRow, sourceFields, targetFields, fieldCount)
    Dim index, count
    count = 0
    For index = 1 To fieldCount
        If Not ValuesEqual(sourceRange.Cells(sourceRow, sourceFields(index)).Text, targetRange.Cells(targetRow, targetFields(index)).Text) Then count = count + 1
    Next
    CountFieldDifferences = count
End Function

Sub WriteFieldDifferences(outputSheet, ByRef outputRow, keyText, sourceRange, targetRange, sourceRow, targetRow, sourceFields, targetFields, fieldCount)
    Dim index, sourceValue, targetValue
    For index = 1 To fieldCount
        sourceValue = CStr(sourceRange.Cells(sourceRow, sourceFields(index)).Text)
        targetValue = CStr(targetRange.Cells(targetRow, targetFields(index)).Text)
        If Not ValuesEqual(sourceValue, targetValue) Then
            outputRow = outputRow + 1
            WriteDifferenceRow outputSheet, outputRow, keyText, "字段不一致", CStr(sourceRange.Cells(1, sourceFields(index)).Text), sourceValue, targetValue, sourceRange.Row + sourceRow - 1, targetRange.Row + targetRow - 1
        End If
    Next
End Sub

Sub WriteDifferenceRow(sheetObj, rowIndex, keyText, statusText, fieldName, sourceValue, targetValue, sourceRow, targetRow)
    sheetObj.Cells(rowIndex, 1).Value = keyText
    sheetObj.Cells(rowIndex, 2).Value = statusText
    sheetObj.Cells(rowIndex, 3).Value = fieldName
    sheetObj.Cells(rowIndex, 4).Value = sourceValue
    sheetObj.Cells(rowIndex, 5).Value = targetValue
    If sourceRow > 0 Then sheetObj.Cells(rowIndex, 6).Value = sourceRow
    If targetRow > 0 Then sheetObj.Cells(rowIndex, 7).Value = targetRow
End Sub

Function ValuesEqual(sourceValue, targetValue)
    ValuesEqual = (StrComp(Trim(CStr(sourceValue)), Trim(CStr(targetValue)), vbTextCompare) = 0)
End Function

Function BuildFieldLabels(sourceRange, sourceFields, fieldCount)
    Dim index, text
    text = ""
    For index = 1 To fieldCount
        If Len(text) > 0 Then text = text & "，"
        text = text & CStr(sourceRange.Cells(1, sourceFields(index)).Text)
    Next
    BuildFieldLabels = text
End Function

Function NormalizeKey(value)
    NormalizeKey = LCase(Trim(CStr(value)))
End Function

Function FindKeyRow(dataRange, keyColumn, keyText)
    Dim rowIndex
    FindKeyRow = 0
    If Len(keyText) = 0 Then Exit Function
    For rowIndex = 2 To dataRange.Rows.Count
        If StrComp(NormalizeKey(dataRange.Cells(rowIndex, keyColumn).Text), keyText, vbTextCompare) = 0 Then FindKeyRow = rowIndex: Exit Function
    Next
End Function

Function CountKeyOccurrences(dataRange, keyColumn, keyText)
    Dim rowIndex, count
    count = 0
    If Len(keyText) = 0 Then CountKeyOccurrences = 0: Exit Function
    For rowIndex = 2 To dataRange.Rows.Count
        If StrComp(NormalizeKey(dataRange.Cells(rowIndex, keyColumn).Text), keyText, vbTextCompare) = 0 Then count = count + 1
    Next
    CountKeyOccurrences = count
End Function

Function FindHeaderColumn(dataRange, fieldName)
    Dim colIndex
    FindHeaderColumn = 0
    For colIndex = 1 To dataRange.Columns.Count
        If StrComp(Trim(CStr(dataRange.Cells(1, colIndex).Text)), Trim(CStr(fieldName)), vbTextCompare) = 0 Then FindHeaderColumn = colIndex: Exit Function
    Next
End Function

Function FindOtherSheetName(workbook, sourceName)
    Dim sheetObj
    FindOtherSheetName = ""
    For Each sheetObj In workbook.Worksheets
        If StrComp(sheetObj.Name, sourceName, vbTextCompare) <> 0 Then FindOtherSheetName = sheetObj.Name: Exit Function
    Next
End Function

Function FindWorksheet(workbook, sheetName)
    Dim sheetObj
    Set FindWorksheet = Nothing
    For Each sheetObj In workbook.Worksheets
        If StrComp(CStr(sheetObj.Name), CStr(sheetName), vbTextCompare) = 0 Then Set FindWorksheet = sheetObj: Exit Function
    Next
End Function

Function TargetAbsoluteRow(targetRange, relativeRow)
    If relativeRow > 0 Then TargetAbsoluteRow = targetRange.Row + relativeRow - 1 Else TargetAbsoluteRow = 0
End Function

Function UniqueSheetName(workbook, baseName)
    Dim candidate, suffix
    candidate = Left(baseName, 31)
    suffix = 1
    Do While Not IsNothing(FindWorksheet(workbook, candidate))
        suffix = suffix + 1
        candidate = Left(baseName, 27) & "_" & CStr(suffix)
    Loop
    UniqueSheetName = candidate
End Function

Function IsNothing(value)
    IsNothing = (TypeName(value) = "Nothing" Or TypeName(value) = "Empty")
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

Function BuildPlanParams(sourceSheet, targetSheet, sourceKey, targetKey, fieldCount, differenceCells)
    BuildPlanParams = "{""sourceSheet"":""" & EscapeJson(sourceSheet) & """,""targetSheet"":""" & EscapeJson(targetSheet) & _
        """,""sourceKey"":""" & EscapeJson(sourceKey) & """,""targetKey"":""" & EscapeJson(targetKey) & _
        """,""fields"":" & CStr(fieldCount) & ",""differenceCells"":" & CStr(differenceCells) & "}"
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
