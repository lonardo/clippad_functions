' 函数名: HostExcelReconcileTwoSheets
' 描述: 按用户指定的主键核对当前区域与另一工作表，预览匹配、缺失和重复数量，确认后把明细写入新工作表；不修改两张源表
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
    Main = HostExcelReconcileTwoSheets(appObj)
End Function

Function HostExcelReconcileTwoSheets(appObj)
    On Error Resume Next
    Dim sourceRange, sourceSheet, workbook
    Set sourceRange = appObj.Selection.CurrentRegion
    Set sourceSheet = sourceRange.Worksheet
    Set workbook = sourceSheet.Parent
    If Err.Number <> 0 Or sourceRange.Rows.Count < 2 Then
        Err.Clear
        HostExcelReconcileTwoSheets = FailureJson("E_NO_SOURCE_RANGE", "请先在来源表选中带表头的数据区域")
        Exit Function
    End If

    Dim targetSheetName, targetSheet
    targetSheetName = Trim(SafePrompt("目标工作表名称", FindOtherSheetName(workbook, sourceSheet.Name)))
    Set targetSheet = FindWorksheet(workbook, targetSheetName)
    If TypeName(targetSheet) = "Nothing" Or TypeName(targetSheet) = "Empty" Then
        HostExcelReconcileTwoSheets = FailureJson("E_TARGET_SHEET", "未找到目标工作表：" & targetSheetName)
        Exit Function
    End If

    Dim targetRange
    Set targetRange = targetSheet.UsedRange
    If targetRange.Rows.Count < 2 Then
        HostExcelReconcileTwoSheets = FailureJson("E_TARGET_RANGE", "目标工作表没有可核对的数据")
        Exit Function
    End If

    Dim sourceKeyName, targetKeyName, sourceKeyCol, targetKeyCol
    sourceKeyName = Trim(SafePrompt("来源表主键字段", CStr(sourceRange.Cells(1, 1).Text)))
    targetKeyName = Trim(SafePrompt("目标表主键字段", sourceKeyName))
    sourceKeyCol = FindHeaderColumn(sourceRange, sourceKeyName)
    targetKeyCol = FindHeaderColumn(targetRange, targetKeyName)
    If sourceKeyCol <= 0 Then
        HostExcelReconcileTwoSheets = FailureJson("E_SOURCE_KEY", "来源表未找到主键字段：" & sourceKeyName)
        Exit Function
    End If
    If targetKeyCol <= 0 Then
        HostExcelReconcileTwoSheets = FailureJson("E_TARGET_KEY", "目标表未找到主键字段：" & targetKeyName)
        Exit Function
    End If

    Dim sourceRow, targetRow, keyText, matchedCount, targetMissingCount, sourceMissingCount, duplicateKeyCount
    matchedCount = 0
    targetMissingCount = 0
    sourceMissingCount = 0
    duplicateKeyCount = 0
    For sourceRow = 2 To sourceRange.Rows.Count
        keyText = NormalizeKey(sourceRange.Cells(sourceRow, sourceKeyCol).Text)
        If Len(keyText) > 0 Then
            targetRow = FindKeyRow(targetRange, targetKeyCol, keyText)
            If targetRow > 0 Then matchedCount = matchedCount + 1 Else targetMissingCount = targetMissingCount + 1
            If CountKeyOccurrences(sourceRange, sourceKeyCol, keyText) > 1 Then duplicateKeyCount = duplicateKeyCount + 1
        End If
    Next
    For targetRow = 2 To targetRange.Rows.Count
        keyText = NormalizeKey(targetRange.Cells(targetRow, targetKeyCol).Text)
        If Len(keyText) > 0 And FindKeyRow(sourceRange, sourceKeyCol, keyText) <= 0 Then sourceMissingCount = sourceMissingCount + 1
    Next

    Dim previewText, planId, planPreview
    previewText = "Excel 两表主键核对预览（尚未写入）" & vbCrLf & _
        "来源=" & sourceSheet.Name & sourceRange.Address & "；主键=" & sourceKeyName & vbCrLf & _
        "目标=" & targetSheet.Name & targetRange.Address & "；主键=" & targetKeyName & vbCrLf & _
        "匹配=" & CStr(matchedCount) & "；目标缺失=" & CStr(targetMissingCount) & "；来源缺失=" & CStr(sourceMissingCount) & _
        "；来源重复行=" & CStr(duplicateKeyCount) & "；结果写入新工作表；两张源表不变"
    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_key_reconciliation", "{""sourceSheet"":""" & EscapeJson(sourceSheet.Name) & """,""targetSheet"":""" & EscapeJson(targetSheet.Name) & """,""matched"":" & CStr(matchedCount) & ",""targetMissing"":" & CStr(targetMissingCount) & ",""sourceMissing"":" & CStr(sourceMissingCount) & "}", "office.excel.keyReconciliation"
    planPreview = SafePreviewWritePlan()
    SafeRollbackWritePlan planId

    If StrComp(Trim(SafePrompt(previewText & vbCrLf & vbCrLf & "确认生成核对明细，请输入：生成", "")), "生成", vbTextCompare) <> 0 Then
        HostExcelReconcileTwoSheets = FailureJson("E_CONFIRM_REQUIRED", "未输入“生成”，已取消且未修改工作簿")
        Exit Function
    End If

    Dim outputSheet, outputName, outputRow, sourceDup, targetDup, statusText
    outputName = UniqueSheetName(workbook, "主键核对")
    Set outputSheet = workbook.Worksheets.Add
    outputSheet.Name = outputName
    outputSheet.Cells(1, 1).Value = "主键"
    outputSheet.Cells(1, 2).Value = "状态"
    outputSheet.Cells(1, 3).Value = "来源行"
    outputSheet.Cells(1, 4).Value = "目标行"
    outputSheet.Cells(1, 5).Value = "来源重复数"
    outputSheet.Cells(1, 6).Value = "目标重复数"
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
        WriteResultRow outputSheet, outputRow, keyText, statusText, sourceRange.Row + sourceRow - 1, TargetAbsoluteRow(targetRange, targetRow), sourceDup, targetDup
    Next

    For targetRow = 2 To targetRange.Rows.Count
        keyText = NormalizeKey(targetRange.Cells(targetRow, targetKeyCol).Text)
        If Len(keyText) > 0 And FindKeyRow(sourceRange, sourceKeyCol, keyText) <= 0 Then
            outputRow = outputRow + 1
            targetDup = CountKeyOccurrences(targetRange, targetKeyCol, keyText)
            statusText = "来源缺失"
            If targetDup > 1 Then statusText = statusText & ";目标重复"
            WriteResultRow outputSheet, outputRow, keyText, statusText, 0, targetRange.Row + targetRow - 1, 0, targetDup
        End If
    Next
    outputSheet.Rows(1).Font.Bold = True
    outputSheet.Columns.AutoFit

    If Err.Number <> 0 Then
        Dim writeError
        writeError = Err.Description
        Err.Clear
        SafeDeleteSheet appObj, outputSheet
        HostExcelReconcileTwoSheets = FailureJson("E_RECONCILE_WRITE", "核对明细写入失败，已删除未完成工作表：" & writeError)
        Exit Function
    End If

    Dim summary
    summary = "Excel 两表主键核对完成；工作表=" & outputName & "；匹配=" & CStr(matchedCount) & _
        "；目标缺失=" & CStr(targetMissingCount) & "；来源缺失=" & CStr(sourceMissingCount) & "；源表未修改"
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostExcelReconcileTwoSheets = "{""ok"":true,""sourceUnchanged"":true,""outputSheet"":""" & EscapeJson(outputName) & _
        """,""matched"":" & CStr(matchedCount) & ",""targetMissing"":" & CStr(targetMissingCount) & _
        ",""sourceMissing"":" & CStr(sourceMissingCount) & ",""writePlanPreview"":""" & EscapeJson(planPreview) & _
        """,""message"":""" & EscapeJson(summary) & """}"
End Function

Sub WriteResultRow(sheetObj, rowIndex, keyText, statusText, sourceRow, targetRow, sourceDup, targetDup)
    sheetObj.Cells(rowIndex, 1).Value = keyText
    sheetObj.Cells(rowIndex, 2).Value = statusText
    If sourceRow > 0 Then sheetObj.Cells(rowIndex, 3).Value = sourceRow
    If targetRow > 0 Then sheetObj.Cells(rowIndex, 4).Value = targetRow
    sheetObj.Cells(rowIndex, 5).Value = sourceDup
    sheetObj.Cells(rowIndex, 6).Value = targetDup
End Sub

Function TargetAbsoluteRow(targetRange, relativeRow)
    If relativeRow > 0 Then TargetAbsoluteRow = targetRange.Row + relativeRow - 1 Else TargetAbsoluteRow = 0
End Function

Function NormalizeKey(value)
    NormalizeKey = LCase(Trim(CStr(value)))
End Function

Function FindKeyRow(dataRange, keyColumn, keyText)
    Dim rowIndex
    FindKeyRow = 0
    For rowIndex = 2 To dataRange.Rows.Count
        If StrComp(NormalizeKey(dataRange.Cells(rowIndex, keyColumn).Text), keyText, vbTextCompare) = 0 Then FindKeyRow = rowIndex: Exit Function
    Next
End Function

Function CountKeyOccurrences(dataRange, keyColumn, keyText)
    Dim rowIndex, count
    count = 0
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
