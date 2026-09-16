' 函数名: HostExcelCreateMappedCopy
' 描述: 按用户确认的字段映射把当前区域写入新工作表，先预览行数、输入字段和输出字段；不修改源数据
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
    Main = HostExcelCreateMappedCopy(appObj)
End Function

Function HostExcelCreateMappedCopy(appObj)
    On Error Resume Next
    Dim sourceRange
    Set sourceRange = appObj.Selection.CurrentRegion
    If Err.Number <> 0 Or TypeName(sourceRange) = "Empty" Or TypeName(sourceRange) = "Nothing" Then
        Err.Clear
        HostExcelCreateMappedCopy = FailureJson("E_NO_RANGE", "请先选中带表头的数据区域")
        Exit Function
    End If
    If sourceRange.Rows.Count < 2 Or sourceRange.Columns.Count < 1 Then
        HostExcelCreateMappedCopy = FailureJson("E_RANGE_TOO_SMALL", "当前区域至少需要一行表头和一行数据")
        Exit Function
    End If

    Dim headerError
    headerError = ValidateSourceHeaders(sourceRange)
    If Len(headerError) > 0 Then
        HostExcelCreateMappedCopy = FailureJson("E_SOURCE_HEADERS", headerError)
        Exit Function
    End If

    Dim mappingText, sourceColumns, targetNames, mappingCount, mappingError
    mappingText = SafePrompt("字段映射（格式：源字段=目标字段，多个映射用逗号分隔；留空表示保留全部字段）", "")
    mappingError = BuildMappings(sourceRange, mappingText, sourceColumns, targetNames, mappingCount)
    If Len(mappingError) > 0 Then
        HostExcelCreateMappedCopy = FailureJson("E_FIELD_MAPPING", mappingError)
        Exit Function
    End If

    Dim previewText, planId, planPreview
    previewText = "Excel 字段映射交付预览（尚未写入）" & vbCrLf & _
        "源工作表：" & sourceRange.Worksheet.Name & "；区域=" & sourceRange.Address & vbCrLf & _
        "数据行=" & CStr(sourceRange.Rows.Count - 1) & "；输入字段=" & CStr(sourceRange.Columns.Count) & _
        "；输出字段=" & CStr(mappingCount) & vbCrLf & _
        "映射：" & BuildMappingSummary(sourceRange, sourceColumns, targetNames, mappingCount) & vbCrLf & _
        "输出=新工作表；源数据不变；不会删除或改写原始字段"
    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_field_mapping_copy", BuildPlanParams(sourceRange, sourceColumns, targetNames, mappingCount), "office.excel.fieldMappingCopy"
    planPreview = SafePreviewWritePlan()
    SafeRollbackWritePlan planId

    If StrComp(Trim(SafePrompt(previewText & vbCrLf & vbCrLf & "确认生成字段映射副本，请输入：生成", "")), "生成", vbTextCompare) <> 0 Then
        HostExcelCreateMappedCopy = FailureJson("E_CONFIRM_REQUIRED", "未输入“生成”，已取消且未修改工作簿")
        Exit Function
    End If

    Dim workbook, outputSheet, outputName, rowIndex, outputRow, outputColumn
    Set workbook = sourceRange.Worksheet.Parent
    outputName = UniqueSheetName(workbook, "字段映射结果")
    Set outputSheet = workbook.Worksheets.Add
    outputSheet.Name = outputName
    If Err.Number <> 0 Then
        Err.Clear
        SafeDeleteSheet appObj, outputSheet
        HostExcelCreateMappedCopy = FailureJson("E_OUTPUT_SHEET", "无法创建字段映射结果工作表")
        Exit Function
    End If

    For outputColumn = 1 To mappingCount
        outputSheet.Cells(1, outputColumn).Value = targetNames(outputColumn)
        outputSheet.Cells(1, outputColumn).NumberFormat = sourceRange.Cells(1, sourceColumns(outputColumn)).NumberFormat
    Next
    outputSheet.Cells(1, mappingCount + 1).Value = "_来源行"
    outputRow = 1
    For rowIndex = 2 To sourceRange.Rows.Count
        outputRow = outputRow + 1
        For outputColumn = 1 To mappingCount
            outputSheet.Cells(outputRow, outputColumn).Value = sourceRange.Cells(rowIndex, sourceColumns(outputColumn)).Value
            outputSheet.Cells(outputRow, outputColumn).NumberFormat = sourceRange.Cells(rowIndex, sourceColumns(outputColumn)).NumberFormat
        Next
        outputSheet.Cells(outputRow, mappingCount + 1).Value = sourceRange.Row + rowIndex - 1
    Next
    outputSheet.Rows(1).Font.Bold = True
    outputSheet.Columns.AutoFit

    If Err.Number <> 0 Then
        Dim writeError
        writeError = Err.Description
        Err.Clear
        SafeDeleteSheet appObj, outputSheet
        HostExcelCreateMappedCopy = FailureJson("E_MAPPING_WRITE", "字段映射副本写入失败，已删除未完成工作表：" & writeError)
        Exit Function
    End If

    Dim summary
    summary = "Excel 字段映射副本已生成；工作表=" & outputName & "；数据行=" & CStr(sourceRange.Rows.Count - 1) & _
        "；输出字段=" & CStr(mappingCount) & "；源数据未修改"
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostExcelCreateMappedCopy = "{""ok"":true,""sourceUnchanged"":true,""outputSheet"":""" & EscapeJson(outputName) & _
        """,""dataRows"":" & CStr(sourceRange.Rows.Count - 1) & ",""outputFields"":" & CStr(mappingCount) & _
        ",""writePlanPreview"":""" & EscapeJson(planPreview) & """,""message"":""" & EscapeJson(summary) & """}"
End Function

Function BuildMappings(sourceRange, mappingText, ByRef sourceColumns, ByRef targetNames, ByRef mappingCount)
    Dim parts, part, pair, sourceName, targetName, sourceColumn, index
    BuildMappings = ""
    mappingCount = 0
    ReDim sourceColumns(sourceRange.Columns.Count)
    ReDim targetNames(sourceRange.Columns.Count)
    mappingText = Replace(Trim(CStr(mappingText)), "，", ",")
    If Len(mappingText) = 0 Then
        For index = 1 To sourceRange.Columns.Count
            mappingCount = mappingCount + 1
            sourceColumns(mappingCount) = index
            targetNames(mappingCount) = Trim(CStr(sourceRange.Cells(1, index).Text))
        Next
        Exit Function
    End If
    parts = Split(mappingText, ",")
    For Each part In parts
        part = Trim(CStr(part))
        pair = Split(part, "=")
        If UBound(pair) <> 1 Then
            BuildMappings = "字段映射格式错误：“" & part & "”；请使用 源字段=目标字段"
            Exit Function
        End If
        sourceName = Trim(CStr(pair(0)))
        targetName = Trim(CStr(pair(1)))
        sourceColumn = FindHeaderColumn(sourceRange, sourceName)
        If sourceColumn <= 0 Then
            BuildMappings = "来源区域未找到字段：" & sourceName
            Exit Function
        End If
        If Len(targetName) = 0 Then
            BuildMappings = "目标字段名称不能为空：" & sourceName
            Exit Function
        End If
        If HasMappedSource(sourceColumns, mappingCount, sourceColumn) Then
            BuildMappings = "同一来源字段不能重复映射：" & sourceName
            Exit Function
        End If
        If HasMappedTarget(targetNames, mappingCount, targetName) Then
            BuildMappings = "目标字段名称不能重复：" & targetName
            Exit Function
        End If
        mappingCount = mappingCount + 1
        sourceColumns(mappingCount) = sourceColumn
        targetNames(mappingCount) = targetName
    Next
    If mappingCount = 0 Then BuildMappings = "至少需要一个字段映射"
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

Function ValidateSourceHeaders(sourceRange)
    Dim colIndex, previousIndex, headerName, previousName
    ValidateSourceHeaders = ""
    For colIndex = 1 To sourceRange.Columns.Count
        headerName = Trim(CStr(sourceRange.Cells(1, colIndex).Text))
        If Len(headerName) = 0 Then
            ValidateSourceHeaders = "第 " & CStr(colIndex) & " 列表头为空，无法安全建立字段映射"
            Exit Function
        End If
        For previousIndex = 1 To colIndex - 1
            previousName = Trim(CStr(sourceRange.Cells(1, previousIndex).Text))
            If StrComp(previousName, headerName, vbTextCompare) = 0 Then
                ValidateSourceHeaders = "来源区域存在重复表头：" & headerName & "；请先改名后再映射"
                Exit Function
            End If
        Next
    Next
End Function

Function HasMappedSource(sourceColumns, mappingCount, sourceColumn)
    Dim index
    HasMappedSource = False
    For index = 1 To mappingCount
        If CLng(sourceColumns(index)) = CLng(sourceColumn) Then HasMappedSource = True: Exit Function
    Next
End Function

Function HasMappedTarget(targetNames, mappingCount, targetName)
    Dim index
    HasMappedTarget = False
    For index = 1 To mappingCount
        If StrComp(CStr(targetNames(index)), CStr(targetName), vbTextCompare) = 0 Then HasMappedTarget = True: Exit Function
    Next
End Function

Function BuildMappingSummary(sourceRange, sourceColumns, targetNames, mappingCount)
    Dim index, text
    text = ""
    For index = 1 To mappingCount
        If Len(text) > 0 Then text = text & "，"
        text = text & CStr(sourceRange.Cells(1, sourceColumns(index)).Text) & "→" & targetNames(index)
    Next
    BuildMappingSummary = text
End Function

Function BuildPlanParams(sourceRange, sourceColumns, targetNames, mappingCount)
    BuildPlanParams = "{""source"":""" & EscapeJson(sourceRange.Worksheet.Name & "!" & sourceRange.Address) & _
        """,""dataRows"":" & CStr(sourceRange.Rows.Count - 1) & ",""outputFields"":" & CStr(mappingCount) & "}"
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
