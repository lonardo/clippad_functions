' 函数名: HostExcelExportMappedDelivery
' 描述: 按用户确认的字段映射把当前区域导出为 CSV 或 TSV，先预览行数、字段、格式和输出路径；不修改工作表
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
    Main = HostExcelExportMappedDelivery(appObj)
End Function

Function HostExcelExportMappedDelivery(appObj)
    On Error Resume Next
    Dim sourceRange
    Set sourceRange = appObj.Selection.CurrentRegion
    If Err.Number <> 0 Or TypeName(sourceRange) = "Empty" Or TypeName(sourceRange) = "Nothing" Or sourceRange.Rows.Count < 2 Then
        Err.Clear
        HostExcelExportMappedDelivery = FailureJson("E_NO_RANGE", "请先选中带表头的数据区域")
        Exit Function
    End If

    Dim formatName, delimiter, extension
    formatName = NormalizeExportFormat(SafePrompt("交付格式：CSV 或 TSV", "CSV"))
    If formatName = "tsv" Then
        delimiter = vbTab
        extension = ".tsv"
    Else
        delimiter = ","
        extension = ".csv"
    End If

    Dim mappingText, sourceColumns, targetNames, mappingCount, mappingError
    mappingText = SafePrompt("字段映射（格式：源字段=目标字段，多个映射用逗号分隔；留空表示保留全部字段）", "")
    mappingError = BuildMappings(sourceRange, mappingText, sourceColumns, targetNames, mappingCount)
    If Len(mappingError) > 0 Then
        HostExcelExportMappedDelivery = FailureJson("E_FIELD_MAPPING", mappingError)
        Exit Function
    End If

    Dim outputPlan, outputPath, defaultName
    defaultName = "excel_delivery_" & FormatFileStamp(Now()) & extension
    outputPlan = Host.ResolveOutputPlan(defaultName, "avoid")
    outputPath = ExtractJsonString(outputPlan, "path")
    If Len(outputPath) = 0 Then
        HostExcelExportMappedDelivery = FailureJson("E_OUTPUT_PATH", "无法生成避免覆盖的交付文件路径；未导出")
        Exit Function
    End If

    Dim previewText, planId, planPreview
    previewText = "Excel 字段映射交付导出预览（尚未写入）" & vbCrLf & _
        "源工作表：" & sourceRange.Worksheet.Name & "；区域=" & sourceRange.Address & vbCrLf & _
        "格式=" & UCase(formatName) & "；数据行=" & CStr(sourceRange.Rows.Count - 1) & "；输入字段=" & CStr(sourceRange.Columns.Count) & "；输出字段=" & CStr(mappingCount) & vbCrLf & _
        "映射：" & BuildMappingSummary(sourceRange, sourceColumns, targetNames, mappingCount) & vbCrLf & _
        "输出路径=" & outputPath & vbCrLf & "源工作表不变；输出文件默认避免覆盖同名文件"
    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_mapped_delivery_export", BuildPlanParams(sourceRange, formatName, outputPath, mappingCount), "host.file.mappedExport"
    planPreview = SafePreviewWritePlan()
    SafeRollbackWritePlan planId
    If StrComp(Trim(SafePrompt(previewText & vbCrLf & vbCrLf & "确认导出交付文件，请输入：导出", "")), "导出", vbTextCompare) <> 0 Then
        HostExcelExportMappedDelivery = FailureJson("E_CONFIRM_REQUIRED", "未输入“导出”，已取消且未修改工作表或文件")
        Exit Function
    End If

    Dim exportText, written
    exportText = BuildDelimitedText(sourceRange, sourceColumns, targetNames, mappingCount, formatName, delimiter)
    written = Host.WriteTextFile(outputPath, exportText, False)
    If Not CBool(written) Then
        HostExcelExportMappedDelivery = FailureJson("E_EXPORT_WRITE", "交付文件写入失败，未修改源工作表：" & outputPath)
        Exit Function
    End If

    Dim summary
    summary = "Excel 字段映射交付文件已导出；格式=" & UCase(formatName) & "；路径=" & outputPath & _
        "；数据行=" & CStr(sourceRange.Rows.Count - 1) & "；输出字段=" & CStr(mappingCount) & "；源工作表未修改"
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostExcelExportMappedDelivery = "{""ok"":true,""sourceUnchanged"":true,""format"":""" & formatName & _
        """,""outputPath"":""" & EscapeJson(outputPath) & """,""dataRows"":" & CStr(sourceRange.Rows.Count - 1) & _
        ",""outputFields"":" & CStr(mappingCount) & ",""writePlanPreview"":""" & EscapeJson(planPreview) & """,""message"":""" & EscapeJson(summary) & """}"
End Function

Function BuildDelimitedText(sourceRange, sourceColumns, targetNames, mappingCount, formatName, delimiter)
    Dim rowIndex, outputColumn, line, text
    text = ""
    line = ""
    For outputColumn = 1 To mappingCount
        If outputColumn > 1 Then line = line & delimiter
        line = line & EscapeExportField(targetNames(outputColumn), formatName)
    Next
    text = line & vbCrLf
    For rowIndex = 2 To sourceRange.Rows.Count
        line = ""
        For outputColumn = 1 To mappingCount
            If outputColumn > 1 Then line = line & delimiter
            line = line & EscapeExportField(CStr(sourceRange.Cells(rowIndex, sourceColumns(outputColumn)).Text), formatName)
        Next
        text = text & line & vbCrLf
    Next
    BuildDelimitedText = text
End Function

Function EscapeExportField(value, formatName)
    Dim text
    text = CStr(value)
    If formatName = "csv" Then
        EscapeExportField = Host.CsvEscape(text)
    Else
        text = Replace(text, Chr(34), Chr(34) & Chr(34))
        EscapeExportField = """" & text & """"
    End If
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
            sourceName = Trim(CStr(sourceRange.Cells(1, index).Text))
            If Len(sourceName) = 0 Then
                BuildMappings = "第 " & CStr(index) & " 列表头为空，无法安全导出"
                Exit Function
            End If
            mappingCount = mappingCount + 1
            sourceColumns(mappingCount) = index
            targetNames(mappingCount) = sourceName
        Next
        Exit Function
    End If
    parts = Split(mappingText, ",")
    For Each part In parts
        pair = Split(Trim(CStr(part)), "=")
        If UBound(pair) <> 1 Then
            BuildMappings = "字段映射格式错误：“" & CStr(part) & "”；请使用 源字段=目标字段"
            Exit Function
        End If
        sourceName = Trim(CStr(pair(0)))
        targetName = Trim(CStr(pair(1)))
        sourceColumn = FindHeaderColumn(sourceRange, sourceName)
        If sourceColumn <= 0 Then
            BuildMappings = "来源区域未找到字段：" & sourceName
            Exit Function
        End If
        If Len(targetName) = 0 Or HasMappedSource(sourceColumns, mappingCount, sourceColumn) Or HasMappedTarget(targetNames, mappingCount, targetName) Then
            BuildMappings = "来源字段重复或目标字段为空/重复：" & sourceName & "=" & targetName
            Exit Function
        End If
        mappingCount = mappingCount + 1
        sourceColumns(mappingCount) = sourceColumn
        targetNames(mappingCount) = targetName
    Next
End Function

Function FindHeaderColumn(sourceRange, fieldName)
    Dim colIndex
    FindHeaderColumn = 0
    For colIndex = 1 To sourceRange.Columns.Count
        If StrComp(Trim(CStr(sourceRange.Cells(1, colIndex).Text)), Trim(CStr(fieldName)), vbTextCompare) = 0 Then FindHeaderColumn = colIndex: Exit Function
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

Function NormalizeExportFormat(value)
    If LCase(Trim(CStr(value))) = "tsv" Or Trim(CStr(value)) = "制表符" Then NormalizeExportFormat = "tsv" Else NormalizeExportFormat = "csv"
End Function

Function FormatFileStamp(value)
    FormatFileStamp = CStr(Year(value)) & Right("0" & CStr(Month(value)), 2) & Right("0" & CStr(Day(value)), 2) & "_" & Right("0" & CStr(Hour(value)), 2) & Right("0" & CStr(Minute(value)), 2) & Right("0" & CStr(Second(value)), 2)
End Function

Function BuildPlanParams(sourceRange, formatName, outputPath, mappingCount)
    BuildPlanParams = "{""source"":""" & EscapeJson(sourceRange.Worksheet.Name & "!" & sourceRange.Address) & _
        """,""format"":""" & formatName & """,""outputPath"":""" & EscapeJson(outputPath) & _
        """,""dataRows"":" & CStr(sourceRange.Rows.Count - 1) & ",""outputFields"":" & CStr(mappingCount) & "}"
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
    Host.RecordWrite actionId, paramsJson, "host", "", "", capabilityId
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

Function ExtractJsonString(jsonText, keyName)
    Dim marker, keyPos, colonPos, quotePos, i, ch, nextCh, result
    ExtractJsonString = ""
    marker = """" & CStr(keyName) & """"
    keyPos = InStr(1, CStr(jsonText), marker, vbTextCompare)
    If keyPos <= 0 Then Exit Function
    colonPos = InStr(keyPos + Len(marker), CStr(jsonText), ":")
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
