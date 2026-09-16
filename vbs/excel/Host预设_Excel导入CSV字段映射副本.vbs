' 函数名: HostExcelImportCsvMappedCopy
' 描述: 预检 CSV 编码、分隔符和字段映射，确认后导入到新工作表；所有输入按文本写入，避免公式解释和覆盖源数据
' 适用应用: Excel
' 搜索范围: 指定范围
' 搜索对象: 无

Option Explicit

Const MaxCsvBytes = 1048576

Function Main()
    On Error Resume Next
    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = FailureJson("E_NO_EXCEL_APP", "未取得 Excel 应用，请在 Excel 中运行该预设")
        Exit Function
    End If
    Main = HostExcelImportCsvMappedCopy(appObj)
End Function

Function HostExcelImportCsvMappedCopy(appObj)
    On Error Resume Next
    Dim csvPath
    csvPath = Trim(SafeSelectFile("选择要预检并导入的 CSV 文件", "CSV 文件|*.csv;*.txt"))
    If Len(csvPath) = 0 Then
        HostExcelImportCsvMappedCopy = FailureJson("E_CSV_REQUIRED", "未选择 CSV 文件；未修改工作簿")
        Exit Function
    End If

    Dim fileInfo, fileSize
    fileInfo = SafeHostText("GetFileInfo", csvPath)
    fileSize = JsonLong(fileInfo, "size")
    If fileSize <= 0 Then
        HostExcelImportCsvMappedCopy = FailureJson("E_CSV_FILE", "无法读取 CSV 文件信息或文件为空")
        Exit Function
    End If
    If fileSize > MaxCsvBytes Then
        HostExcelImportCsvMappedCopy = FailureJson("E_CSV_TOO_LARGE", "CSV 超过 1 MB 预检上限；为避免截断导入，已拒绝写入")
        Exit Function
    End If

    Dim encodingInfo, encodingName
    encodingInfo = SafeInspectTextFileEncoding(csvPath)
    encodingName = ExtractJsonString(encodingInfo, "encoding")
    If Not ExtractJsonBoolean(encodingInfo, "ok") Then
        HostExcelImportCsvMappedCopy = FailureJson("E_CSV_ENCODING_INSPECT", "无法完成 CSV 编码预检；未读取或写入数据")
        Exit Function
    End If
    If Not ExtractJsonBoolean(encodingInfo, "readSupported") Then
        HostExcelImportCsvMappedCopy = FailureJson("E_CSV_ENCODING_UNSUPPORTED", "CSV 编码为 " & encodingName & "，当前不支持安全读取；请另存为 UTF-8、UTF-16LE 或本机 ANSI 后再导入")
        Exit Function
    End If

    Dim delimiter, csvText, rows, parseError
    delimiter = NormalizeDelimiter(SafePrompt("CSV 分隔符：逗号、分号或制表符", "逗号"))
    csvText = SafeReadTextFile(csvPath, 262144)
    If Len(csvText) = 0 Then
        HostExcelImportCsvMappedCopy = FailureJson("E_CSV_READ", "CSV 读取为空；未写入工作簿")
        Exit Function
    End If
    parseError = ""
    Set rows = ParseCsv(csvText, delimiter, parseError)
    If Len(parseError) > 0 Then
        HostExcelImportCsvMappedCopy = FailureJson("E_CSV_PARSE", parseError)
        Exit Function
    End If
    If rows.Count < 2 Then
        HostExcelImportCsvMappedCopy = FailureJson("E_CSV_ROWS", "CSV 至少需要一行表头和一行数据")
        Exit Function
    End If

    Dim headerRow, headerError
    Set headerRow = rows.Item(1)
    headerError = ValidateHeaders(headerRow)
    If Len(headerError) > 0 Then
        HostExcelImportCsvMappedCopy = FailureJson("E_CSV_HEADERS", headerError)
        Exit Function
    End If

    Dim mappingText, sourceColumns, targetNames, mappingCount, mappingError
    mappingText = SafePrompt("字段映射（格式：源字段=目标字段，多个映射用逗号分隔；留空表示保留全部字段）", "")
    mappingError = BuildMappings(headerRow, mappingText, sourceColumns, targetNames, mappingCount)
    If Len(mappingError) > 0 Then
        HostExcelImportCsvMappedCopy = FailureJson("E_FIELD_MAPPING", mappingError)
        Exit Function
    End If

    Dim previewText, planId, planPreview
    previewText = "CSV 导入与字段映射预览（尚未写入）" & vbCrLf & _
        "文件=" & csvPath & "；编码=" & encodingName & "；分隔符=" & DelimiterLabel(delimiter) & vbCrLf & _
        "数据行=" & CStr(rows.Count - 1) & "；输入字段=" & CStr(headerRow.Count) & "；输出字段=" & CStr(mappingCount) & vbCrLf & _
        "映射：" & BuildMappingSummary(headerRow, sourceColumns, targetNames, mappingCount) & vbCrLf & _
        "将创建新工作表；所有 CSV 值按文本写入，源文件和现有工作表不变"
    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_csv_import_mapped_copy", BuildPlanParams(csvPath, encodingName, delimiter, rows.Count - 1, mappingCount), "office.excel.csvMappedImport"
    planPreview = SafePreviewWritePlan()
    SafeRollbackWritePlan planId
    If StrComp(Trim(SafePrompt(previewText & vbCrLf & vbCrLf & "确认导入 CSV 并生成副本，请输入：导入", "")), "导入", vbTextCompare) <> 0 Then
        HostExcelImportCsvMappedCopy = FailureJson("E_CONFIRM_REQUIRED", "未输入“导入”，已取消且未修改工作簿")
        Exit Function
    End If

    Dim workbook, outputSheet, outputName, rowIndex, outputRow, outputColumn, dataRow
    Set workbook = appObj.ActiveWorkbook
    If TypeName(workbook) = "Empty" Or TypeName(workbook) = "Nothing" Then
        HostExcelImportCsvMappedCopy = FailureJson("E_NO_WORKBOOK", "未取得当前 Excel 工作簿")
        Exit Function
    End If
    outputName = UniqueSheetName(workbook, "CSV导入结果")
    Set outputSheet = workbook.Worksheets.Add
    outputSheet.Name = outputName
    If Err.Number <> 0 Then
        Err.Clear
        SafeDeleteSheet appObj, outputSheet
        HostExcelImportCsvMappedCopy = FailureJson("E_OUTPUT_SHEET", "无法创建 CSV 导入结果工作表")
        Exit Function
    End If
    For outputColumn = 1 To mappingCount
        outputSheet.Cells(1, outputColumn).NumberFormat = "@"
        outputSheet.Cells(1, outputColumn).Value = targetNames(outputColumn)
    Next
    outputSheet.Cells(1, mappingCount + 1).Value = "_来源行"
    outputRow = 1
    For rowIndex = 2 To rows.Count
        Set dataRow = rows.Item(rowIndex)
        outputRow = outputRow + 1
        For outputColumn = 1 To mappingCount
            outputSheet.Cells(outputRow, outputColumn).NumberFormat = "@"
            outputSheet.Cells(outputRow, outputColumn).Value = SafeRowValue(dataRow, sourceColumns(outputColumn))
        Next
        outputSheet.Cells(outputRow, mappingCount + 1).Value = rowIndex
    Next
    outputSheet.Rows(1).Font.Bold = True
    outputSheet.Columns.AutoFit
    If Err.Number <> 0 Then
        Dim writeError
        writeError = Err.Description
        Err.Clear
        SafeDeleteSheet appObj, outputSheet
        HostExcelImportCsvMappedCopy = FailureJson("E_CSV_IMPORT_WRITE", "CSV 导入写入失败，已删除未完成工作表：" & writeError)
        Exit Function
    End If

    Dim summary
    summary = "CSV 字段映射副本已生成；工作表=" & outputName & "；编码=" & encodingName & _
        "；数据行=" & CStr(rows.Count - 1) & "；输出字段=" & CStr(mappingCount) & "；源文件和现有数据未修改"
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostExcelImportCsvMappedCopy = "{""ok"":true,""sourceUnchanged"":true,""encoding"":""" & EscapeJson(encodingName) & _
        """,""outputSheet"":""" & EscapeJson(outputName) & """,""dataRows"":" & CStr(rows.Count - 1) & _
        ",""outputFields"":" & CStr(mappingCount) & ",""writePlanPreview"":""" & EscapeJson(planPreview) & """,""message"":""" & EscapeJson(summary) & """}"
End Function

Function ParseCsv(csvText, delimiter, ByRef errorMessage)
    Dim rows, row, fieldValue, i, ch, nextCh, inQuotes
    errorMessage = ""
    Set rows = New Collection
    Set row = New Collection
    fieldValue = ""
    inQuotes = False
    i = 1
    Do While i <= Len(csvText)
        ch = Mid(csvText, i, 1)
        If inQuotes Then
            If ch = """" Then
                If i < Len(csvText) And Mid(csvText, i + 1, 1) = """" Then
                    fieldValue = fieldValue & """"
                    i = i + 2
                Else
                    inQuotes = False
                    i = i + 1
                End If
            Else
                fieldValue = fieldValue & ch
                i = i + 1
            End If
        ElseIf ch = """" Then
            If Len(fieldValue) > 0 Then
                errorMessage = "第 " & CStr(rows.Count + 1) & " 行的引号必须从字段开头开始"
                Exit Do
            End If
            inQuotes = True
            i = i + 1
        ElseIf ch = delimiter Then
            row.Add fieldValue
            fieldValue = ""
            i = i + 1
        ElseIf ch = vbCr Or ch = vbLf Then
            row.Add fieldValue
            rows.Add row
            Set row = New Collection
            fieldValue = ""
            If ch = vbCr And i < Len(csvText) And Mid(csvText, i + 1, 1) = vbLf Then i = i + 1
            i = i + 1
        Else
            fieldValue = fieldValue & ch
            i = i + 1
        End If
    Loop
    If Len(errorMessage) > 0 Then
        Set ParseCsv = rows
        Exit Function
    End If
    If inQuotes Then
        errorMessage = "CSV 存在未闭合的引号，已拒绝导入"
        Set ParseCsv = rows
        Exit Function
    End If
    If Len(fieldValue) > 0 Or row.Count > 0 Then
        row.Add fieldValue
        rows.Add row
    End If
    Set ParseCsv = rows
End Function

Function ValidateHeaders(headerRow)
    Dim colIndex, previousIndex, headerName
    ValidateHeaders = ""
    For colIndex = 1 To headerRow.Count
        headerName = Trim(CStr(headerRow.Item(colIndex)))
        If Len(headerName) = 0 Then
            ValidateHeaders = "第 " & CStr(colIndex) & " 列表头为空，无法安全映射"
            Exit Function
        End If
        For previousIndex = 1 To colIndex - 1
            If StrComp(Trim(CStr(headerRow.Item(previousIndex))), headerName, vbTextCompare) = 0 Then
                ValidateHeaders = "CSV 存在重复表头：" & headerName & "；请先改名后再导入"
                Exit Function
            End If
        Next
    Next
End Function

Function BuildMappings(headerRow, mappingText, ByRef sourceColumns, ByRef targetNames, ByRef mappingCount)
    Dim parts, part, pair, sourceName, targetName, sourceColumn, index
    BuildMappings = ""
    mappingCount = 0
    ReDim sourceColumns(headerRow.Count)
    ReDim targetNames(headerRow.Count)
    mappingText = Replace(Trim(CStr(mappingText)), "，", ",")
    If Len(mappingText) = 0 Then
        For index = 1 To headerRow.Count
            mappingCount = mappingCount + 1
            sourceColumns(mappingCount) = index
            targetNames(mappingCount) = CStr(headerRow.Item(index))
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
        sourceColumn = FindHeaderColumn(headerRow, sourceName)
        If sourceColumn <= 0 Then
            BuildMappings = "CSV 未找到来源字段：" & sourceName
            Exit Function
        End If
        If Len(targetName) = 0 Or HasMappedTarget(targetNames, mappingCount, targetName) Then
            BuildMappings = "目标字段为空或重复：" & targetName
            Exit Function
        End If
        mappingCount = mappingCount + 1
        sourceColumns(mappingCount) = sourceColumn
        targetNames(mappingCount) = targetName
    Next
End Function

Function FindHeaderColumn(headerRow, fieldName)
    Dim colIndex
    FindHeaderColumn = 0
    For colIndex = 1 To headerRow.Count
        If StrComp(Trim(CStr(headerRow.Item(colIndex))), Trim(CStr(fieldName)), vbTextCompare) = 0 Then FindHeaderColumn = colIndex: Exit Function
    Next
End Function

Function HasMappedTarget(targetNames, mappingCount, targetName)
    Dim index
    HasMappedTarget = False
    For index = 1 To mappingCount
        If StrComp(CStr(targetNames(index)), CStr(targetName), vbTextCompare) = 0 Then HasMappedTarget = True: Exit Function
    Next
End Function

Function SafeRowValue(dataRow, columnIndex)
    If CLng(columnIndex) <= dataRow.Count Then SafeRowValue = CStr(dataRow.Item(CLng(columnIndex))) Else SafeRowValue = ""
End Function

Function NormalizeDelimiter(value)
    Dim text
    text = LCase(Trim(CStr(value)))
    If text = "分号" Or text = "semicolon" Then
        NormalizeDelimiter = ";"
    ElseIf text = "制表符" Or text = "tab" Then
        NormalizeDelimiter = vbTab
    Else
        NormalizeDelimiter = ","
    End If
End Function

Function DelimiterLabel(value)
    If value = ";" Then
        DelimiterLabel = "分号"
    ElseIf value = vbTab Then
        DelimiterLabel = "制表符"
    Else
        DelimiterLabel = "逗号"
    End If
End Function

Function BuildMappingSummary(headerRow, sourceColumns, targetNames, mappingCount)
    Dim index, text
    text = ""
    For index = 1 To mappingCount
        If Len(text) > 0 Then text = text & "，"
        text = text & CStr(headerRow.Item(sourceColumns(index))) & "→" & targetNames(index)
    Next
    BuildMappingSummary = text
End Function

Function BuildPlanParams(csvPath, encodingName, delimiter, dataRows, mappingCount)
    BuildPlanParams = "{""csvPath"":""" & EscapeJson(csvPath) & """,""encoding"":""" & EscapeJson(encodingName) & _
        """,""delimiter"":""" & EscapeJson(DelimiterLabel(delimiter)) & """,""dataRows"":" & CStr(dataRows) & ",""outputFields"":" & CStr(mappingCount) & "}"
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

Function SafeSelectFile(title, filter)
    On Error Resume Next
    SafeSelectFile = Host.SelectFile(title, filter)
    If Err.Number <> 0 Then SafeSelectFile = ""
    Err.Clear
End Function

Function SafeReadTextFile(path, maxChars)
    On Error Resume Next
    SafeReadTextFile = Host.ReadTextFile(path, maxChars)
    If Err.Number <> 0 Then SafeReadTextFile = ""
    Err.Clear
End Function

Function SafeInspectTextFileEncoding(path)
    On Error Resume Next
    SafeInspectTextFileEncoding = Host.InspectTextFileEncoding(path)
    If Err.Number <> 0 Then SafeInspectTextFileEncoding = ""
    Err.Clear
End Function

Function SafeHostText(methodName, value)
    On Error Resume Next
    If methodName = "GetFileInfo" Then SafeHostText = Host.GetFileInfo(value) Else SafeHostText = ""
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

Function ExtractJsonBoolean(jsonText, keyName)
    Dim marker, keyPos, colonPos, valueText
    ExtractJsonBoolean = False
    marker = """" & CStr(keyName) & """"
    keyPos = InStr(1, CStr(jsonText), marker, vbTextCompare)
    If keyPos <= 0 Then Exit Function
    colonPos = InStr(keyPos + Len(marker), CStr(jsonText), ":")
    If colonPos <= 0 Then Exit Function
    valueText = LCase(LTrim(Mid(CStr(jsonText), colonPos + 1)))
    ExtractJsonBoolean = (Left(valueText, 4) = "true")
End Function

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

Function JsonLong(jsonText, keyName)
    Dim marker, keyPos, colonPos, valueText
    JsonLong = 0
    marker = """" & CStr(keyName) & """"
    keyPos = InStr(1, CStr(jsonText), marker, vbTextCompare)
    If keyPos <= 0 Then Exit Function
    colonPos = InStr(keyPos + Len(marker), CStr(jsonText), ":")
    If colonPos <= 0 Then Exit Function
    valueText = Trim(Mid(CStr(jsonText), colonPos + 1))
    JsonLong = CLng(Val(valueText))
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
