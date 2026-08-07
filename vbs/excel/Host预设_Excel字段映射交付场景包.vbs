' 函数名: HostExcelFieldMappingDeliveryPack
' 描述: 字段映射交付场景包：按映射生成副本和/或 CSV/TSV，可选敏感字段脱敏；确认后只写输出并生成处理摘要，不改源表
' 适用应用: Excel
' 搜索范围: 当前范围
' 搜索对象: 无
' 作用范围: 当前范围
' 输出类型: 表

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
    Main = HostExcelFieldMappingDeliveryPack(appObj)
End Function

Function HostExcelFieldMappingDeliveryPack(appObj)
    On Error Resume Next
    Dim sourceRange, sourceSheet, sourceName
    Set sourceRange = appObj.Selection.CurrentRegion
    If Err.Number <> 0 Or TypeName(sourceRange) = "Empty" Or TypeName(sourceRange) = "Nothing" Then
        Err.Clear
        HostExcelFieldMappingDeliveryPack = FailureJson("E_NO_RANGE", "请先选中带表头的数据区域")
        Exit Function
    End If
    If sourceRange.Rows.Count < 2 Or sourceRange.Columns.Count < 1 Then
        HostExcelFieldMappingDeliveryPack = FailureJson("E_RANGE_TOO_SMALL", "当前区域至少需要一行表头和一行数据")
        Exit Function
    End If
    Set sourceSheet = sourceRange.Worksheet
    sourceName = sourceSheet.Name

    Dim deliveryMode, outputMode, mappingText, sourceColumns, targetNames, mappingCount, mappingError
    deliveryMode = NormalizeDeliveryMode(SafePrompt("交付方式：sheet / export / both（默认 both）", "both"))
    outputMode = NormalizeOutputMode(SafePrompt("输出模式：NewSheet 或 NewWorkbook（默认 NewSheet）", "NewSheet"))
    mappingText = SafePrompt("字段映射（格式：源字段=目标字段，多个映射用逗号分隔；留空=保留全部字段）", "")
    mappingError = BuildMappings(sourceRange, mappingText, sourceColumns, targetNames, mappingCount)
    If Len(mappingError) > 0 Then
        HostExcelFieldMappingDeliveryPack = FailureJson("E_FIELD_MAPPING", mappingError)
        Exit Function
    End If

    Dim maskEnabled, maskField, maskKind, maskAllFields, maskCol, matchedMaskCount, exampleMask
    Dim r, c, rawText, maskedText
    maskEnabled = False
    matchedMaskCount = 0
    exampleMask = ""
    maskField = Trim(SafePrompt("可选脱敏字段（留空=不脱敏；可输字段名或 全部字段）", ""))
    If Len(maskField) > 0 Then
        maskEnabled = True
        maskAllFields = (maskField = "全部字段" Or LCase(maskField) = "all")
        maskCol = FindHeaderColumn(sourceRange, maskField)
        If (Not maskAllFields) And maskCol <= 0 Then
            HostExcelFieldMappingDeliveryPack = FailureJson("E_MASK_FIELD", "未找到要脱敏的字段：" & maskField)
            Exit Function
        End If
        maskKind = NormalizeMaskKind(SafePrompt("脱敏类型：自动、手机号、身份证或邮箱", "自动"))
        For r = 2 To sourceRange.Rows.Count
            For c = 1 To sourceRange.Columns.Count
                If maskAllFields Or c = maskCol Then
                    rawText = CStr(sourceRange.Cells(r, c).Text)
                    maskedText = MaskSensitiveValue(rawText, maskKind)
                    If Len(maskedText) > 0 Then
                        matchedMaskCount = matchedMaskCount + 1
                        If Len(exampleMask) = 0 Then exampleMask = maskedText
                    End If
                End If
            Next
        Next
    End If

    Dim formatName, delimiter, extension, outputPlan, outputPath, defaultName, needExport, needSheet
    needExport = (deliveryMode = "export" Or deliveryMode = "both")
    needSheet = (deliveryMode = "sheet" Or deliveryMode = "both")
    formatName = "csv"
    delimiter = ","
    extension = ".csv"
    outputPath = ""
    If needExport Then
        formatName = NormalizeExportFormat(SafePrompt("导出格式：CSV 或 TSV", "CSV"))
        If formatName = "tsv" Then
            delimiter = vbTab
            extension = ".tsv"
        End If
        defaultName = "excel_delivery_" & FormatFileStamp(Now()) & extension
        On Error Resume Next
        outputPlan = Host.ResolveOutputPlan(defaultName, "avoid")
        If Err.Number <> 0 Then
            Err.Clear
            outputPlan = ""
        End If
        outputPath = ExtractJsonString(outputPlan, "path")
        If Len(outputPath) = 0 Then
            HostExcelFieldMappingDeliveryPack = FailureJson("E_OUTPUT_PATH", "无法生成避免覆盖的交付文件路径；未导出")
            Exit Function
        End If
    End If

    Dim previewText, planId, planPreview
    previewText = "Excel 字段映射交付场景包预览（尚未写入）" & vbCrLf & _
        "命令ID：excel.field_mapping_delivery_pack" & vbCrLf & _
        "源表=" & sourceName & "；区域=" & sourceRange.Address & vbCrLf & _
        "交付方式=" & deliveryMode & "；输出模式=" & outputMode & vbCrLf & _
        "数据行=" & CStr(sourceRange.Rows.Count - 1) & "；输入字段=" & CStr(sourceRange.Columns.Count) & "；输出字段=" & CStr(mappingCount) & vbCrLf & _
        "映射=" & BuildMappingSummary(sourceRange, sourceColumns, targetNames, mappingCount) & vbCrLf
    If needExport Then previewText = previewText & "导出=" & UCase(formatName) & "；路径=" & outputPath & vbCrLf
    If maskEnabled Then previewText = previewText & "脱敏字段=" & maskField & "；类型=" & MaskKindLabel(maskKind) & "；命中=" & CStr(matchedMaskCount) & "；示例=" & exampleMask & vbCrLf
    previewText = previewText & "sourceUnchanged=true；不改源表"

    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_field_mapping_delivery_pack", "{""source"":""" & EscapeJson(sourceName & "!" & sourceRange.Address) & """,""deliveryMode"":""" & EscapeJson(deliveryMode) & """,""mappingCount"":" & CStr(mappingCount) & ",""maskEnabled"":" & LCase(CStr(maskEnabled)) & ",""exportPath"":""" & EscapeJson(outputPath) & """}", "office.excel.fieldMappingDeliveryPack"
    planPreview = SafePreviewWritePlan()
    SafeRollbackWritePlan planId

    If Not SafeConfirmStep(previewText & vbCrLf & vbCrLf & "确认执行字段映射交付？", "excel.field_mapping_delivery_pack") Then
        HostExcelFieldMappingDeliveryPack = FailureJson("E_CONFIRM_REQUIRED", "用户取消或未确认，未修改工作簿或文件")
        Exit Function
    End If

    Dim createdSheets(), createdCount, mappedSheet, summarySheet, outRows, written, exportText
    createdCount = 0
    ReDim createdSheets(6)
    outRows = 0
    written = False
    Set mappedSheet = Nothing

    If needSheet Then
        Set mappedSheet = CreateOutputSheet(appObj, sourceName, "字段映射", outputMode)
        If mappedSheet Is Nothing Then
            HostExcelFieldMappingDeliveryPack = FailureJson("E_OUTPUT_SHEET", "无法创建字段映射输出工作表")
            Exit Function
        End If
        createdCount = createdCount + 1: Set createdSheets(createdCount) = mappedSheet
        If Not WriteMappedSheet(sourceRange, sourceColumns, targetNames, mappingCount, mappedSheet, maskEnabled, maskField, maskKind, maskAllFields, maskCol, outRows) Then
            RollbackCreatedSheets appObj, createdSheets, createdCount
            HostExcelFieldMappingDeliveryPack = FailureJson("E_MAP_WRITE", "字段映射写入失败，已删除未完成输出")
            Exit Function
        End If
    End If

    If needExport Then
        exportText = BuildDelimitedText(sourceRange, sourceColumns, targetNames, mappingCount, formatName, delimiter, maskEnabled, maskField, maskKind, maskAllFields, maskCol)
        On Error Resume Next
        written = Host.WriteTextFile(outputPath, exportText, False)
        If Err.Number <> 0 Then
            Err.Clear
            written = False
        End If
        If Not CBool(written) Then
            RollbackCreatedSheets appObj, createdSheets, createdCount
            HostExcelFieldMappingDeliveryPack = FailureJson("E_EXPORT_WRITE", "交付文件写入失败，未修改源工作表：" & outputPath)
            Exit Function
        End If
        If outRows = 0 Then outRows = sourceRange.Rows.Count - 1
    End If

    Set summarySheet = CreateOutputSheet(appObj, sourceName, "处理摘要", outputMode)
    If summarySheet Is Nothing Then
        RollbackCreatedSheets appObj, createdSheets, createdCount
        HostExcelFieldMappingDeliveryPack = FailureJson("E_OUTPUT_SHEET", "无法创建处理摘要工作表")
        Exit Function
    End If
    createdCount = createdCount + 1: Set createdSheets(createdCount) = summarySheet

    Dim summaryLines, impactJson, summary, outputNames
    outputNames = ""
    If needSheet Then outputNames = mappedSheet.Name
    If needExport Then
        If Len(outputNames) > 0 Then outputNames = outputNames & ","
        outputNames = outputNames & outputPath
    End If
    If Len(outputNames) > 0 Then outputNames = outputNames & ","
    outputNames = outputNames & summarySheet.Name

    summaryLines = "场景名=字段映射交付场景包" & vbCrLf & _
        "命令ID=excel.field_mapping_delivery_pack" & vbCrLf & _
        "源表名称=" & sourceName & vbCrLf & _
        "源区域=" & sourceRange.Address & vbCrLf & _
        "输出模式=" & outputMode & vbCrLf & _
        "交付方式=" & deliveryMode & vbCrLf & _
        "输出位置=" & outputNames & vbCrLf & _
        "处理前行数=" & CStr(sourceRange.Rows.Count - 1) & vbCrLf & _
        "处理后行数=" & CStr(outRows) & vbCrLf & _
        "参数=mappingCount=" & CStr(mappingCount) & ";maskEnabled=" & CStr(maskEnabled) & vbCrLf & _
        "输出字段数=" & CStr(mappingCount) & vbCrLf & _
        "脱敏命中=" & CStr(matchedMaskCount) & vbCrLf & _
        "导出路径=" & outputPath & vbCrLf & _
        "执行时间=" & Now & vbCrLf & _
        "结果=成功"
    impactJson = SafeHostText("ExcelWriteImpactSummary", summaryLines, summarySheet.Name)
    If Not ExtractJsonBoolean(impactJson, "ok") Then WriteSummaryFallback summarySheet, summaryLines

    summary = "Excel 字段映射交付完成：方式=" & deliveryMode & "；输出字段=" & CStr(mappingCount)
    If needSheet Then summary = summary & "；表=" & mappedSheet.Name
    If needExport Then summary = summary & "；文件=" & outputPath
    summary = summary & "；sourceUnchanged=true"
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostExcelFieldMappingDeliveryPack = "{""ok"":true,""sourceUnchanged"":true,""commandId"":""excel.field_mapping_delivery_pack"",""deliveryMode"":""" & EscapeJson(deliveryMode) & """,""outputSheet"":""" & EscapeJson(IIf(needSheet, mappedSheet.Name, "")) & """,""exportPath"":""" & EscapeJson(outputPath) & """,""summarySheet"":""" & EscapeJson(summarySheet.Name) & """,""outputRows"":" & CStr(outRows) & ",""mappingCount"":" & CStr(mappingCount) & ",""writePlanPreview"":""" & EscapeJson(planPreview) & """,""message"":""" & EscapeJson(summary) & """}"
End Function

Function NormalizeDeliveryMode(value)
    Dim t
    t = LCase(Trim(CStr(value)))
    If t = "sheet" Or t = "工作表" Or t = "表" Then
        NormalizeDeliveryMode = "sheet"
    ElseIf t = "export" Or t = "file" Or t = "导出" Or t = "csv" Or t = "tsv" Then
        NormalizeDeliveryMode = "export"
    Else
        NormalizeDeliveryMode = "both"
    End If
End Function

Function BuildMappings(sourceRange, mappingText, ByRef sourceColumns, ByRef targetNames, ByRef mappingCount)
    Dim text, parts, i, pair, eqPos, sourceName, targetName, sourceCol, n
    BuildMappings = ""
    mappingCount = 0
    text = Trim(CStr(mappingText))
    If Len(text) = 0 Then
        n = sourceRange.Columns.Count
        ReDim sourceColumns(n)
        ReDim targetNames(n)
        For i = 1 To n
            sourceColumns(i) = i
            targetNames(i) = NormalizeCellText(sourceRange.Cells(1, i).Text)
            If Len(targetNames(i)) = 0 Then
                BuildMappings = "源表第 " & CStr(i) & " 列表头为空"
                Exit Function
            End If
        Next
        mappingCount = n
        Exit Function
    End If

    text = Replace(Replace(text, "，", ","), "；", ";")
    parts = Split(text, ",")
    n = 0
    ReDim sourceColumns(UBound(parts) + 1)
    ReDim targetNames(UBound(parts) + 1)
    For i = 0 To UBound(parts)
        pair = Trim(parts(i))
        If Len(pair) > 0 Then
            eqPos = InStr(pair, "=")
            If eqPos <= 0 Then eqPos = InStr(pair, "＝")
            If eqPos <= 0 Then eqPos = InStr(pair, ":")
            If eqPos <= 0 Then eqPos = InStr(pair, "：")
            If eqPos <= 0 Then
                sourceName = pair
                targetName = pair
            Else
                sourceName = Trim(Left(pair, eqPos - 1))
                targetName = Trim(Mid(pair, eqPos + 1))
            End If
            sourceCol = FindHeaderColumn(sourceRange, sourceName)
            If sourceCol <= 0 Then
                BuildMappings = "未找到源字段：" & sourceName
                Exit Function
            End If
            If Len(targetName) = 0 Then targetName = sourceName
            If HasMappedSource(sourceColumns, n, sourceCol) Then
                BuildMappings = "源字段重复映射：" & sourceName
                Exit Function
            End If
            If HasMappedTarget(targetNames, n, targetName) Then
                BuildMappings = "目标字段重复：" & targetName
                Exit Function
            End If
            n = n + 1
            sourceColumns(n) = sourceCol
            targetNames(n) = targetName
        End If
    Next
    If n = 0 Then
        BuildMappings = "字段映射为空"
        Exit Function
    End If
    ReDim Preserve sourceColumns(n)
    ReDim Preserve targetNames(n)
    mappingCount = n
End Function

Function HasMappedSource(sourceColumns, mappingCount, sourceColumn)
    Dim i
    HasMappedSource = False
    For i = 1 To mappingCount
        If CLng(sourceColumns(i)) = CLng(sourceColumn) Then
            HasMappedSource = True
            Exit Function
        End If
    Next
End Function

Function HasMappedTarget(targetNames, mappingCount, targetName)
    Dim i
    HasMappedTarget = False
    For i = 1 To mappingCount
        If StrComp(CStr(targetNames(i)), CStr(targetName), vbTextCompare) = 0 Then
            HasMappedTarget = True
            Exit Function
        End If
    Next
End Function

Function BuildMappingSummary(sourceRange, sourceColumns, targetNames, mappingCount)
    Dim i, buf, sourceName
    buf = ""
    For i = 1 To mappingCount
        sourceName = NormalizeCellText(sourceRange.Cells(1, sourceColumns(i)).Text)
        If Len(buf) > 0 Then buf = buf & ", "
        buf = buf & sourceName & "->" & CStr(targetNames(i))
        If Len(buf) > 180 Then
            buf = buf & "..."
            Exit For
        End If
    Next
    BuildMappingSummary = buf
End Function

Function ShouldMaskMappedColumn(sourceRange, sourceCol, maskEnabled, maskField, maskAllFields, maskCol)
    ShouldMaskMappedColumn = False
    If Not maskEnabled Then Exit Function
    If maskAllFields Then
        ShouldMaskMappedColumn = True
        Exit Function
    End If
    If sourceCol = maskCol Then
        ShouldMaskMappedColumn = True
        Exit Function
    End If
    If StrComp(NormalizeCellText(sourceRange.Cells(1, sourceCol).Text), NormalizeCellText(maskField), vbTextCompare) = 0 Then ShouldMaskMappedColumn = True
End Function

Function MappedCellValue(sourceRange, rowIndex, sourceCol, maskEnabled, maskField, maskKind, maskAllFields, maskCol)
    Dim rawText, masked
    rawText = CStr(sourceRange.Cells(rowIndex, sourceCol).Value)
    MappedCellValue = rawText
    If ShouldMaskMappedColumn(sourceRange, sourceCol, maskEnabled, maskField, maskAllFields, maskCol) Then
        masked = MaskSensitiveValue(CStr(sourceRange.Cells(rowIndex, sourceCol).Text), maskKind)
        If Len(masked) > 0 Then MappedCellValue = masked
    End If
End Function

Function WriteMappedSheet(sourceRange, sourceColumns, targetNames, mappingCount, mappedSheet, maskEnabled, maskField, maskKind, maskAllFields, maskCol, ByRef outRows)
    On Error Resume Next
    Dim c, r, outR
    WriteMappedSheet = False
    For c = 1 To mappingCount
        mappedSheet.Cells(1, c).Value = targetNames(c)
    Next
    outR = 1
    For r = 2 To sourceRange.Rows.Count
        outR = outR + 1
        For c = 1 To mappingCount
            mappedSheet.Cells(outR, c).Value = MappedCellValue(sourceRange, r, sourceColumns(c), maskEnabled, maskField, maskKind, maskAllFields, maskCol)
        Next
    Next
    outRows = outR - 1
    mappedSheet.Rows(1).Font.Bold = True
    mappedSheet.Columns.AutoFit
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    WriteMappedSheet = True
End Function

Function BuildDelimitedText(sourceRange, sourceColumns, targetNames, mappingCount, formatName, delimiter, maskEnabled, maskField, maskKind, maskAllFields, maskCol)
    Dim lines, r, c, rowText, valueText
    lines = ""
    rowText = ""
    For c = 1 To mappingCount
        If c > 1 Then rowText = rowText & delimiter
        rowText = rowText & EscapeExportField(CStr(targetNames(c)), formatName)
    Next
    lines = rowText
    For r = 2 To sourceRange.Rows.Count
        rowText = ""
        For c = 1 To mappingCount
            If c > 1 Then rowText = rowText & delimiter
            valueText = CStr(MappedCellValue(sourceRange, r, sourceColumns(c), maskEnabled, maskField, maskKind, maskAllFields, maskCol))
            rowText = rowText & EscapeExportField(valueText, formatName)
        Next
        lines = lines & vbCrLf & rowText
    Next
    BuildDelimitedText = lines
End Function

Function EscapeExportField(value, formatName)
    Dim text
    text = CStr(value)
    If LCase(CStr(formatName)) = "tsv" Then
        text = Replace(text, vbCrLf, " ")
        text = Replace(text, vbCr, " ")
        text = Replace(text, vbLf, " ")
        text = Replace(text, vbTab, " ")
        EscapeExportField = text
        Exit Function
    End If
    On Error Resume Next
    EscapeExportField = Host.CsvEscape(text)
    If Err.Number <> 0 Or Len(CStr(EscapeExportField)) = 0 And Len(text) > 0 Then
        Err.Clear
        text = Replace(text, """", """""")
        If InStr(text, ",") > 0 Or InStr(text, """") > 0 Or InStr(text, vbCr) > 0 Or InStr(text, vbLf) > 0 Then
            EscapeExportField = """" & text & """"
        Else
            EscapeExportField = text
        End If
    End If
End Function


Function IsNothing(value)
    IsNothing = (TypeName(value) = "Nothing" Or TypeName(value) = "Empty")
End Function

Function ValuesEqual(sourceValue, targetValue)
    ValuesEqual = (StrComp(NormalizeCellText(sourceValue), NormalizeCellText(targetValue), vbTextCompare) = 0)
End Function

Function NormalizeExportFormat(value)
    Dim t
    t = LCase(Trim(CStr(value)))
    If t = "tsv" Or t = "tab" Or t = "制表符" Then
        NormalizeExportFormat = "tsv"
    Else
        NormalizeExportFormat = "csv"
    End If
End Function

Function FormatFileStamp(value)
    Dim d
    d = CDate(value)
    FormatFileStamp = Year(d) & Right("0" & Month(d), 2) & Right("0" & Day(d), 2) & "_" & Right("0" & Hour(d), 2) & Right("0" & Minute(d), 2) & Right("0" & Second(d), 2)
End Function

Function NormalizeLogic(value)
    Dim t
    t = UCase(Trim(CStr(value)))
    If t = "OR" Or t = "或" Or t = "或者" Then
        NormalizeLogic = "OR"
    Else
        NormalizeLogic = "AND"
    End If
End Function

Function NormalizeMatchMode(value)
    Dim t
    t = LCase(Trim(CStr(value)))
    If t = "contains" Or t = "包含" Or t = "like" Or t = "fuzzy" Then
        NormalizeMatchMode = "contains"
    Else
        NormalizeMatchMode = "exact"
    End If
End Function

Function MatchModeLabel(value)
    If NormalizeMatchMode(value) = "contains" Then
        MatchModeLabel = "包含"
    Else
        MatchModeLabel = "精确"
    End If
End Function

Function IsMatch(cellText, keyword, matchMode)
    Dim leftText, rightText
    leftText = NormalizeCellText(cellText)
    rightText = NormalizeCellText(keyword)
    If Len(rightText) = 0 Then
        IsMatch = (Len(leftText) = 0)
        Exit Function
    End If
    If NormalizeMatchMode(matchMode) = "contains" Then
        IsMatch = (InStr(1, leftText, rightText, vbTextCompare) > 0)
    Else
        IsMatch = (StrComp(leftText, rightText, vbTextCompare) = 0)
    End If
End Function

Function NormalizeMaskKind(value)
    Dim t
    t = LCase(Trim(CStr(value)))
    If t = "phone" Or t = "mobile" Or t = "手机" Or t = "手机号" Then
        NormalizeMaskKind = "phone"
    ElseIf t = "id" Or t = "idcard" Or t = "身份证" Then
        NormalizeMaskKind = "idcard"
    ElseIf t = "email" Or t = "mail" Or t = "邮箱" Then
        NormalizeMaskKind = "email"
    Else
        NormalizeMaskKind = "auto"
    End If
End Function

Function MaskKindLabel(value)
    Dim k
    k = NormalizeMaskKind(value)
    If k = "phone" Then
        MaskKindLabel = "手机号"
    ElseIf k = "idcard" Then
        MaskKindLabel = "身份证"
    ElseIf k = "email" Then
        MaskKindLabel = "邮箱"
    Else
        MaskKindLabel = "自动"
    End If
End Function

Function DigitsOnly(value)
    Dim i, ch, buf
    buf = ""
    For i = 1 To Len(CStr(value))
        ch = Mid(CStr(value), i, 1)
        If ch >= "0" And ch <= "9" Then buf = buf & ch
    Next
    DigitsOnly = buf
End Function

Function LooksLikeIdCard(value)
    Dim digits, text
    text = UCase(Trim(CStr(value)))
    digits = DigitsOnly(text)
    LooksLikeIdCard = (Len(digits) = 15 Or Len(digits) = 18 Or (Len(text) = 18 And Right(text, 1) = "X"))
End Function

Function MaskSensitiveValue(value, maskKind)
    Dim text, digits, kind, atPos, localPart, domainPart
    MaskSensitiveValue = ""
    text = NormalizeCellText(value)
    If Len(text) = 0 Then Exit Function
    kind = NormalizeMaskKind(maskKind)
    digits = DigitsOnly(text)
    If kind = "auto" Then
        If InStr(text, "@") > 0 Then
            kind = "email"
        ElseIf LooksLikeIdCard(text) Then
            kind = "idcard"
        ElseIf Len(digits) = 11 And Left(digits, 1) = "1" Then
            kind = "phone"
        Else
            Exit Function
        End If
    End If
    If kind = "phone" Then
        If Len(digits) < 7 Then Exit Function
        MaskSensitiveValue = Left(digits, 3) & "****" & Right(digits, 4)
    ElseIf kind = "idcard" Then
        If Len(digits) < 8 Then Exit Function
        If Len(digits) >= 18 Then
            MaskSensitiveValue = Left(digits, 4) & "**********" & Right(digits, 4)
        Else
            MaskSensitiveValue = Left(digits, 3) & "*********" & Right(digits, 3)
        End If
    ElseIf kind = "email" Then
        atPos = InStr(text, "@")
        If atPos <= 1 Then Exit Function
        localPart = Left(text, atPos - 1)
        domainPart = Mid(text, atPos)
        If Len(localPart) <= 1 Then
            MaskSensitiveValue = "*" & domainPart
        Else
            MaskSensitiveValue = Left(localPart, 1) & "***" & domainPart
        End If
    End If
End Function

Function RowHasContent(rowRange)
    Dim cell
    RowHasContent = False
    For Each cell In rowRange.Cells
        If Len(NormalizeCellText(cell.Text)) > 0 Then
            RowHasContent = True
            Exit Function
        End If
    Next
End Function

Function IIf(condition, trueValue, falseValue)
    If condition Then
        IIf = trueValue
    Else
        IIf = falseValue
    End If
End Function


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
    If colonPos <= 0 Then Exit Function
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

Function ExtractJsonNumber(jsonText, keyName)
    Dim marker, keyPos, colonPos, valueText, ch, i, buf
    ExtractJsonNumber = 0
    marker = """" & CStr(keyName) & """"
    keyPos = InStr(1, CStr(jsonText), marker, vbTextCompare)
    If keyPos <= 0 Then Exit Function
    colonPos = InStr(keyPos + Len(marker), CStr(jsonText), ":")
    If colonPos <= 0 Then Exit Function
    valueText = LTrim(Mid(CStr(jsonText), colonPos + 1))
    buf = ""
    For i = 1 To Len(valueText)
        ch = Mid(valueText, i, 1)
        If (ch >= "0" And ch <= "9") Or ch = "-" Or ch = "." Then
            buf = buf & ch
        ElseIf Len(buf) > 0 Then
            Exit For
        End If
    Next
    If Len(buf) > 0 Then ExtractJsonNumber = CDbl(buf)
End Function

Function ParseJsonIntArray(jsonText, keyName)
    Dim marker, keyPos, bracketStart, bracketEnd, body, parts, i, n, values()
    ParseJsonIntArray = Array()
    marker = """" & CStr(keyName) & """"
    keyPos = InStr(1, CStr(jsonText), marker, vbTextCompare)
    If keyPos <= 0 Then Exit Function
    bracketStart = InStr(keyPos, CStr(jsonText), "[")
    If bracketStart <= 0 Then Exit Function
    bracketEnd = InStr(bracketStart + 1, CStr(jsonText), "]")
    If bracketEnd <= bracketStart Then Exit Function
    body = Trim(Mid(CStr(jsonText), bracketStart + 1, bracketEnd - bracketStart - 1))
    If Len(body) = 0 Then Exit Function
    parts = Split(body, ",")
    n = -1
    ReDim values(UBound(parts))
    For i = 0 To UBound(parts)
        If Len(Trim(parts(i))) > 0 Then
            n = n + 1
            values(n) = CLng(Trim(parts(i)))
        End If
    Next
    If n >= 0 Then
        ReDim Preserve values(n)
        ParseJsonIntArray = values
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
    If Err.Number <> 0 Then
        SafePrompt = CStr(defaultValue)
        Err.Clear
    End If
End Function

Function SafeConfirmStep(message, stepInfo)
    On Error Resume Next
    Dim confirmed, answer
    confirmed = Host.ConfirmStep(message, stepInfo)
    If Err.Number = 0 Then
        SafeConfirmStep = CBool(confirmed)
        Exit Function
    End If
    Err.Clear
    answer = Trim(SafePrompt(message & vbCrLf & vbCrLf & "确认继续请输入：确认", ""))
    SafeConfirmStep = (StrComp(answer, "确认", vbTextCompare) = 0)
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

Function SafeHostText(methodName, arg1, arg2, arg3, arg4)
    On Error Resume Next
    SafeHostText = ""
    If methodName = "ExcelGetHeaderMap" Then
        SafeHostText = Host.ExcelGetHeaderMap(arg1)
    ElseIf methodName = "ExcelCreateOutputSheet" Then
        SafeHostText = Host.ExcelCreateOutputSheet(arg1, arg2, arg3, arg4)
    ElseIf methodName = "ExcelWriteImpactSummary" Then
        SafeHostText = Host.ExcelWriteImpactSummary(arg1, arg2)
    ElseIf methodName = "ExcelPreviewRowSamples" Then
        SafeHostText = Host.ExcelPreviewRowSamples(arg1, arg2)
    ElseIf methodName = "ExcelDedupePlan" Then
        SafeHostText = Host.ExcelDedupePlan(arg1, arg2)
    ElseIf methodName = "ExcelValidatePlan" Then
        SafeHostText = Host.ExcelValidatePlan(arg1)
    End If
    If Err.Number <> 0 Then SafeHostText = ""
    Err.Clear
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

Function FindHeaderColumn(sourceRange, fieldName)
    Dim colIndex
    FindHeaderColumn = 0
    For colIndex = 1 To sourceRange.Columns.Count
        If StrComp(NormalizeCellText(sourceRange.Cells(1, colIndex).Text), NormalizeCellText(fieldName), vbTextCompare) = 0 Then
            FindHeaderColumn = colIndex
            Exit Function
        End If
    Next
End Function

Function SplitCsvFields(text)
    Dim cleaned, parts, i, n, values()
    cleaned = Replace(Replace(Trim(CStr(text)), "，", ","), "；", ",")
    If Len(cleaned) = 0 Then
        SplitCsvFields = Array()
        Exit Function
    End If
    parts = Split(cleaned, ",")
    n = -1
    ReDim values(UBound(parts))
    For i = 0 To UBound(parts)
        If Len(Trim(parts(i))) > 0 Then
            n = n + 1
            values(n) = Trim(parts(i))
        End If
    Next
    If n < 0 Then
        SplitCsvFields = Array()
    Else
        ReDim Preserve values(n)
        SplitCsvFields = values
    End If
End Function

Function ResolveHeaderColumns(sourceRange, fieldNames)
    Dim i, colIndex, n, cols()
    ResolveHeaderColumns = Array()
    If Not IsArray(fieldNames) Then Exit Function
    On Error Resume Next
    If UBound(fieldNames) < 0 Then Exit Function
    Err.Clear
    n = -1
    ReDim cols(UBound(fieldNames))
    For i = 0 To UBound(fieldNames)
        colIndex = FindHeaderColumn(sourceRange, fieldNames(i))
        If colIndex <= 0 Then
            ResolveHeaderColumns = Array()
            Exit Function
        End If
        n = n + 1
        cols(n) = colIndex
    Next
    If n >= 0 Then
        ReDim Preserve cols(n)
        ResolveHeaderColumns = cols
    End If
End Function

Function UniqueSheetName(workbook, baseName)
    Dim candidate, index, exists, ws
    index = 1
    Do
        If index = 1 Then
            candidate = Left(CStr(baseName), 31)
        Else
            candidate = Left(CStr(baseName), 28) & CStr(index)
        End If
        exists = False
        For Each ws In workbook.Worksheets
            If StrComp(ws.Name, candidate, vbTextCompare) = 0 Then exists = True
        Next
        If Not exists Then
            UniqueSheetName = candidate
            Exit Function
        End If
        index = index + 1
    Loop
End Function

Function BuildTimestamp()
    Dim d
    d = Now
    BuildTimestamp = Year(d) & Right("0" & Month(d), 2) & Right("0" & Day(d), 2) & "_" & Right("0" & Hour(d), 2) & Right("0" & Minute(d), 2) & Right("0" & Second(d), 2)
End Function

Function BuildOutputSheetName(sourceSheetName, functionName)
    Dim baseName
    baseName = CStr(sourceSheetName) & "_" & CStr(functionName) & "_" & BuildTimestamp()
    baseName = Replace(baseName, "[", "(")
    baseName = Replace(baseName, "]", ")")
    baseName = Replace(baseName, ":", "-")
    baseName = Replace(baseName, "\\", "-")
    baseName = Replace(baseName, "/", "-")
    baseName = Replace(baseName, "?", "")
    baseName = Replace(baseName, "*", "")
    If Len(baseName) > 31 Then baseName = Left(baseName, 31)
    BuildOutputSheetName = baseName
End Function

Function CreateOutputSheet(appObj, sourceSheetName, functionName, outputMode)
    On Error Resume Next
    Dim hostJson, sheetName, workbook, sheetObj, modeText
    Set CreateOutputSheet = Nothing
    modeText = NormalizeOutputMode(outputMode)
    hostJson = SafeHostText("ExcelCreateOutputSheet", functionName, modeText, sourceSheetName, True)
    sheetName = ExtractJsonString(hostJson, "sheetName")
    If Len(sheetName) = 0 Then sheetName = ExtractJsonString(hostJson, "name")
    If Len(sheetName) > 0 Then
        Set sheetObj = Nothing
        Set sheetObj = appObj.ActiveWorkbook.Worksheets(sheetName)
        If Not sheetObj Is Nothing Then
            Set CreateOutputSheet = sheetObj
            Exit Function
        End If
    End If

    Set workbook = appObj.ActiveWorkbook
    If modeText = "NewWorkbook" Then
        Set workbook = appObj.Workbooks.Add
    End If
    sheetName = UniqueSheetName(workbook, BuildOutputSheetName(sourceSheetName, functionName))
    Set sheetObj = workbook.Worksheets.Add
    sheetObj.Name = sheetName
    If Err.Number <> 0 Then
        Err.Clear
        Set CreateOutputSheet = Nothing
        Exit Function
    End If
    Set CreateOutputSheet = sheetObj
End Function

Sub SafeDeleteSheet(appObj, sheetObj)
    On Error Resume Next
    Dim prev
    If sheetObj Is Nothing Then Exit Sub
    prev = appObj.DisplayAlerts
    appObj.DisplayAlerts = False
    sheetObj.Delete
    appObj.DisplayAlerts = prev
    Err.Clear
End Sub

Function NormalizeOutputMode(text)
    Dim t
    t = LCase(Trim(CStr(text)))
    If t = "newworkbook" Or t = "workbook" Or t = "新工作簿" Then
        NormalizeOutputMode = "NewWorkbook"
    Else
        NormalizeOutputMode = "NewSheet"
    End If
End Function

Sub WriteSummaryFallback(sheetObj, summaryLines)
    Dim lines, i
    lines = Split(CStr(summaryLines), vbCrLf)
    sheetObj.Cells(1, 1).Value = "项目"
    sheetObj.Cells(1, 2).Value = "内容"
    For i = 0 To UBound(lines)
        sheetObj.Cells(i + 2, 1).Value = "行" & CStr(i + 1)
        sheetObj.Cells(i + 2, 2).Value = lines(i)
    Next
    sheetObj.Rows(1).Font.Bold = True
    sheetObj.Columns.AutoFit
End Sub

Sub RollbackCreatedSheets(appObj, createdSheets, createdCount)
    Dim i
    For i = createdCount To 1 Step -1
        SafeDeleteSheet appObj, createdSheets(i)
        Set createdSheets(i) = Nothing
    Next
End Sub

Function IsBlankCell(value)
    IsBlankCell = (Len(NormalizeCellText(value)) = 0)
End Function

Function FindWorksheet(workbook, sheetName)
    On Error Resume Next
    Dim ws
    Set FindWorksheet = Nothing
    For Each ws In workbook.Worksheets
        If StrComp(ws.Name, Trim(CStr(sheetName)), vbTextCompare) = 0 Then
            Set FindWorksheet = ws
            Exit Function
        End If
    Next
End Function

Function FindOtherSheetName(workbook, currentName)
    Dim ws
    FindOtherSheetName = ""
    For Each ws In workbook.Worksheets
        If StrComp(ws.Name, currentName, vbTextCompare) <> 0 Then
            FindOtherSheetName = ws.Name
            Exit Function
        End If
    Next
End Function

Function FindKeyRow(dataRange, keyColumn, keyText)
    Dim rowIndex
    FindKeyRow = 0
    For rowIndex = 2 To dataRange.Rows.Count
        If NormalizeKey(dataRange.Cells(rowIndex, keyColumn).Text) = keyText Then
            FindKeyRow = rowIndex
            Exit Function
        End If
    Next
End Function

Function CountKeyOccurrences(dataRange, keyColumn, keyText)
    Dim rowIndex, total
    total = 0
    For rowIndex = 2 To dataRange.Rows.Count
        If NormalizeKey(dataRange.Cells(rowIndex, keyColumn).Text) = keyText Then total = total + 1
    Next
    CountKeyOccurrences = total
End Function

Function TargetAbsoluteRow(targetRange, relativeRow)
    If relativeRow > 0 Then TargetAbsoluteRow = targetRange.Row + relativeRow - 1 Else TargetAbsoluteRow = 0
End Function

