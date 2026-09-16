' 函数名: HostExcelCreateMaskedCopy
' 描述: 对用户指定字段预览身份证、手机号或邮箱脱敏结果，确认后生成新工作表；不改写源数据
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
    Main = HostExcelCreateMaskedCopy(appObj)
End Function

Function HostExcelCreateMaskedCopy(appObj)
    On Error Resume Next
    Dim sourceRange
    Set sourceRange = appObj.Selection.CurrentRegion
    If Err.Number <> 0 Or TypeName(sourceRange) = "Empty" Or TypeName(sourceRange) = "Nothing" Or sourceRange.Rows.Count < 2 Then
        Err.Clear
        HostExcelCreateMaskedCopy = FailureJson("E_NO_RANGE", "请先选中带表头的数据区域")
        Exit Function
    End If

    Dim fieldName, fieldColumn, maskKind, maskAllFields
    fieldName = Trim(SafePrompt("要脱敏的字段名称（必须与表头一致；输入 全部字段 可扫描全部列）", CStr(sourceRange.Cells(1, 1).Text)))
    maskAllFields = (fieldName = "全部字段" Or LCase(fieldName) = "all")
    fieldColumn = FindHeaderColumn(sourceRange, fieldName)
    If (Not maskAllFields) And fieldColumn <= 0 Then
        HostExcelCreateMaskedCopy = FailureJson("E_FIELD_NOT_FOUND", "未找到要脱敏的字段：" & fieldName)
        Exit Function
    End If
    maskKind = NormalizeMaskKind(SafePrompt("脱敏类型：自动、手机号、身份证或邮箱", "自动"))

    Dim rowIndex, colIndex, rawText, maskedText, matchedCount, exampleMask
    matchedCount = 0
    exampleMask = ""
    For rowIndex = 2 To sourceRange.Rows.Count
        For colIndex = 1 To sourceRange.Columns.Count
            If maskAllFields Or colIndex = fieldColumn Then
                rawText = CStr(sourceRange.Cells(rowIndex, colIndex).Text)
                maskedText = MaskSensitiveValue(rawText, maskKind)
                If Len(maskedText) > 0 Then
                    matchedCount = matchedCount + 1
                    If Len(exampleMask) = 0 Then exampleMask = maskedText
                End If
            End If
        Next
    Next
    If matchedCount = 0 Then
        HostExcelCreateMaskedCopy = FailureJson("E_NO_SENSITIVE_MATCH", "指定字段中未识别到与脱敏类型匹配的值；未创建副本")
        Exit Function
    End If

    Dim previewText, planId, planPreview
    previewText = "Excel 敏感字段脱敏预览（尚未写入）" & vbCrLf & _
        "源工作表：" & sourceRange.Worksheet.Name & "；区域=" & sourceRange.Address & vbCrLf & _
        "字段=" & fieldName & "；类型=" & MaskKindLabel(maskKind) & "；数据行=" & CStr(sourceRange.Rows.Count - 1) & _
        "；识别并脱敏=" & CStr(matchedCount) & vbCrLf & _
        "掩码示例=" & exampleMask & "；输出=新工作表；源数据不变"
    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_sensitive_field_mask_copy", "{""source"":""" & EscapeJson(sourceRange.Worksheet.Name & "!" & sourceRange.Address) & _
        """,""field"":""" & EscapeJson(fieldName) & """,""kind"":""" & maskKind & """,""matched"":" & CStr(matchedCount) & "}", "office.excel.sensitiveFieldMaskCopy"
    planPreview = SafePreviewWritePlan()
    SafeRollbackWritePlan planId
    If StrComp(Trim(SafePrompt(previewText & vbCrLf & vbCrLf & "确认生成脱敏副本，请输入：脱敏", "")), "脱敏", vbTextCompare) <> 0 Then
        HostExcelCreateMaskedCopy = FailureJson("E_CONFIRM_REQUIRED", "未输入“脱敏”，已取消且未修改工作簿")
        Exit Function
    End If

    Dim workbook, outputSheet, outputName, outputRow
    Set workbook = sourceRange.Worksheet.Parent
    outputName = UniqueSheetName(workbook, "脱敏结果")
    Set outputSheet = workbook.Worksheets.Add
    outputSheet.Name = outputName
    If Err.Number <> 0 Then
        Err.Clear
        SafeDeleteSheet appObj, outputSheet
        HostExcelCreateMaskedCopy = FailureJson("E_OUTPUT_SHEET", "无法创建脱敏结果工作表")
        Exit Function
    End If
    For colIndex = 1 To sourceRange.Columns.Count
        outputSheet.Cells(1, colIndex).Value = CStr(sourceRange.Cells(1, colIndex).Value)
        If maskAllFields Or colIndex = fieldColumn Then outputSheet.Cells(1, colIndex).Value = CStr(sourceRange.Cells(1, colIndex).Value) & "（脱敏）"
    Next
    outputSheet.Cells(1, sourceRange.Columns.Count + 1).Value = "_来源行"
    outputRow = 1
    For rowIndex = 2 To sourceRange.Rows.Count
        outputRow = outputRow + 1
        For colIndex = 1 To sourceRange.Columns.Count
            If maskAllFields Or colIndex = fieldColumn Then
                maskedText = MaskSensitiveValue(CStr(sourceRange.Cells(rowIndex, colIndex).Text), maskKind)
                outputSheet.Cells(outputRow, colIndex).NumberFormat = "@"
                If Len(maskedText) > 0 Then
                    outputSheet.Cells(outputRow, colIndex).Value = maskedText
                Else
                    outputSheet.Cells(outputRow, colIndex).Value = CStr(sourceRange.Cells(rowIndex, colIndex).Text)
                End If
            Else
                outputSheet.Cells(outputRow, colIndex).Value = sourceRange.Cells(rowIndex, colIndex).Value
                outputSheet.Cells(outputRow, colIndex).NumberFormat = sourceRange.Cells(rowIndex, colIndex).NumberFormat
            End If
        Next
        outputSheet.Cells(outputRow, sourceRange.Columns.Count + 1).Value = sourceRange.Row + rowIndex - 1
    Next
    outputSheet.Rows(1).Font.Bold = True
    outputSheet.Columns.AutoFit
    If Err.Number <> 0 Then
        Dim writeError
        writeError = Err.Description
        Err.Clear
        SafeDeleteSheet appObj, outputSheet
        HostExcelCreateMaskedCopy = FailureJson("E_MASK_WRITE", "脱敏副本写入失败，已删除未完成工作表：" & writeError)
        Exit Function
    End If

    Dim summary
    summary = "Excel 脱敏副本已生成；工作表=" & outputName & "；字段=" & fieldName & "；脱敏=" & CStr(matchedCount) & "；源数据未修改"
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostExcelCreateMaskedCopy = "{""ok"":true,""sourceUnchanged"":true,""outputSheet"":""" & EscapeJson(outputName) & _
        """,""maskedRows"":" & CStr(matchedCount) & ",""writePlanPreview"":""" & EscapeJson(planPreview) & """,""message"":""" & EscapeJson(summary) & """}"
End Function

Function FindHeaderColumn(sourceRange, fieldName)
    Dim colIndex
    FindHeaderColumn = 0
    For colIndex = 1 To sourceRange.Columns.Count
        If StrComp(Trim(CStr(sourceRange.Cells(1, colIndex).Text)), Trim(CStr(fieldName)), vbTextCompare) = 0 Then FindHeaderColumn = colIndex: Exit Function
    Next
End Function

Function NormalizeMaskKind(value)
    Dim text
    text = LCase(Trim(CStr(value)))
    If text = "手机号" Or text = "phone" Then
        NormalizeMaskKind = "phone"
    ElseIf text = "身份证" Or text = "idcard" Then
        NormalizeMaskKind = "idcard"
    ElseIf text = "邮箱" Or text = "email" Then
        NormalizeMaskKind = "email"
    Else
        NormalizeMaskKind = "auto"
    End If
End Function

Function MaskKindLabel(value)
    If value = "phone" Then
        MaskKindLabel = "手机号"
    ElseIf value = "idcard" Then
        MaskKindLabel = "身份证"
    ElseIf value = "email" Then
        MaskKindLabel = "邮箱"
    Else
        MaskKindLabel = "自动"
    End If
End Function

Function MaskSensitiveValue(value, maskKind)
    Dim text, digits, atPos, localPart, domainPart
    text = Trim(CStr(value))
    MaskSensitiveValue = ""
    If maskKind = "phone" Or maskKind = "auto" Then
        digits = DigitsOnly(text)
        If Len(digits) = 11 And Left(digits, 1) = "1" Then
            MaskSensitiveValue = Left(digits, 3) & "****" & Right(digits, 4)
            Exit Function
        End If
    End If
    If maskKind = "idcard" Or maskKind = "auto" Then
        If LooksLikeIdCard(text) Then
            MaskSensitiveValue = Left(text, 4) & String(Len(text) - 8, "*") & Right(text, 4)
            Exit Function
        End If
    End If
    If maskKind = "email" Or maskKind = "auto" Then
        atPos = InStr(1, text, "@")
        If atPos > 1 And InStr(atPos + 1, text, ".") > atPos + 1 Then
            localPart = Left(text, atPos - 1)
            domainPart = Mid(text, atPos)
            If Len(localPart) = 1 Then MaskSensitiveValue = "*" & domainPart Else MaskSensitiveValue = Left(localPart, 1) & String(Len(localPart) - 1, "*") & domainPart
        End If
    End If
End Function

Function DigitsOnly(value)
    Dim index, ch, digits
    digits = ""
    For index = 1 To Len(CStr(value))
        ch = Mid(CStr(value), index, 1)
        If ch >= "0" And ch <= "9" Then digits = digits & ch
    Next
    DigitsOnly = digits
End Function

Function LooksLikeIdCard(value)
    Dim text, index, ch
    text = UCase(Trim(CStr(value)))
    LooksLikeIdCard = False
    If Len(text) <> 15 And Len(text) <> 18 Then Exit Function
    For index = 1 To Len(text)
        ch = Mid(text, index, 1)
        If index = Len(text) And Len(text) = 18 Then
            If Not ((ch >= "0" And ch <= "9") Or ch = "X") Then Exit Function
        ElseIf Not (ch >= "0" And ch <= "9") Then
            Exit Function
        End If
    Next
    LooksLikeIdCard = True
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
