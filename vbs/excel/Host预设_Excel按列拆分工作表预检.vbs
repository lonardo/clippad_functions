' 函数名: HostExcelSplitSheetsByColumnPreflight
' 描述: 按指定列唯一值预览分组后，确认生成多个新工作表副本并保留来源行号；源表不变，空值组可跳过
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
    Main = HostExcelSplitSheetsByColumnPreflight(appObj)
End Function

Function HostExcelSplitSheetsByColumnPreflight(appObj)
    On Error Resume Next

    Dim sourceRange
    Set sourceRange = appObj.Selection.CurrentRegion
    If Err.Number <> 0 Or TypeName(sourceRange) = "Empty" Or TypeName(sourceRange) = "Nothing" Then
        Err.Clear
        HostExcelSplitSheetsByColumnPreflight = FailureJson("E_NO_RANGE", "请先选中带表头的数据区域")
        Exit Function
    End If
    If sourceRange.Rows.Count < 2 Or sourceRange.Columns.Count < 1 Then
        HostExcelSplitSheetsByColumnPreflight = FailureJson("E_RANGE_TOO_SMALL", "当前区域至少需要一行表头和一行数据")
        Exit Function
    End If

    Dim splitField, splitCol, skipEmptyText, skipEmpty, maxSheetsText, maxSheets
    splitField = Trim(SafePrompt("拆分字段名称（按该列唯一值分表）", CStr(sourceRange.Cells(1, 1).Text)))
    If Len(splitField) = 0 Then
        HostExcelSplitSheetsByColumnPreflight = FailureJson("E_FIELD_REQUIRED", "必须指定拆分字段")
        Exit Function
    End If
    splitCol = FindHeaderColumn(sourceRange, splitField)
    If splitCol <= 0 Then
        HostExcelSplitSheetsByColumnPreflight = FailureJson("E_FIELD_NOT_FOUND", "未找到拆分字段：" & splitField)
        Exit Function
    End If

    skipEmptyText = Trim(SafePrompt("空值组：跳过 / 单独成表", "跳过"))
    skipEmpty = (InStr(1, skipEmptyText, "跳", vbTextCompare) > 0) Or (StrComp(skipEmptyText, "skip", vbTextCompare) = 0)
    maxSheetsText = Trim(SafePrompt("最多生成工作表数（防止过多）", "30"))
    If IsNumeric(maxSheetsText) Then
        maxSheets = CLng(maxSheetsText)
    Else
        maxSheets = 30
    End If
    If maxSheets < 1 Then maxSheets = 1
    If maxSheets > 100 Then maxSheets = 100

    Dim keys(), counts(), keyCount, emptyRows, totalRows, rowIndex, key, i
    Dim sampleText, sampleCount, maxGroup
    keyCount = 0
    emptyRows = 0
    totalRows = sourceRange.Rows.Count - 1
    sampleText = ""
    sampleCount = 0
    maxGroup = 0

    For rowIndex = 2 To sourceRange.Rows.Count
        key = NormalizeCellText(sourceRange.Cells(rowIndex, splitCol).Text)
        If Len(key) = 0 Then
            emptyRows = emptyRows + 1
            If Not skipEmpty Then
                AddSplitKey keys, counts, keyCount, "(空值)"
            End If
        Else
            AddSplitKey keys, counts, keyCount, key
        End If
    Next

    For i = 0 To keyCount - 1
        If counts(i) > maxGroup Then maxGroup = counts(i)
        If sampleCount < 8 Then
            sampleText = sampleText & keys(i) & " x" & CStr(counts(i)) & "; "
            sampleCount = sampleCount + 1
        End If
    Next

    If keyCount = 0 Then
        HostExcelSplitSheetsByColumnPreflight = FailureJson("E_NO_GROUPS", "没有可拆分的分组（可能全是空值且选择了跳过）")
        Exit Function
    End If
    If keyCount > maxSheets Then
        HostExcelSplitSheetsByColumnPreflight = FailureJson("E_TOO_MANY_GROUPS", "唯一值组数=" & CStr(keyCount) & " 超过上限 " & CStr(maxSheets) & "，请缩小范围或提高上限")
        Exit Function
    End If

    Dim headersJson, rangeSummary, previewText, planId, planPreview
    headersJson = SafeHostTextExcel("GetHeaders")
    rangeSummary = SafeHostTextExcel("GetRangeSummary")
    previewText = "Excel 按列拆分工作表预检（尚未写入）" & vbCrLf & _
        "区域=" & sourceRange.Address & "；拆分字段=" & splitField & vbCrLf & _
        "数据行=" & CStr(totalRows) & "；分组数=" & CStr(keyCount) & "；空值行=" & CStr(emptyRows) & _
        "；空值策略=" & EmptyPolicyLabel(skipEmpty) & vbCrLf & _
        "最大组内行数=" & CStr(maxGroup) & "；输出=多个新工作表；源数据不变" & vbCrLf & _
        "样例组=" & sampleText & vbCrLf & _
        "Host.GetHeaders: " & headersJson & vbCrLf & _
        "Host.GetRangeSummary: " & rangeSummary

    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_split_sheets_by_column", "{""field"":""" & EscapeJson(splitField) & """,""groups"":" & CStr(keyCount) & ",""skipEmpty"":" & LCase(CStr(skipEmpty)) & "}", "office.excel.splitSheets"
    planPreview = SafePreviewWritePlan()
    SafeRollbackWritePlan planId

    If StrComp(Trim(SafePrompt(previewText & vbCrLf & vbCrLf & "确认按列拆分生成工作表副本，请输入：生成", "")), "生成", vbTextCompare) <> 0 Then
        HostExcelSplitSheetsByColumnPreflight = FailureJson("E_CONFIRM_REQUIRED", "未输入“生成”，已取消且未修改工作簿")
        Exit Function
    End If

    Dim workbook, createdNames, createdCount, groupIndex, outputName, outputSheet
    Dim colIndex, outputRow, memberCount, sourceRowNum
    Set workbook = sourceRange.Worksheet.Parent
    createdNames = ""
    createdCount = 0

    For groupIndex = 0 To keyCount - 1
        outputName = UniqueSheetName(workbook, keys(groupIndex))
        Set outputSheet = workbook.Worksheets.Add
        outputSheet.Name = outputName
        If Err.Number <> 0 Then
            Err.Clear
            SafeDeleteSheet appObj, outputSheet
            HostExcelSplitSheetsByColumnPreflight = FailureJson("E_OUTPUT_SHEET", "无法创建拆分工作表：" & keys(groupIndex))
            Exit Function
        End If

        For colIndex = 1 To sourceRange.Columns.Count
            outputSheet.Cells(1, colIndex).Value = sourceRange.Cells(1, colIndex).Value
            outputSheet.Cells(1, colIndex).NumberFormat = sourceRange.Cells(1, colIndex).NumberFormat
        Next
        outputSheet.Cells(1, sourceRange.Columns.Count + 1).Value = "_来源行"
        outputSheet.Cells(1, sourceRange.Columns.Count + 2).Value = "_拆分值"

        outputRow = 1
        memberCount = 0
        For rowIndex = 2 To sourceRange.Rows.Count
            key = NormalizeCellText(sourceRange.Cells(rowIndex, splitCol).Text)
            If Len(key) = 0 Then
                If skipEmpty Then
                    key = ""
                Else
                    key = "(空值)"
                End If
            End If
            If Len(key) > 0 Then
                If StrComp(key, keys(groupIndex), vbBinaryCompare) = 0 Then
                    outputRow = outputRow + 1
                    memberCount = memberCount + 1
                    For colIndex = 1 To sourceRange.Columns.Count
                        outputSheet.Cells(outputRow, colIndex).Value = sourceRange.Cells(rowIndex, colIndex).Value
                        outputSheet.Cells(outputRow, colIndex).NumberFormat = sourceRange.Cells(rowIndex, colIndex).NumberFormat
                    Next
                    sourceRowNum = sourceRange.Row + rowIndex - 1
                    outputSheet.Cells(outputRow, sourceRange.Columns.Count + 1).Value = sourceRowNum
                    outputSheet.Cells(outputRow, sourceRange.Columns.Count + 2).Value = keys(groupIndex)
                End If
            End If
        Next

        outputSheet.Rows(1).Font.Bold = True
        outputSheet.Columns.AutoFit
        createdCount = createdCount + 1
        If createdCount > 1 Then createdNames = createdNames & "; "
        createdNames = createdNames & outputName & "(" & CStr(memberCount) & ")"
        Err.Clear
    Next

    Dim summary
    summary = "Excel 按列拆分工作表完成；字段=" & splitField & "；新建表=" & CStr(createdCount) & _
        "；源数据不变；表=" & createdNames & vbCrLf & "Preview: " & planPreview
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostExcelSplitSheetsByColumnPreflight = "{""ok"":true,""sourceUnchanged"":true,""splitField"":""" & EscapeJson(splitField) & _
        """,""groupCount"":" & CStr(keyCount) & ",""createdSheets"":" & CStr(createdCount) & _
        ",""message"":""" & EscapeJson(summary) & """}"
End Function

Sub AddSplitKey(ByRef keys, ByRef counts, ByRef keyCount, key)
    Dim i
    For i = 0 To keyCount - 1
        If StrComp(keys(i), key, vbBinaryCompare) = 0 Then
            counts(i) = counts(i) + 1
            Exit Sub
        End If
    Next
    ReDim Preserve keys(keyCount)
    ReDim Preserve counts(keyCount)
    keys(keyCount) = key
    counts(keyCount) = 1
    keyCount = keyCount + 1
End Sub

Function EmptyPolicyLabel(skipEmpty)
    If skipEmpty Then
        EmptyPolicyLabel = "跳过空值组"
    Else
        EmptyPolicyLabel = "空值单独成表"
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
    candidate = Left(SanitizeSheetName(baseName), 31)
    If Len(candidate) = 0 Then candidate = "split"
    suffix = 1
    Do While SheetExists(workbook, candidate)
        suffix = suffix + 1
        candidate = Left(SanitizeSheetName(baseName), 27) & "_" & CStr(suffix)
    Loop
    UniqueSheetName = candidate
End Function

Function SanitizeSheetName(value)
    Dim text
    text = CStr(value)
    text = Replace(text, ":", "_")
    text = Replace(text, "\", "_")
    text = Replace(text, "/", "_")
    text = Replace(text, "?", "_")
    text = Replace(text, "*", "_")
    text = Replace(text, "[", "(")
    text = Replace(text, "]", ")")
    text = Replace(text, "'", "_")
    text = Trim(text)
    If Len(text) = 0 Then text = "empty"
    SanitizeSheetName = text
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
    Else
        SafeHostTextExcel = ""
    End If
    If Err.Number <> 0 Then SafeHostTextExcel = ""
    Err.Clear
End Function
