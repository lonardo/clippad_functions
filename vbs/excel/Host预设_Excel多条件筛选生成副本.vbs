' 函数名: HostExcelMultiConditionFilterCopy

' 描述: 支持最多 3 组字段+关键词+匹配方式，按 AND/OR 组合预览命中后生成新表副本，并保留来源行号；源表不变

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

    Main = HostExcelMultiConditionFilterCopy(appObj)

End Function

Function HostExcelMultiConditionFilterCopy(appObj)

    On Error Resume Next

    Dim sourceRange

    Set sourceRange = appObj.Selection.CurrentRegion

    If Err.Number <> 0 Or TypeName(sourceRange) = "Empty" Or TypeName(sourceRange) = "Nothing" Then

        Err.Clear

        HostExcelMultiConditionFilterCopy = FailureJson("E_NO_RANGE", "请先选中带表头的数据区域")

        Exit Function

    End If

    If sourceRange.Rows.Count < 2 Then

        HostExcelMultiConditionFilterCopy = FailureJson("E_RANGE_TOO_SMALL", "当前区域至少需要一行表头和一行数据")

        Exit Function

    End If

    Dim logicText, logicCode

    logicText = Trim(SafePrompt("条件关系：AND 或 OR", "AND"))

    logicCode = NormalizeLogic(logicText)

    Dim fields(3), keywords(3), modes(3), columns(3)

    Dim condCount, idx, defaultField

    condCount = 0

    defaultField = CStr(sourceRange.Cells(1, 1).Text)

    For idx = 1 To 3

        fields(idx) = Trim(SafePrompt("条件" & CStr(idx) & " 字段名（留空结束）", IIf(idx = 1, defaultField, "")))

        If Len(fields(idx)) = 0 Then Exit For

        columns(idx) = FindHeaderColumn(sourceRange, fields(idx))

        If columns(idx) <= 0 Then

            HostExcelMultiConditionFilterCopy = FailureJson("E_FIELD_NOT_FOUND", "未找到字段：" & fields(idx))

            Exit Function

        End If

        keywords(idx) = SafePrompt("条件" & CStr(idx) & " 关键词", "")

        modes(idx) = NormalizeMatchMode(SafePrompt("条件" & CStr(idx) & " 匹配方式：精确 或 包含", "精确"))

        condCount = condCount + 1

    Next

    If condCount = 0 Then

        HostExcelMultiConditionFilterCopy = FailureJson("E_NO_CONDITION", "至少需要 1 个筛选条件")

        Exit Function

    End If

    Dim rowIndex, matchCount, condSummary

    matchCount = 0

    For rowIndex = 2 To sourceRange.Rows.Count

        If RowMatchesConditions(sourceRange, rowIndex, columns, keywords, modes, condCount, logicCode) Then matchCount = matchCount + 1

    Next

    condSummary = BuildConditionSummary(fields, keywords, modes, condCount, logicCode)

    Dim previewText, planId, planPreview

    previewText = "Excel 多条件筛选预检（尚未写入）" & vbCrLf & _
        "区域=" & sourceRange.Address & "；关系=" & UCase(logicCode) & "；条件数=" & CStr(condCount) & vbCrLf & _
        "条件：" & condSummary & vbCrLf & _
        "数据行=" & CStr(sourceRange.Rows.Count - 1) & "；命中=" & CStr(matchCount) & "；输出=新工作表；源数据不变"

    planId = SafeBeginWritePlan()

    SafeRecordWrite "excel_multi_condition_filter_copy", "{""logic"":""" & EscapeJson(logicCode) & """,""conditions"":" & CStr(condCount) & ",""matches"":" & CStr(matchCount) & "}", "office.excel.filteredCopy"

    planPreview = SafePreviewWritePlan()

    SafeRollbackWritePlan planId

    If StrComp(Trim(SafePrompt(previewText & vbCrLf & vbCrLf & "确认生成多条件筛选副本，请输入：生成", "")), "生成", vbTextCompare) <> 0 Then

        HostExcelMultiConditionFilterCopy = FailureJson("E_CONFIRM_REQUIRED", "未输入“生成”，已取消且未修改工作簿")

        Exit Function

    End If

    Dim workbook, outputSheet, outputName, colIndex, outputRow

    Set workbook = sourceRange.Worksheet.Parent

    outputName = UniqueSheetName(workbook, "多条件筛选")

    Set outputSheet = workbook.Worksheets.Add

    outputSheet.Name = outputName

    If Err.Number <> 0 Then

        Err.Clear

        SafeDeleteSheet appObj, outputSheet

        HostExcelMultiConditionFilterCopy = FailureJson("E_OUTPUT_SHEET", "无法创建筛选结果工作表")

        Exit Function

    End If

    For colIndex = 1 To sourceRange.Columns.Count

        outputSheet.Cells(1, colIndex).Value = sourceRange.Cells(1, colIndex).Value

        outputSheet.Cells(1, colIndex).NumberFormat = sourceRange.Cells(1, colIndex).NumberFormat

    Next

    outputSheet.Cells(1, sourceRange.Columns.Count + 1).Value = "_来源行"

    outputRow = 1

    For rowIndex = 2 To sourceRange.Rows.Count

        If RowMatchesConditions(sourceRange, rowIndex, columns, keywords, modes, condCount, logicCode) Then

            outputRow = outputRow + 1

            For colIndex = 1 To sourceRange.Columns.Count

                outputSheet.Cells(outputRow, colIndex).Value = sourceRange.Cells(rowIndex, colIndex).Value

                outputSheet.Cells(outputRow, colIndex).NumberFormat = sourceRange.Cells(rowIndex, colIndex).NumberFormat

            Next

            outputSheet.Cells(outputRow, sourceRange.Columns.Count + 1).Value = sourceRange.Row + rowIndex - 1

        End If

    Next

    outputSheet.Rows(1).Font.Bold = True

    outputSheet.Columns.AutoFit

    If Err.Number <> 0 Then

        Dim writeError

        writeError = Err.Description

        Err.Clear

        SafeDeleteSheet appObj, outputSheet

        HostExcelMultiConditionFilterCopy = FailureJson("E_COPY_FAILED", "筛选副本写入失败，已删除未完成工作表：" & writeError)

        Exit Function

    End If

    Dim summary

    summary = "Excel 多条件筛选副本已生成；工作表=" & outputName & "；命中行=" & CStr(matchCount) & "；条件=" & condSummary & "；源数据未修改"

    Host.WriteClipboard summary

    SafeWriteLog summary

    HostExcelMultiConditionFilterCopy = "{""ok"":true,""sourceUnchanged"":true,""outputSheet"":""" & EscapeJson(outputName) & _
        """,""matchedRows"":" & CStr(matchCount) & ",""conditionCount"":" & CStr(condCount) & ",""logic"":""" & EscapeJson(logicCode) & _
        """,""writePlanPreview"":""" & EscapeJson(planPreview) & """,""message"":""" & EscapeJson(summary) & """}"

End Function

Function RowMatchesConditions(sourceRange, rowIndex, ByRef columns, ByRef keywords, ByRef modes, condCount, logicCode)

    Dim i, cellText, matched, anyMatched, allMatched

    anyMatched = False

    allMatched = True

    For i = 1 To condCount

        cellText = CStr(sourceRange.Cells(rowIndex, columns(i)).Text)

        matched = IsMatch(cellText, keywords(i), modes(i))

        If matched Then anyMatched = True Else allMatched = False

    Next

    If logicCode = "or" Then

        RowMatchesConditions = anyMatched

    Else

        RowMatchesConditions = allMatched

    End If

End Function

Function BuildConditionSummary(ByRef fields, ByRef keywords, ByRef modes, condCount, logicCode)

    Dim i, parts

    parts = ""

    For i = 1 To condCount

        If i > 1 Then parts = parts & " " & UCase(logicCode) & " "

        parts = parts & fields(i) & "(" & MatchModeLabel(modes(i)) & ")=" & keywords(i)

    Next

    BuildConditionSummary = parts

End Function

Function NormalizeLogic(value)

    Dim text

    text = LCase(Trim(CStr(value)))

    If text = "or" Or text = "或" Then NormalizeLogic = "or" Else NormalizeLogic = "and"

End Function

Function NormalizeMatchMode(value)

    Dim text

    text = LCase(Trim(CStr(value)))

    If text = "包含" Or text = "contains" Then NormalizeMatchMode = "contains" Else NormalizeMatchMode = "exact"

End Function

Function MatchModeLabel(value)

    If value = "contains" Then MatchModeLabel = "包含" Else MatchModeLabel = "精确"

End Function

Function IsMatch(cellText, keyword, matchMode)

    If matchMode = "contains" Then

        IsMatch = (InStr(1, CStr(cellText), CStr(keyword), vbTextCompare) > 0)

    Else

        IsMatch = (StrComp(Trim(CStr(cellText)), Trim(CStr(keyword)), vbTextCompare) = 0)

    End If

End Function

Function IIf(condition, trueValue, falseValue)

    If condition Then IIf = trueValue Else IIf = falseValue

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
