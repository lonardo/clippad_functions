' 函数名: HostExcelSplitToRowsOrNameCopy

' 描述: 按分隔符把目标列展开到多行，或按中文姓名拆成姓/名；先预览输出行数，确认后写入新表并保留来源行号，源表不变

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

    Main = HostExcelSplitToRowsOrNameCopy(appObj)

End Function

Function HostExcelSplitToRowsOrNameCopy(appObj)

    On Error Resume Next

    Dim sourceRange, selectedRange

    Set selectedRange = appObj.Selection

    Set sourceRange = selectedRange.CurrentRegion

    If Err.Number <> 0 Or TypeName(sourceRange) = "Empty" Or TypeName(sourceRange) = "Nothing" Then

        Err.Clear

        HostExcelSplitToRowsOrNameCopy = FailureJson("E_NO_RANGE", "请先选中带表头的数据区域")

        Exit Function

    End If

    If sourceRange.Rows.Count < 2 Then

        HostExcelSplitToRowsOrNameCopy = FailureJson("E_RANGE_TOO_SMALL", "当前区域至少需要一行表头和一行数据")

        Exit Function

    End If

    Dim modeText, modeCode, fieldName, fieldColumn, delimiter

    modeText = Trim(SafePrompt("模式：分隔符到行 / 姓名拆分", "分隔符到行"))

    modeCode = NormalizeSplitMode(modeText)

    If Len(modeCode) = 0 Then

        HostExcelSplitToRowsOrNameCopy = FailureJson("E_BAD_MODE", "无法识别模式，请输入：分隔符到行 或 姓名拆分")

        Exit Function

    End If

    fieldName = Trim(SafePrompt("目标列字段名（必须与表头一致）", CStr(sourceRange.Cells(1, selectedRange.Column - sourceRange.Column + 1).Text)))

    fieldColumn = FindHeaderColumn(sourceRange, fieldName)

    If fieldColumn <= 0 Then

        HostExcelSplitToRowsOrNameCopy = FailureJson("E_FIELD_NOT_FOUND", "未找到目标字段：" & fieldName)

        Exit Function

    End If

    delimiter = ","

    If modeCode = "delim" Then

        delimiter = SafePrompt("分隔符（例如逗号、分号、斜杠、竖线）", ",")

        If Len(delimiter) = 0 Then

            HostExcelSplitToRowsOrNameCopy = FailureJson("E_EMPTY_DELIMITER", "分隔符为空，已取消")

            Exit Function

        End If

    End If

    Dim rowIndex, outputRows, sampleText, sampleCount, cellText, parts, p

    outputRows = 0

    sampleText = ""

    sampleCount = 0

    For rowIndex = 2 To sourceRange.Rows.Count

        cellText = NormalizeCellText(sourceRange.Cells(rowIndex, fieldColumn).Text)

        If modeCode = "delim" Then

            parts = SplitKeepEmpty(cellText, delimiter)

            If UBound(parts) < 0 Then

                outputRows = outputRows + 1

            Else

                outputRows = outputRows + (UBound(parts) + 1)

            End If

            If sampleCount < 5 And Len(cellText) > 0 Then

                sampleText = sampleText & cellText & " => " & CStr(UBound(parts) + 1) & "行; "

                sampleCount = sampleCount + 1

            End If

        Else

            outputRows = outputRows + 1

            If sampleCount < 5 And Len(cellText) > 0 Then

                sampleText = sampleText & cellText & " => " & SplitChineseNamePreview(cellText) & "; "

                sampleCount = sampleCount + 1

            End If

        End If

    Next

    Dim rangeSummary, budgetJson, previewText, planId, planPreview

    rangeSummary = SafeHostText("GetRangeSummary")

    budgetJson = SafeRunBudget(outputRows)

    previewText = "Excel 分列到行/姓名拆分预检（尚未写入）" & vbCrLf & _
        "模式=" & ModeLabel(modeCode) & "；字段=" & fieldName & "；源数据行=" & CStr(sourceRange.Rows.Count - 1) & "；预计输出行=" & CStr(outputRows) & vbCrLf & _
        "样例=" & sampleText & vbCrLf & _
        "输出=新工作表并保留来源行号；源表不变" & vbCrLf & _
        "Host.GetRangeSummary: " & rangeSummary & vbCrLf & _
        "Host.GetRunBudgetPlan: " & budgetJson

    planId = SafeBeginWritePlan()

    SafeRecordWrite "excel_split_to_rows_or_name", "{""mode"":""" & EscapeJson(modeCode) & """,""field"":""" & EscapeJson(fieldName) & """,""outputRows"":" & CStr(outputRows) & "}", "office.excel.transform"

    planPreview = SafePreviewWritePlan()

    SafeRollbackWritePlan planId

    If StrComp(Trim(SafePrompt(previewText & vbCrLf & vbCrLf & "确认生成拆分副本，请输入：生成", "")), "生成", vbTextCompare) <> 0 Then

        HostExcelSplitToRowsOrNameCopy = FailureJson("E_CONFIRM_REQUIRED", "未输入“生成”，已取消且未修改工作簿")

        Exit Function

    End If

    Dim workbook, outputSheet, outputName, colIndex, outputRow, surname, givenName, partValue

    Set workbook = sourceRange.Worksheet.Parent

    If modeCode = "delim" Then

        outputName = UniqueSheetName(workbook, "分列到行")

    Else

        outputName = UniqueSheetName(workbook, "姓名拆分")

    End If

    Set outputSheet = workbook.Worksheets.Add

    outputSheet.Name = outputName

    If Err.Number <> 0 Then

        Err.Clear

        SafeDeleteSheet appObj, outputSheet

        HostExcelSplitToRowsOrNameCopy = FailureJson("E_OUTPUT_SHEET", "无法创建拆分结果工作表")

        Exit Function

    End If

    For colIndex = 1 To sourceRange.Columns.Count

        outputSheet.Cells(1, colIndex).Value = sourceRange.Cells(1, colIndex).Value

    Next

    outputSheet.Cells(1, sourceRange.Columns.Count + 1).Value = "_来源行"

    If modeCode = "delim" Then

        outputSheet.Cells(1, sourceRange.Columns.Count + 2).Value = "_拆分值"

        outputSheet.Cells(1, sourceRange.Columns.Count + 3).Value = "_拆分段序号"

    Else

        outputSheet.Cells(1, sourceRange.Columns.Count + 2).Value = "_姓"

        outputSheet.Cells(1, sourceRange.Columns.Count + 3).Value = "_名"

    End If

    outputRow = 1

    For rowIndex = 2 To sourceRange.Rows.Count

        cellText = NormalizeCellText(sourceRange.Cells(rowIndex, fieldColumn).Text)

        If modeCode = "delim" Then

            parts = SplitKeepEmpty(cellText, delimiter)

            If UBound(parts) < 0 Then

                outputRow = outputRow + 1

                CopySourceRow sourceRange, rowIndex, outputSheet, outputRow

                outputSheet.Cells(outputRow, sourceRange.Columns.Count + 1).Value = sourceRange.Row + rowIndex - 1

                outputSheet.Cells(outputRow, sourceRange.Columns.Count + 2).Value = ""

                outputSheet.Cells(outputRow, sourceRange.Columns.Count + 3).Value = 1

            Else

                For p = 0 To UBound(parts)

                    partValue = NormalizeCellText(parts(p))

                    outputRow = outputRow + 1

                    CopySourceRow sourceRange, rowIndex, outputSheet, outputRow

                    outputSheet.Cells(outputRow, fieldColumn).Value = partValue

                    outputSheet.Cells(outputRow, sourceRange.Columns.Count + 1).Value = sourceRange.Row + rowIndex - 1

                    outputSheet.Cells(outputRow, sourceRange.Columns.Count + 2).Value = partValue

                    outputSheet.Cells(outputRow, sourceRange.Columns.Count + 3).Value = p + 1

                Next

            End If

        Else

            SplitChineseName cellText, surname, givenName

            outputRow = outputRow + 1

            CopySourceRow sourceRange, rowIndex, outputSheet, outputRow

            outputSheet.Cells(outputRow, sourceRange.Columns.Count + 1).Value = sourceRange.Row + rowIndex - 1

            outputSheet.Cells(outputRow, sourceRange.Columns.Count + 2).Value = surname

            outputSheet.Cells(outputRow, sourceRange.Columns.Count + 3).Value = givenName

        End If

    Next

    outputSheet.Rows(1).Font.Bold = True

    outputSheet.Columns.AutoFit

    If Err.Number <> 0 Then

        Dim writeError

        writeError = Err.Description

        Err.Clear

        SafeDeleteSheet appObj, outputSheet

        HostExcelSplitToRowsOrNameCopy = FailureJson("E_SPLIT_WRITE", "拆分写入失败，已删除未完成工作表：" & writeError)

        Exit Function

    End If

    Dim summary

    summary = "Excel 拆分副本已生成；模式=" & ModeLabel(modeCode) & "；工作表=" & outputName & "；输出行=" & CStr(outputRow - 1) & "；源表未修改"

    Host.WriteClipboard summary

    SafeWriteLog summary

    HostExcelSplitToRowsOrNameCopy = "{""ok"":true,""sourceUnchanged"":true,""mode"":""" & EscapeJson(modeCode) & """,""outputSheet"":""" & EscapeJson(outputName) & _
        """,""outputRows"":" & CStr(outputRow - 1) & ",""writePlanPreview"":""" & EscapeJson(planPreview) & """,""message"":""" & EscapeJson(summary) & """}"

End Function

Sub CopySourceRow(sourceRange, sourceRowIndex, outputSheet, outputRow)

    Dim colIndex

    For colIndex = 1 To sourceRange.Columns.Count

        outputSheet.Cells(outputRow, colIndex).Value = sourceRange.Cells(sourceRowIndex, colIndex).Value

        outputSheet.Cells(outputRow, colIndex).NumberFormat = sourceRange.Cells(sourceRowIndex, colIndex).NumberFormat

    Next

End Sub

Function SplitKeepEmpty(textValue, delimiter)

    Dim text

    text = CStr(textValue)

    If Len(text) = 0 Then

        SplitKeepEmpty = Split("", delimiter)

    Else

        SplitKeepEmpty = Split(text, delimiter)

    End If

End Function

Sub SplitChineseName(fullName, ByRef surname, ByRef givenName)

    Dim text, compound, i, names

    text = NormalizeCellText(fullName)

    surname = ""

    givenName = ""

    If Len(text) = 0 Then Exit Sub

    names = Array("欧阳", "司马", "上官", "诸葛", "东方", "南宫", "西门", "夏侯", "皇甫", "尉迟", "公孙", "慕容", "长孙", "司徒", "司空", "端木", "百里", "呼延", "宇文", "司寇")

    For i = 0 To UBound(names)

        compound = CStr(names(i))

        If Len(text) >= Len(compound) Then

            If StrComp(Left(text, Len(compound)), compound, vbTextCompare) = 0 Then

                surname = Left(text, Len(compound))

                givenName = Mid(text, Len(compound) + 1)

                Exit Sub

            End If

        End If

    Next

    If Len(text) = 1 Then

        surname = text

        givenName = ""

    Else

        surname = Left(text, 1)

        givenName = Mid(text, 2)

    End If

End Sub

Function SplitChineseNamePreview(fullName)

    Dim surname, givenName

    SplitChineseName fullName, surname, givenName

    SplitChineseNamePreview = surname & "/" & givenName

End Function

Function NormalizeSplitMode(value)

    Dim text

    text = LCase(Trim(CStr(value)))

    If InStr(text, "姓名") > 0 Or text = "name" Then

        NormalizeSplitMode = "name"

    ElseIf InStr(text, "分隔") > 0 Or InStr(text, "到行") > 0 Or text = "delim" Or text = "split" Then

        NormalizeSplitMode = "delim"

    Else

        NormalizeSplitMode = ""

    End If

End Function

Function ModeLabel(code)

    If code = "name" Then ModeLabel = "姓名拆分" Else ModeLabel = "分隔符到行"

End Function

Function SafeHostText(methodName)

    On Error Resume Next

    If methodName = "GetRangeSummary" Then

        SafeHostText = Host.GetRangeSummary()

    Else

        SafeHostText = ""

    End If

    If Err.Number <> 0 Then SafeHostText = ""

    Err.Clear

End Function

Function SafeRunBudget(rowCount)

    On Error Resume Next

    SafeRunBudget = Host.GetRunBudgetPlan("excel_split_to_rows", rowCount, 20000, 10, 120)

    If Err.Number <> 0 Then SafeRunBudget = ""

    Err.Clear

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
