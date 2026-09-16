' 函数名: HostExcelSplitColumnAndTrim
' 描述: 将当前 Excel 区域中选中列按分隔符拆分到新工作表，并清理每段首尾空白；适合名单、地址、标签和导入数据拆分
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
        Main = "{""ok"":false,""code"":""E_NO_EXCEL_APP"",""message"":""未取得 Excel 应用，请在 Excel 中运行该预设""}"
        Exit Function
    End If

    Main = HostExcelSplitColumnAndTrim(appObj)
End Function

Function HostExcelSplitColumnAndTrim(appObj)
    On Error Resume Next

    Dim ws, selectedRange, dataRange
    Set ws = appObj.ActiveSheet
    Set selectedRange = appObj.Selection
    If Err.Number <> 0 Or TypeName(ws) = "Empty" Or TypeName(selectedRange) = "Empty" Then
        Err.Clear
        HostExcelSplitColumnAndTrim = "{""ok"":false,""code"":""E_NO_SELECTION"",""message"":""请先选择要拆分的 Excel 列中的一个单元格""}"
        Exit Function
    End If

    Set dataRange = selectedRange.CurrentRegion
    If Err.Number <> 0 Or dataRange.Rows.Count < 1 Then
        Err.Clear
        HostExcelSplitColumnAndTrim = "{""ok"":false,""code"":""E_EMPTY_RANGE"",""message"":""当前区域为空，无法分列""}"
        Exit Function
    End If

    Dim delimiter
    delimiter = SafePrompt("请输入分列分隔符，例如逗号、分号、空格或 /", ",")
    If Len(delimiter) = 0 Then
        HostExcelSplitColumnAndTrim = "{""ok"":false,""code"":""E_EMPTY_DELIMITER"",""message"":""分隔符为空，已取消""}"
        Exit Function
    End If

    Dim sourceColumn
    sourceColumn = selectedRange.Column - dataRange.Column + 1
    If sourceColumn < 1 Or sourceColumn > dataRange.Columns.Count Then sourceColumn = 1

    Dim rangeSummary, budgetJson, planId, previewJson
    rangeSummary = SafeHostText("GetRangeSummary")
    budgetJson = SafeRunBudget(dataRange.Rows.Count)
    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_split_column_trim", "{""scope"":""currentRegion"",""delimiter"":""" & EscapeJson(delimiter) & """}", "office.excel.transform"
    previewJson = SafePreviewWritePlan()

    Dim values(), maxParts, r, parts, p, cleaned
    ReDim values(dataRange.Rows.Count - 1)
    maxParts = 0
    For r = 1 To dataRange.Rows.Count
        cleaned = NormalizeCellText(dataRange.Cells(r, sourceColumn).Text)
        parts = Split(cleaned, delimiter)
        values(r - 1) = parts
        If UBound(parts) + 1 > maxParts Then maxParts = UBound(parts) + 1
    Next
    If maxParts < 1 Then maxParts = 1

    Dim outSheet
    Set outSheet = appObj.Worksheets.Add
    outSheet.Name = SafeSheetName(appObj, "Host分列结果")
    outSheet.Cells(1, 1).Value = "来源行"
    outSheet.Cells(1, 2).Value = "原始文本"
    For p = 1 To maxParts
        outSheet.Cells(1, p + 2).Value = "拆分" & CStr(p)
    Next

    Dim partArray
    For r = 1 To dataRange.Rows.Count
        outSheet.Cells(r + 1, 1).Value = dataRange.Row + r - 1
        outSheet.Cells(r + 1, 2).Value = dataRange.Cells(r, sourceColumn).Text
        partArray = values(r - 1)
        For p = 0 To UBound(partArray)
            outSheet.Cells(r + 1, p + 3).Value = NormalizeCellText(partArray(p))
        Next
    Next
    outSheet.Rows(1).Font.Bold = True
    outSheet.Columns.AutoFit
    SafeCloseWritePlan planId

    Dim summary
    summary = "Excel 分列完成；来源=" & ws.Name & "!" & dataRange.Address(False, False) & _
        "；列序号=" & CStr(sourceColumn) & "；行数=" & CStr(dataRange.Rows.Count) & "；最大拆分段=" & CStr(maxParts) & vbCrLf & _
        "Host.GetRangeSummary: " & rangeSummary & vbCrLf & _
        "Host.GetRunBudgetPlan: " & budgetJson & vbCrLf & _
        "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostExcelSplitColumnAndTrim = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""rows"":" & CStr(dataRange.Rows.Count) & _
        ",""maxParts"":" & CStr(maxParts) & ",""outputSheet"":""" & EscapeJson(outSheet.Name) & """}"

End Function

Function NormalizeCellText(value)
    Dim text
    text = CStr(value)
    text = Replace(text, ChrW(160), " ")
    text = Replace(text, ChrW(12288), " ")
    NormalizeCellText = Trim(text)
End Function

Function SafePrompt(message, defaultValue)
    On Error Resume Next
    SafePrompt = Host.Prompt(message, defaultValue)
    If Err.Number <> 0 Then
        SafePrompt = defaultValue
        Err.Clear
    End If
End Function

Function SafeSheetName(appObj, baseName)
    On Error Resume Next
    Dim candidate, index, exists, ws
    index = 1
    Do
        If index = 1 Then
            candidate = baseName
        Else
            candidate = baseName & CStr(index)
        End If
        exists = False
        For Each ws In appObj.Worksheets
            If StrComp(ws.Name, candidate, vbTextCompare) = 0 Then exists = True
        Next
        If Not exists Then
            SafeSheetName = candidate
            Exit Function
        End If
        index = index + 1
    Loop
End Function

Function SafeHostText(methodName)
    On Error Resume Next
    If methodName = "GetRangeSummary" Then
        SafeHostText = Host.GetRangeSummary()
    Else
        SafeHostText = ""
    End If
    If Err.Number <> 0 Then
        SafeHostText = ""
        Err.Clear
    End If
End Function

Function SafeRunBudget(rowCount)
    On Error Resume Next
    SafeRunBudget = Host.GetRunBudgetPlan("excel_split_column", rowCount, 20000, 10, 120)
    If Err.Number <> 0 Then
        SafeRunBudget = ""
        Err.Clear
    End If
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

Sub SafeCloseWritePlan(planId)
    On Error Resume Next
    Host.RollbackWritePlan planId
    Err.Clear
End Sub

Sub SafeWriteLog(message)
    On Error Resume Next
    Host.WriteLog message
    Err.Clear
End Sub

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
