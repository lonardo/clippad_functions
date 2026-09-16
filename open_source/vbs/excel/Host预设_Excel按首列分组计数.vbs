' 函数名: HostExcelGroupCountByFirstColumn
' 描述: 按当前 Excel 数据区域首列分组计数并生成汇总工作表；验证 Host.GetHeaders、GetRangeSummary、GetRunBudgetPlan 和剪贴板摘要
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

    Main = HostExcelGroupCountByFirstColumn(appObj)
End Function

Function HostExcelGroupCountByFirstColumn(appObj)
    On Error Resume Next

    Dim ws, selectedRange, dataRange
    Set ws = appObj.ActiveSheet
    Set selectedRange = appObj.Selection
    If Err.Number <> 0 Or TypeName(ws) = "Empty" Or TypeName(selectedRange) = "Empty" Then
        Err.Clear
        HostExcelGroupCountByFirstColumn = "{""ok"":false,""code"":""E_NO_SELECTION"",""message"":""请先选择 Excel 数据区域中的一个单元格""}"
        Exit Function
    End If

    Set dataRange = selectedRange.CurrentRegion
    If Err.Number <> 0 Or dataRange.Rows.Count < 2 Then
        Err.Clear
        HostExcelGroupCountByFirstColumn = "{""ok"":false,""code"":""E_SMALL_RANGE"",""message"":""当前区域不足两行，无法分组计数""}"
        Exit Function
    End If

    Dim headersJson, rangeSummaryJson, budgetJson
    headersJson = SafeHostText("GetHeaders")
    rangeSummaryJson = SafeHostText("GetRangeSummary")
    budgetJson = SafeRunBudget(dataRange.Rows.Count)

    Dim keys(), counts(), groupCount, r, keyText
    groupCount = 0
    For r = 2 To dataRange.Rows.Count
        keyText = Trim(CStr(dataRange.Cells(r, 1).Text))
        If Len(keyText) = 0 Then keyText = "(空白)"
        AddGroupCount keys, counts, groupCount, keyText
    Next

    Dim outSheet
    Set outSheet = appObj.Worksheets.Add
    outSheet.Name = SafeSheetName(appObj, "Host分组计数")
    outSheet.Range("A1").Value = "分组"
    outSheet.Range("B1").Value = "数量"
    For r = 0 To groupCount - 1
        outSheet.Cells(r + 2, 1).Value = keys(r)
        outSheet.Cells(r + 2, 2).Value = counts(r)
    Next
    outSheet.Columns("A:B").AutoFit

    Dim summary
    summary = "Excel 分组计数完成；来源=" & ws.Name & "!" & dataRange.Address(False, False) & _
        "；分组数=" & CStr(groupCount) & vbCrLf & _
        "Host.GetHeaders: " & headersJson & vbCrLf & _
        "Host.GetRangeSummary: " & rangeSummaryJson & vbCrLf & _
        "Host.GetRunBudgetPlan: " & budgetJson
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostExcelGroupCountByFirstColumn = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""groupCount"":" & CStr(groupCount) & _
        ",""sourceRange"":""" & EscapeJson(dataRange.Address(False, False)) & """,""outputSheet"":""" & EscapeJson(outSheet.Name) & """}"

End Function

Sub AddGroupCount(ByRef keys, ByRef counts, ByRef groupCount, keyText)
    Dim i
    For i = 0 To groupCount - 1
        If StrComp(keys(i), keyText, vbTextCompare) = 0 Then
            counts(i) = counts(i) + 1
            Exit Sub
        End If
    Next
    ReDim Preserve keys(groupCount)
    ReDim Preserve counts(groupCount)
    keys(groupCount) = keyText
    counts(groupCount) = 1
    groupCount = groupCount + 1
End Sub

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
    If methodName = "GetHeaders" Then
        SafeHostText = Host.GetHeaders()
    ElseIf methodName = "GetRangeSummary" Then
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
    SafeRunBudget = Host.GetRunBudgetPlan("group_count", rowCount, 20000, 10, 120)
    If Err.Number <> 0 Then
        SafeRunBudget = ""
        Err.Clear
    End If
End Function

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
