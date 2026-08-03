' 函数名: HostExcelCleanCurrentRegion
' 描述: 清理当前 Excel 数据区域文本首尾空白、删除重复行并复制清洗摘要；验证 Excel 区域摘要、运行预检和 Host 日志能力
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

    Main = HostExcelCleanCurrentRegion(appObj)
End Function

Function HostExcelCleanCurrentRegion(appObj)
    On Error Resume Next

    Dim ws, selectedRange, dataRange
    Set ws = appObj.ActiveSheet
    Set selectedRange = appObj.Selection
    If Err.Number <> 0 Or TypeName(ws) = "Empty" Or TypeName(selectedRange) = "Empty" Then
        Err.Clear
        HostExcelCleanCurrentRegion = "{""ok"":false,""code"":""E_NO_SELECTION"",""message"":""请先选择 Excel 数据区域中的一个单元格""}"
        Exit Function
    End If

    Set dataRange = selectedRange.CurrentRegion
    If Err.Number <> 0 Or dataRange.Rows.Count < 2 Then
        Err.Clear
        HostExcelCleanCurrentRegion = "{""ok"":false,""code"":""E_SMALL_RANGE"",""message"":""当前区域不足两行，无法清洗""}"
        Exit Function
    End If

    Dim preflightJson, statsJson, beforeRows
    preflightJson = SafeHostText("GetRunPreflightPlan")
    statsJson = SafeHostText("GetTypeStats")
    beforeRows = dataRange.Rows.Count

    Dim changedCells
    changedCells = TrimTextCells(dataRange)
    RemoveDuplicateRows dataRange

    Dim afterRows, summary
    afterRows = selectedRange.CurrentRegion.Rows.Count
    summary = "Excel 数据清洗完成" & vbCrLf & _
        "工作表: " & CStr(ws.Name) & vbCrLf & _
        "区域: " & CStr(dataRange.Address(False, False)) & vbCrLf & _
        "清理文本单元格: " & CStr(changedCells) & vbCrLf & _
        "清洗前行数: " & CStr(beforeRows) & vbCrLf & _
        "清洗后行数: " & CStr(afterRows) & vbCrLf & _
        "Host.GetTypeStats: " & statsJson & vbCrLf & _
        "Host.GetRunPreflightPlan: " & preflightJson
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostExcelCleanCurrentRegion = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""trimmedCells"":" & CStr(changedCells) & _
        ",""beforeRows"":" & CStr(beforeRows) & ",""afterRows"":" & CStr(afterRows) & "}"

End Function

Function TrimTextCells(rng)
    On Error Resume Next
    Dim changed, cell, beforeValue, afterValue
    changed = 0
    For Each cell In rng.Cells
        beforeValue = CStr(cell.Value)
        afterValue = Trim(Replace(beforeValue, ChrW(12288), " "))
        If afterValue <> beforeValue Then
            cell.Value = afterValue
            changed = changed + 1
        End If
    Next
    TrimTextCells = changed
End Function

Sub RemoveDuplicateRows(rng)
    On Error Resume Next
    Dim cols(), i
    ReDim cols(rng.Columns.Count - 1)
    For i = 1 To rng.Columns.Count
        cols(i - 1) = i
    Next
    rng.RemoveDuplicates cols, 1
    Err.Clear
End Sub

Function SafeHostText(methodName)
    On Error Resume Next
    If methodName = "GetRunPreflightPlan" Then
        SafeHostText = Host.GetRunPreflightPlan("Excel", True, True, False)
    ElseIf methodName = "GetTypeStats" Then
        SafeHostText = Host.GetTypeStats()
    Else
        SafeHostText = ""
    End If
    If Err.Number <> 0 Then
        SafeHostText = ""
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
