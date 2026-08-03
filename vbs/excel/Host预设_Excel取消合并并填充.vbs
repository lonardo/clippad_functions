' 函数名: HostExcelUnmergeAndFill
' 描述: 取消当前选区中的合并单元格，并用原左上角内容填充拆分后的单元格
' 适用应用: Excel
' 搜索范围: 选区
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

    Main = HostExcelUnmergeAndFill(appObj)
End Function

Function HostExcelUnmergeAndFill(appObj)
    On Error Resume Next

    Dim selectedRange, cell, mergeArea, mergedCount, planId
    Set selectedRange = appObj.Selection
    If Err.Number <> 0 Or TypeName(selectedRange) = "Empty" Or TypeName(selectedRange) = "Nothing" Then
        Err.Clear
        HostExcelUnmergeAndFill = "{""ok"":false,""code"":""E_NO_SELECTION"",""message"":""请先选择需要取消合并的 Excel 区域""}"
        Exit Function
    End If

    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_unmerge_and_fill", "{""scope"":""selection""}", "office.excel.cell.cleanup"
    mergedCount = 0
    For Each cell In selectedRange.Cells
        If cell.MergeCells Then
            Set mergeArea = cell.MergeArea
            If cell.Address = mergeArea.Cells(1, 1).Address Then
                UnmergeAndFill mergeArea
                If Err.Number = 0 Then mergedCount = mergedCount + 1
                Err.Clear
            End If
        End If
        Err.Clear
    Next
    SafeCloseWritePlan planId

    Dim summary
    summary = "Excel 取消合并完成；处理合并区域=" & CStr(mergedCount) & "；已填充原左上角内容"
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostExcelUnmergeAndFill = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""mergedAreaCount"":" & CStr(mergedCount) & "}"
End Function

Sub UnmergeAndFill(mergeArea)
    On Error Resume Next
    Dim originalValue
    originalValue = mergeArea.Cells(1, 1).Value
    mergeArea.UnMerge
    mergeArea.Value = originalValue
    Err.Clear
End Sub

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
    EscapeJson = text
End Function
