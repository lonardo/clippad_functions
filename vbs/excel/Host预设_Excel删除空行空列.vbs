' 函数名: HostExcelRemoveBlankRowsColumns
' 描述: 删除当前 Excel 数据区域中的全空行和全空列，并复制清理摘要；适合导入表、问卷表和系统导出表整理
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

    Main = HostExcelRemoveBlankRowsColumns(appObj)
End Function

Function HostExcelRemoveBlankRowsColumns(appObj)
    On Error Resume Next

    Dim ws, selectedRange, dataRange
    Set ws = appObj.ActiveSheet
    Set selectedRange = appObj.Selection
    If Err.Number <> 0 Or TypeName(ws) = "Empty" Or TypeName(selectedRange) = "Empty" Then
        Err.Clear
        HostExcelRemoveBlankRowsColumns = "{""ok"":false,""code"":""E_NO_SELECTION"",""message"":""请先选择 Excel 数据区域中的一个单元格""}"
        Exit Function
    End If

    Set dataRange = selectedRange.CurrentRegion
    If Err.Number <> 0 Or dataRange.Rows.Count < 1 Or dataRange.Columns.Count < 1 Then
        Err.Clear
        HostExcelRemoveBlankRowsColumns = "{""ok"":false,""code"":""E_EMPTY_RANGE"",""message"":""当前区域为空，无法清理""}"
        Exit Function
    End If

    Dim rangeSummary, planId, previewJson
    rangeSummary = SafeHostText("GetRangeSummary")
    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_remove_blank_rows_columns", "{""scope"":""currentRegion""}", "office.excel.clean"
    previewJson = SafePreviewWritePlan()

    Dim beforeRows, beforeCols, removedRows, removedCols
    beforeRows = dataRange.Rows.Count
    beforeCols = dataRange.Columns.Count
    removedRows = RemoveBlankRows(appObj, dataRange)
    Set dataRange = selectedRange.CurrentRegion
    removedCols = RemoveBlankColumns(appObj, dataRange)
    SafeCloseWritePlan planId

    Dim afterRows, afterCols, summary
    Set dataRange = selectedRange.CurrentRegion
    afterRows = dataRange.Rows.Count
    afterCols = dataRange.Columns.Count
    summary = "Excel 空行空列清理完成；工作表=" & ws.Name & _
        "；删除空行=" & CStr(removedRows) & "；删除空列=" & CStr(removedCols) & _
        "；清理前=" & CStr(beforeRows) & "x" & CStr(beforeCols) & _
        "；清理后=" & CStr(afterRows) & "x" & CStr(afterCols) & vbCrLf & _
        "Host.GetRangeSummary: " & rangeSummary & vbCrLf & _
        "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostExcelRemoveBlankRowsColumns = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""removedRows"":" & CStr(removedRows) & _
        ",""removedColumns"":" & CStr(removedCols) & ",""afterRows"":" & CStr(afterRows) & ",""afterColumns"":" & CStr(afterCols) & "}"

End Function

Function RemoveBlankRows(appObj, dataRange)
    On Error Resume Next
    Dim r, count
    count = 0
    For r = dataRange.Rows.Count To 1 Step -1
        If appObj.WorksheetFunction.CountA(dataRange.Rows(r)) = 0 Then
            dataRange.Rows(r).Delete
            count = count + 1
        End If
        If Err.Number <> 0 Then Err.Clear
    Next
    RemoveBlankRows = count
End Function

Function RemoveBlankColumns(appObj, dataRange)
    On Error Resume Next
    Dim c, count
    count = 0
    For c = dataRange.Columns.Count To 1 Step -1
        If appObj.WorksheetFunction.CountA(dataRange.Columns(c)) = 0 Then
            dataRange.Columns(c).Delete
            count = count + 1
        End If
        If Err.Number <> 0 Then Err.Clear
    Next
    RemoveBlankColumns = count
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
