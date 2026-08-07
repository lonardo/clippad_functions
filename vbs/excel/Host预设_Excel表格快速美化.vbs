' 函数名: HostExcelFormatCurrentTable
' 描述: 对当前 Excel 数据区域一键设置标题行、边框、隔行底色、筛选和列宽；适合日报、清单和导出表快速交付
' 适用应用: Excel
' 搜索范围: 当前范围
' 搜索对象: 无

Option Explicit

Const xlCenter = -4108
Const xlContinuous = 1
Const xlThin = 2
Const xlEdgeLeft = 7
Const xlEdgeTop = 8
Const xlEdgeBottom = 9
Const xlEdgeRight = 10
Const xlInsideVertical = 11
Const xlInsideHorizontal = 12

Function Main()
    On Error Resume Next

    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = "{""ok"":false,""code"":""E_NO_EXCEL_APP"",""message"":""未取得 Excel 应用，请在 Excel 中运行该预设""}"
        Exit Function
    End If

    Main = HostExcelFormatCurrentTable(appObj)
End Function

Function HostExcelFormatCurrentTable(appObj)
    On Error Resume Next

    Dim ws, selectedRange, dataRange
    Set ws = appObj.ActiveSheet
    Set selectedRange = appObj.Selection
    If Err.Number <> 0 Or TypeName(ws) = "Empty" Or TypeName(selectedRange) = "Empty" Then
        Err.Clear
        HostExcelFormatCurrentTable = "{""ok"":false,""code"":""E_NO_SELECTION"",""message"":""请先选择 Excel 数据区域中的一个单元格""}"
        Exit Function
    End If

    Set dataRange = selectedRange.CurrentRegion
    If Err.Number <> 0 Or dataRange.Rows.Count < 1 Or dataRange.Columns.Count < 1 Then
        Err.Clear
        HostExcelFormatCurrentTable = "{""ok"":false,""code"":""E_EMPTY_RANGE"",""message"":""当前区域为空，无法美化""}"
        Exit Function
    End If

    Dim preflightJson, planId, previewJson
    preflightJson = SafePreflight()
    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_format_current_table", "{""scope"":""currentRegion""}", "office.excel.format"
    previewJson = SafePreviewWritePlan()

    FormatTableRange dataRange
    SafeCloseWritePlan planId

    Dim summary
    summary = "Excel 表格快速美化完成；工作表=" & ws.Name & "；区域=" & dataRange.Address(False, False) & _
        "；行数=" & CStr(dataRange.Rows.Count) & "；列数=" & CStr(dataRange.Columns.Count) & vbCrLf & _
        "Host.GetRunPreflightPlan: " & preflightJson & vbCrLf & _
        "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostExcelFormatCurrentTable = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""rows"":" & CStr(dataRange.Rows.Count) & _
        ",""columns"":" & CStr(dataRange.Columns.Count) & ",""planId"":""" & EscapeJson(planId) & """}"

End Function

Sub FormatTableRange(dataRange)
    On Error Resume Next
    Dim header, body, rowIndex
    Set header = dataRange.Rows(1)
    header.Font.Bold = True
    header.Font.Color = RGB(255, 255, 255)
    header.Interior.Color = RGB(31, 78, 121)
    header.HorizontalAlignment = xlCenter
    header.VerticalAlignment = xlCenter

    dataRange.Borders(xlEdgeLeft).LineStyle = xlContinuous
    dataRange.Borders(xlEdgeTop).LineStyle = xlContinuous
    dataRange.Borders(xlEdgeBottom).LineStyle = xlContinuous
    dataRange.Borders(xlEdgeRight).LineStyle = xlContinuous
    dataRange.Borders(xlInsideVertical).LineStyle = xlContinuous
    dataRange.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    dataRange.Borders.Weight = xlThin

    If dataRange.Rows.Count > 1 Then
        Set body = dataRange.Offset(1, 0).Resize(dataRange.Rows.Count - 1, dataRange.Columns.Count)
        body.Font.Color = RGB(31, 31, 31)
        For rowIndex = 1 To body.Rows.Count
            If rowIndex Mod 2 = 0 Then
                body.Rows(rowIndex).Interior.Color = RGB(242, 246, 250)
            Else
                body.Rows(rowIndex).Interior.Color = RGB(255, 255, 255)
            End If
        Next
    End If

    dataRange.Columns.AutoFit
    dataRange.AutoFilter
    Err.Clear
End Sub

Function SafePreflight()
    On Error Resume Next
    SafePreflight = Host.GetRunPreflightPlan("Excel", True, True, False)
    If Err.Number <> 0 Then
        SafePreflight = ""
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
