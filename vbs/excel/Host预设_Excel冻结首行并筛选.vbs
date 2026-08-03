' 函数名: HostExcelFreezeHeaderAndFilter
' 描述: 对当前 Excel 数据区域一键冻结首行、开启自动筛选并加粗表头；适合清单、台账和导出表快速进入可浏览状态
' 适用应用: Excel
' 搜索范围: 当前范围
' 搜索对象: 无

Option Explicit

Const xlCenter = -4108

Function Main()
    On Error Resume Next

    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = "{""ok"":false,""code"":""E_NO_EXCEL_APP"",""message"":""未取得 Excel 应用，请在 Excel 中运行该预设""}"
        Exit Function
    End If

    Main = HostExcelFreezeHeaderAndFilter(appObj)
End Function

Function HostExcelFreezeHeaderAndFilter(appObj)
    On Error Resume Next

    Dim ws, selectedRange, dataRange
    Set ws = appObj.ActiveSheet
    Set selectedRange = appObj.Selection
    If Err.Number <> 0 Or TypeName(ws) = "Empty" Or TypeName(selectedRange) = "Empty" Then
        Err.Clear
        HostExcelFreezeHeaderAndFilter = "{""ok"":false,""code"":""E_NO_SELECTION"",""message"":""请先选择 Excel 数据区域中的一个单元格""}"
        Exit Function
    End If

    Set dataRange = selectedRange.CurrentRegion
    If Err.Number <> 0 Or dataRange.Rows.Count < 1 Then
        Err.Clear
        HostExcelFreezeHeaderAndFilter = "{""ok"":false,""code"":""E_EMPTY_RANGE"",""message"":""当前区域为空，无法冻结和筛选""}"
        Exit Function
    End If

    Dim rangeSummary, preflightJson, planId, previewJson
    rangeSummary = SafeHostText("GetRangeSummary")
    preflightJson = SafePreflight()
    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_freeze_header_filter", "{""scope"":""currentRegion""}", "office.excel.view"
    previewJson = SafePreviewWritePlan()

    Dim header
    Set header = dataRange.Rows(1)
    header.Font.Bold = True
    header.HorizontalAlignment = xlCenter
    header.Interior.Color = RGB(221, 235, 247)

    appObj.ActiveWindow.FreezePanes = False
    dataRange.Cells(2, 1).Select
    appObj.ActiveWindow.FreezePanes = True
    dataRange.AutoFilter
    dataRange.Columns.AutoFit
    SafeCloseWritePlan planId

    Dim summary
    summary = "Excel 冻结首行并筛选完成；工作表=" & ws.Name & "；区域=" & dataRange.Address(False, False) & _
        "；行数=" & CStr(dataRange.Rows.Count) & "；列数=" & CStr(dataRange.Columns.Count) & vbCrLf & _
        "Host.GetRangeSummary: " & rangeSummary & vbCrLf & _
        "Host.GetRunPreflightPlan: " & preflightJson & vbCrLf & _
        "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostExcelFreezeHeaderAndFilter = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""rows"":" & CStr(dataRange.Rows.Count) & _
        ",""columns"":" & CStr(dataRange.Columns.Count) & ",""planId"":""" & EscapeJson(planId) & """}"
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

Function SafePreflight()
    On Error Resume Next
    SafePreflight = Host.GetRunPreflightPlan("Excel", True, True, False)
    If Err.Number <> 0 Then
        SafePreflight = ""
        Err.Clear
    End If
End Function

Function SafePrompt(message, defaultValue)
    On Error Resume Next
    SafePrompt = Host.Prompt(message, defaultValue)
    If Err.Number <> 0 Then
        SafePrompt = defaultValue
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

Function SafeValidateWritePlan(planId)
    On Error Resume Next
    SafeValidateWritePlan = Host.ValidateWritePlan(planId, False)
    If Err.Number <> 0 Then SafeValidateWritePlan = ""
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
