' 函数名: HostExcelFillDownBlanks
' 描述: 将当前 Excel 区域中的空白单元格按列向下填充最近非空值；适合取消合并后的分类表、透视导出表和台账整理
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

    Main = HostExcelFillDownBlanks(appObj)
End Function

Function HostExcelFillDownBlanks(appObj)
    On Error Resume Next

    Dim ws, selectedRange, dataRange
    Set ws = appObj.ActiveSheet
    Set selectedRange = appObj.Selection
    If Err.Number <> 0 Or TypeName(ws) = "Empty" Or TypeName(selectedRange) = "Empty" Then
        Err.Clear
        HostExcelFillDownBlanks = "{""ok"":false,""code"":""E_NO_SELECTION"",""message"":""请先选择 Excel 数据区域中的一个单元格""}"
        Exit Function
    End If

    Set dataRange = selectedRange.CurrentRegion
    If Err.Number <> 0 Or dataRange.Rows.Count < 2 Then
        Err.Clear
        HostExcelFillDownBlanks = "{""ok"":false,""code"":""E_SMALL_RANGE"",""message"":""当前区域不足两行，无法向下填充""}"
        Exit Function
    End If

    Dim rangeSummary, blankHint, planId, previewJson, filled, colIndex, rowIndex, lastValue, cellText
    rangeSummary = SafeHostText("GetRangeSummary")
    blankHint = SafeHostText("HasBlank")
    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_fill_down_blanks", "{""scope"":""currentRegion""}", "office.excel.cell.cleanup"
    previewJson = SafePreviewWritePlan()

    filled = 0
    For colIndex = 1 To dataRange.Columns.Count
        lastValue = ""
        For rowIndex = 1 To dataRange.Rows.Count
            cellText = Trim(CStr(dataRange.Cells(rowIndex, colIndex).Text))
            If Len(cellText) = 0 Then
                If Len(lastValue) > 0 Then
                    dataRange.Cells(rowIndex, colIndex).Value = lastValue
                    If Err.Number = 0 Then filled = filled + 1
                    Err.Clear
                End If
            Else
                lastValue = dataRange.Cells(rowIndex, colIndex).Value
            End If
        Next
    Next
    SafeCloseWritePlan planId

    Dim summary
    summary = "Excel 向下填充空白完成；填充单元格=" & CStr(filled) & "；区域=" & dataRange.Address(False, False) & vbCrLf & _
        "Host.GetRangeSummary: " & rangeSummary & vbCrLf & _
        "Host.HasBlank: " & CStr(blankHint) & vbCrLf & _
        "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostExcelFillDownBlanks = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""filled"":" & CStr(filled) & _
        ",""planId"":""" & EscapeJson(planId) & """}"
End Function

Function SafeHostText(methodName)
    On Error Resume Next
    If methodName = "GetRangeSummary" Then
        SafeHostText = Host.GetRangeSummary()
    ElseIf methodName = "HasBlank" Then
        SafeHostText = CStr(Host.HasBlank())
    Else
        SafeHostText = ""
    End If
    If Err.Number <> 0 Then
        SafeHostText = ""
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
