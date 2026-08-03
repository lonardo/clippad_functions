' 函数名: HostExcelSortByFirstColumn
' 描述: 按当前区域首列对数据排序，默认升序；对标表格整理插件中的一键排序
' 适用应用: Excel
' 搜索范围: 当前范围
' 搜索对象: 无

Option Explicit

Const xlYes = 1
Const xlAscending = 1
Const xlDescending = 2
Const xlSortColumns = 1

Function Main()
    On Error Resume Next
    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = "{""ok"":false,""code"":""E_NO_EXCEL_APP"",""message"":""未取得 Excel 应用，请在 Excel 中运行该预设""}"
        Exit Function
    End If
    Main = HostExcelSortByFirstColumn(appObj)
End Function

Function HostExcelSortByFirstColumn(appObj)
    On Error Resume Next
    Dim ws, selectedRange, dataRange
    Set ws = appObj.ActiveSheet
    Set selectedRange = appObj.Selection
    If Err.Number <> 0 Or TypeName(ws) = "Empty" Or TypeName(selectedRange) = "Empty" Then
        Err.Clear
        HostExcelSortByFirstColumn = "{""ok"":false,""code"":""E_NO_SELECTION"",""message"":""请先选择 Excel 数据区域中的一个单元格""}"
        Exit Function
    End If
    Set dataRange = selectedRange.CurrentRegion
    If Err.Number <> 0 Or dataRange.Rows.Count < 2 Then
        Err.Clear
        HostExcelSortByFirstColumn = "{""ok"":false,""code"":""E_SMALL_RANGE"",""message"":""当前区域不足两行，无法排序""}"
        Exit Function
    End If

    Dim orderText, orderValue
    orderText = LCase(Trim(SafePrompt("请输入排序方向：asc 或 desc，默认 asc", "asc")))
    If orderText = "desc" Then
        orderValue = xlDescending
    Else
        orderValue = xlAscending
        orderText = "asc"
    End If

    Dim rangeSummary, planId, previewJson
    rangeSummary = SafeHostText("GetRangeSummary")
    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_sort_by_first_column", "{""order"":""" & orderText & """}", "office.excel.transform"
    previewJson = SafePreviewWritePlan()

    dataRange.Sort dataRange.Columns(1), orderValue, , , , , , xlYes, , , xlSortColumns
    SafeCloseWritePlan planId

    Dim summary
    summary = "Excel 首列排序完成；方向=" & orderText & "；区域=" & dataRange.Address(False, False) & vbCrLf & _
        "Host.GetRangeSummary: " & rangeSummary & vbCrLf & "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostExcelSortByFirstColumn = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""order"":""" & orderText & _
        """,""rows"":" & CStr(dataRange.Rows.Count) & ",""planId"":""" & EscapeJson(planId) & """}"
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