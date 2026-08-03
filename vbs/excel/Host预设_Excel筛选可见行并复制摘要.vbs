' 函数名: HostExcelFilterVisibleRows
' 描述: 对当前 Excel 选区所在数据区域按当前列筛选关键字，并复制筛选摘要；验证 Host.Prompt、Excel 上下文、表头/区域摘要与剪贴板能力
' 适用应用: Excel
' 搜索范围: 当前范围
' 搜索对象: 无

Option Explicit

Const xlCellTypeVisible = 12

Function Main()
    On Error Resume Next

    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = "{""ok"":false,""code"":""E_NO_EXCEL_APP"",""message"":""未取得 Excel 应用，请在 Excel 中运行该预设""}"
        Exit Function
    End If

    Main = HostExcelFilterVisibleRows(appObj)
End Function

Function HostExcelFilterVisibleRows(appObj)
    On Error Resume Next

    Dim ws, selectedRange, dataRange
    Set ws = appObj.ActiveSheet
    Set selectedRange = appObj.Selection
    If Err.Number <> 0 Or TypeName(ws) = "Empty" Or TypeName(selectedRange) = "Empty" Then
        Err.Clear
        HostExcelFilterVisibleRows = "{""ok"":false,""code"":""E_NO_SELECTION"",""message"":""请先选择 Excel 数据区域中的一个单元格""}"
        Exit Function
    End If

    Set dataRange = selectedRange.CurrentRegion
    If Err.Number <> 0 Or dataRange.Rows.Count < 2 Then
        Err.Clear
        HostExcelFilterVisibleRows = "{""ok"":false,""code"":""E_SMALL_RANGE"",""message"":""当前区域不足两行，无法执行筛选""}"
        Exit Function
    End If

    Dim relativeColumn
    relativeColumn = selectedRange.Column - dataRange.Column + 1
    If relativeColumn < 1 Or relativeColumn > dataRange.Columns.Count Then relativeColumn = 1

    Dim defaultKeyword, keyword
    defaultKeyword = CStr(selectedRange.Cells(1, 1).Text)
    keyword = SafePrompt("请输入筛选关键字。留空则使用当前单元格文本。", defaultKeyword)
    keyword = Trim(keyword)
    If Len(keyword) = 0 Then keyword = defaultKeyword
    If Len(keyword) = 0 Then
        HostExcelFilterVisibleRows = "{""ok"":false,""code"":""E_EMPTY_KEYWORD"",""message"":""筛选关键字为空""}"
        Exit Function
    End If

    Dim headersJson, rangeSummaryJson, excelContextJson
    headersJson = SafeHostText("GetHeaders")
    rangeSummaryJson = SafeHostText("GetRangeSummary")
    excelContextJson = SafeHostText("GetExcelContextInfo")

    Err.Clear
    dataRange.AutoFilter relativeColumn, "*" & keyword & "*"
    If Err.Number <> 0 Then
        Dim filterErr
        filterErr = Err.Description
        Err.Clear
        HostExcelFilterVisibleRows = "{""ok"":false,""code"":""E_FILTER_FAILED"",""message"":""" & EscapeJson(filterErr) & """}"
        Exit Function
    End If

    Dim visibleCount
    visibleCount = CountVisibleDataRows(dataRange)

    Dim summary
    summary = "Excel 筛选完成" & vbCrLf & _
        "工作表: " & CStr(ws.Name) & vbCrLf & _
        "区域: " & CStr(dataRange.Address(False, False)) & vbCrLf & _
        "筛选列序号: " & CStr(relativeColumn) & vbCrLf & _
        "关键字: " & keyword & vbCrLf & _
        "可见数据行: " & CStr(visibleCount) & vbCrLf & _
        "Host.GetHeaders: " & headersJson & vbCrLf & _
        "Host.GetRangeSummary: " & rangeSummaryJson & vbCrLf & _
        "Host.GetExcelContextInfo: " & excelContextJson

    Dim copied
    copied = False
    Err.Clear
    copied = Host.WriteClipboard(summary)
    If Err.Number <> 0 Then
        copied = False
        Err.Clear
    End If
    SafeWriteLog summary

    HostExcelFilterVisibleRows = "{""ok"":true,""message"":""已完成 Excel 数据筛选并复制摘要"",""visibleRows"":" & CStr(visibleCount) & _
        ",""keyword"":""" & EscapeJson(keyword) & """,""copied"":" & JsonBool(copied) & "}"

End Function

Function CountVisibleDataRows(dataRange)
    On Error Resume Next
    Dim count, r
    count = 0
    For r = 2 To dataRange.Rows.Count
        If dataRange.Rows(r).EntireRow.Hidden = False Then count = count + 1
    Next
    CountVisibleDataRows = count
End Function

Function SafePrompt(message, defaultValue)
    On Error Resume Next
    SafePrompt = Host.Prompt(message, defaultValue)
    If Err.Number <> 0 Then
        SafePrompt = defaultValue
        Err.Clear
    End If
End Function

Function SafeHostText(methodName)
    On Error Resume Next
    Err.Clear
    Select Case methodName
        Case "GetHeaders"
            SafeHostText = Host.GetHeaders()
        Case "GetRangeSummary"
            SafeHostText = Host.GetRangeSummary()
        Case "GetExcelContextInfo"
            SafeHostText = Host.GetExcelContextInfo()
        Case Else
            SafeHostText = ""
    End Select
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

Function JsonBool(value)
    If CBool(value) Then
        JsonBool = "true"
    Else
        JsonBool = "false"
    End If
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
