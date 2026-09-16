' 函数名: HostExcelNormalizeDateText
' 描述: 将当前 Excel 区域中常见文本日期转为真正日期并统一显示格式；对标导入表清洗插件
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
    Main = HostExcelNormalizeDateText(appObj)
End Function

Function HostExcelNormalizeDateText(appObj)
    On Error Resume Next
    Dim ws, selectedRange, dataRange
    Set ws = appObj.ActiveSheet
    Set selectedRange = appObj.Selection
    If Err.Number <> 0 Or TypeName(ws) = "Empty" Or TypeName(selectedRange) = "Empty" Then
        Err.Clear
        HostExcelNormalizeDateText = "{""ok"":false,""code"":""E_NO_SELECTION"",""message"":""请先选择 Excel 数据区域中的一个单元格""}"
        Exit Function
    End If
    Set dataRange = selectedRange.CurrentRegion
    If Err.Number <> 0 Or dataRange.Rows.Count < 1 Then
        Err.Clear
        HostExcelNormalizeDateText = "{""ok"":false,""code"":""E_EMPTY_RANGE"",""message"":""当前区域为空""}"
        Exit Function
    End If

    Dim formatText
    formatText = SafePrompt("请输入目标日期格式，默认 yyyy-mm-dd", "yyyy-mm-dd")
    If Len(Trim(CStr(formatText))) = 0 Then formatText = "yyyy-mm-dd"

    Dim typeStats, planId, previewJson, cell, changed, scanned, textValue, dateValue
    typeStats = SafeHostText("GetTypeStats")
    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_normalize_date_text", "{""format"":""" & EscapeJson(formatText) & """}", "office.excel.cell.cleanup"
    previewJson = SafePreviewWritePlan()

    changed = 0
    scanned = 0
    For Each cell In dataRange.Cells
        scanned = scanned + 1
        If Not IsError(cell.Value) Then
            textValue = Trim(CStr(cell.Text))
            If LooksLikeDateText(textValue) Then
                dateValue = ParseLooseDate(textValue)
                If IsDate(dateValue) Then
                    cell.Value = CDate(dateValue)
                    cell.NumberFormat = formatText
                    If Err.Number = 0 Then changed = changed + 1
                    Err.Clear
                End If
            End If
        End If
    Next
    SafeCloseWritePlan planId

    Dim summary
    summary = "Excel 日期规范化完成；扫描=" & CStr(scanned) & "；转换=" & CStr(changed) & _
        "；格式=" & formatText & "；区域=" & dataRange.Address(False, False) & vbCrLf & _
        "Host.GetTypeStats: " & typeStats & vbCrLf & "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostExcelNormalizeDateText = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""scanned"":" & CStr(scanned) & _
        ",""changed"":" & CStr(changed) & ",""planId"":""" & EscapeJson(planId) & """}"
End Function


Function LooksLikeDateText(value)
    Dim text
    text = Trim(CStr(value))
    LooksLikeDateText = False
    If Len(text) < 6 Or Len(text) > 20 Then Exit Function
    If InStr(text, "年") > 0 Or InStr(text, "月") > 0 Or InStr(text, "日") > 0 Then
        LooksLikeDateText = True
        Exit Function
    End If
    If InStr(text, "-") > 0 Or InStr(text, "/") > 0 Or InStr(text, ".") > 0 Then
        LooksLikeDateText = True
        Exit Function
    End If
End Function
Function ParseLooseDate(value)
    On Error Resume Next
    Dim text
    text = CStr(value)
    text = Replace(text, ".", "-")
    text = Replace(text, "/", "-")
    text = Replace(text, "年", "-")
    text = Replace(text, "月", "-")
    text = Replace(text, "日", "")
    text = Replace(text, " ", "")
    If IsDate(text) Then
        ParseLooseDate = CDate(text)
    Else
        ParseLooseDate = ""
    End If
End Function

Function SafeHostText(methodName)
    On Error Resume Next
    If methodName = "GetTypeStats" Then
        SafeHostText = Host.GetTypeStats()
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