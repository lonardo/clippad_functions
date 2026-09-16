' 函数名: HostExcelNormalizeCase
' 描述: 将当前 Excel 区域文本按选择转换为大写、小写或首字母大写；对标 Excel 清洗插件常见大小写规范
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
    Main = HostExcelNormalizeCase(appObj)
End Function

Function HostExcelNormalizeCase(appObj)
    On Error Resume Next
    Dim ws, selectedRange, dataRange
    Set ws = appObj.ActiveSheet
    Set selectedRange = appObj.Selection
    If Err.Number <> 0 Or TypeName(ws) = "Empty" Or TypeName(selectedRange) = "Empty" Then
        Err.Clear
        HostExcelNormalizeCase = "{""ok"":false,""code"":""E_NO_SELECTION"",""message"":""请先选择 Excel 数据区域中的一个单元格""}"
        Exit Function
    End If
    Set dataRange = selectedRange.CurrentRegion
    If Err.Number <> 0 Or dataRange.Rows.Count < 1 Then
        Err.Clear
        HostExcelNormalizeCase = "{""ok"":false,""code"":""E_EMPTY_RANGE"",""message"":""当前区域为空""}"
        Exit Function
    End If

    Dim modeText, mode
    modeText = SafePrompt("请输入大小写模式：upper / lower / proper，默认 upper", "upper")
    mode = LCase(Trim(CStr(modeText)))
    If mode <> "upper" And mode <> "lower" And mode <> "proper" Then mode = "upper"

    Dim typeStats, planId, previewJson, cell, changed, scanned, textValue, newValue
    typeStats = SafeHostText("GetTypeStats")
    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_normalize_case", "{""mode"":""" & mode & """}", "office.excel.cell.cleanup"
    previewJson = SafePreviewWritePlan()

    changed = 0
    scanned = 0
    For Each cell In dataRange.Cells
        scanned = scanned + 1
        If Not IsError(cell.Value) Then
            If VarType(cell.Value) = 8 Then
                textValue = CStr(cell.Value)
                If mode = "upper" Then
                    newValue = UCase(textValue)
                ElseIf mode = "lower" Then
                    newValue = LCase(textValue)
                Else
                    newValue = StrConv(textValue, 3)
                End If
                If newValue <> textValue Then
                    cell.Value = newValue
                    If Err.Number = 0 Then changed = changed + 1
                    Err.Clear
                End If
            End If
        End If
    Next
    SafeCloseWritePlan planId

    Dim summary
    summary = "Excel 大小写规范完成；模式=" & mode & "；扫描=" & CStr(scanned) & "；修改=" & CStr(changed) & _
        "；区域=" & dataRange.Address(False, False) & vbCrLf & _
        "Host.GetTypeStats: " & typeStats & vbCrLf & "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostExcelNormalizeCase = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""mode"":""" & mode & """,""scanned"":" & CStr(scanned) & _
        ",""changed"":" & CStr(changed) & ",""planId"":""" & EscapeJson(planId) & """}"
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