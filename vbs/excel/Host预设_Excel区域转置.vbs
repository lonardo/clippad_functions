' 函数名: HostExcelTransposeCurrentRegion
' 描述: 将当前 Excel 区域转置输出到新工作表；对标表格工具插件中的一键转置
' 适用应用: Excel
' 搜索范围: 当前范围
' 搜索对象: 无

Option Explicit

Const xlPasteAll = -4104
Const xlPasteSpecialOperationNone = -4142

Function Main()
    On Error Resume Next
    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = "{""ok"":false,""code"":""E_NO_EXCEL_APP"",""message"":""未取得 Excel 应用，请在 Excel 中运行该预设""}"
        Exit Function
    End If
    Main = HostExcelTransposeCurrentRegion(appObj)
End Function

Function HostExcelTransposeCurrentRegion(appObj)
    On Error Resume Next
    Dim ws, selectedRange, dataRange
    Set ws = appObj.ActiveSheet
    Set selectedRange = appObj.Selection
    If Err.Number <> 0 Or TypeName(ws) = "Empty" Or TypeName(selectedRange) = "Empty" Then
        Err.Clear
        HostExcelTransposeCurrentRegion = "{""ok"":false,""code"":""E_NO_SELECTION"",""message"":""请先选择 Excel 数据区域中的一个单元格""}"
        Exit Function
    End If
    Set dataRange = selectedRange.CurrentRegion
    If Err.Number <> 0 Or dataRange.Rows.Count < 1 Then
        Err.Clear
        HostExcelTransposeCurrentRegion = "{""ok"":false,""code"":""E_EMPTY_RANGE"",""message"":""当前区域为空""}"
        Exit Function
    End If

    Dim rangeSummary, planId, previewJson, outSheet
    rangeSummary = SafeHostText("GetRangeSummary")
    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_transpose_current_region", "{""scope"":""currentRegion""}", "office.excel.transform"
    previewJson = SafePreviewWritePlan()

    Set outSheet = appObj.Worksheets.Add
    outSheet.Name = SafeSheetName(appObj, "Host转置结果")
    dataRange.Copy
    outSheet.Range("A1").PasteSpecial xlPasteAll, xlPasteSpecialOperationNone, False, True
    appObj.CutCopyMode = False
    outSheet.Columns.AutoFit
    SafeCloseWritePlan planId

    Dim summary
    summary = "Excel 区域转置完成；来源=" & ws.Name & "!" & dataRange.Address(False, False) & _
        "；输出=" & outSheet.Name & vbCrLf & _
        "Host.GetRangeSummary: " & rangeSummary & vbCrLf & "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostExcelTransposeCurrentRegion = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""sourceRange"":""" & EscapeJson(dataRange.Address(False, False)) & _
        """,""outputSheet"":""" & EscapeJson(outSheet.Name) & """,""planId"":""" & EscapeJson(planId) & """}"
End Function

Function SafeSheetName(appObj, baseName)
    On Error Resume Next
    Dim candidate, index, exists, ws
    index = 1
    Do
        If index = 1 Then
            candidate = baseName
        Else
            candidate = baseName & CStr(index)
        End If
        exists = False
        For Each ws In appObj.Worksheets
            If StrComp(ws.Name, candidate, vbTextCompare) = 0 Then exists = True
        Next
        If Not exists Then
            SafeSheetName = candidate
            Exit Function
        End If
        index = index + 1
    Loop
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