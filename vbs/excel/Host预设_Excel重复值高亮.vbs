' 函数名: HostExcelHighlightDuplicateValues
' 描述: 高亮当前 Excel 区域中的重复值并生成重复样例摘要；适合名单去重前预检和导入质量检查
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

    Main = HostExcelHighlightDuplicateValues(appObj)
End Function

Function HostExcelHighlightDuplicateValues(appObj)
    On Error Resume Next

    Dim ws, selectedRange, dataRange
    Set ws = appObj.ActiveSheet
    Set selectedRange = appObj.Selection
    If Err.Number <> 0 Or TypeName(ws) = "Empty" Or TypeName(selectedRange) = "Empty" Then
        Err.Clear
        HostExcelHighlightDuplicateValues = "{""ok"":false,""code"":""E_NO_SELECTION"",""message"":""请先选择 Excel 数据区域中的一个单元格""}"
        Exit Function
    End If

    Set dataRange = selectedRange.CurrentRegion
    If Err.Number <> 0 Or dataRange.Rows.Count < 2 Then
        Err.Clear
        HostExcelHighlightDuplicateValues = "{""ok"":false,""code"":""E_SMALL_RANGE"",""message"":""当前区域不足两行，无法检查重复值""}"
        Exit Function
    End If

    Dim headersJson, rangeSummary, planId, previewJson
    headersJson = SafeHostText("GetHeaders")
    rangeSummary = SafeHostText("GetRangeSummary")
    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_highlight_duplicates", "{""scope"":""currentRegion""}", "office.excel.quality"
    previewJson = SafePreviewWritePlan()

    Dim keys(), counts(), keyCount, cell, key, duplicateCells, uniqueCells, sampleCount, sampleText, i
    keyCount = 0
    For Each cell In dataRange.Cells
        key = NormalizeKey(cell.Text)
        If Len(key) > 0 Then
            AddKeyCount keys, counts, keyCount, key
        End If
    Next

    duplicateCells = 0
    uniqueCells = 0
    sampleCount = 0
    sampleText = ""
    For Each cell In dataRange.Cells
        key = NormalizeKey(cell.Text)
        If Len(key) > 0 Then
            If GetKeyCount(keys, counts, keyCount, key) > 1 Then
                cell.Interior.Color = RGB(255, 199, 206)
                cell.Font.Color = RGB(156, 0, 6)
                duplicateCells = duplicateCells + 1
                If sampleCount < 8 Then
                    If InStr(1, sampleText, key & " x", vbTextCompare) = 0 Then
                        sampleText = sampleText & key & " x" & CStr(GetKeyCount(keys, counts, keyCount, key)) & "; "
                        sampleCount = sampleCount + 1
                    End If
                End If
            Else
                uniqueCells = uniqueCells + 1
            End If
        End If
    Next
    SafeCloseWritePlan planId

    Dim summary
    summary = "Excel 重复值高亮完成；重复单元格=" & CStr(duplicateCells) & "；唯一出现单元格=" & CStr(uniqueCells) & _
        "；唯一键数=" & CStr(keyCount) & "；样例=" & sampleText & vbCrLf & _
        "Host.GetHeaders: " & headersJson & vbCrLf & _
        "Host.GetRangeSummary: " & rangeSummary & vbCrLf & _
        "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostExcelHighlightDuplicateValues = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""duplicateCells"":" & CStr(duplicateCells) & _
        ",""uniqueCells"":" & CStr(uniqueCells) & ",""uniqueKeys"":" & CStr(keyCount) & ",""planId"":""" & EscapeJson(planId) & """}"
End Function

Sub AddKeyCount(ByRef keys, ByRef counts, ByRef keyCount, keyText)
    Dim i
    For i = 0 To keyCount - 1
        If StrComp(keys(i), keyText, vbTextCompare) = 0 Then
            counts(i) = counts(i) + 1
            Exit Sub
        End If
    Next
    ReDim Preserve keys(keyCount)
    ReDim Preserve counts(keyCount)
    keys(keyCount) = keyText
    counts(keyCount) = 1
    keyCount = keyCount + 1
End Sub

Function GetKeyCount(ByRef keys, ByRef counts, keyCount, keyText)
    Dim i
    For i = 0 To keyCount - 1
        If StrComp(keys(i), keyText, vbTextCompare) = 0 Then
            GetKeyCount = counts(i)
            Exit Function
        End If
    Next
    GetKeyCount = 0
End Function

Function NormalizeKey(value)
    Dim text
    text = CStr(value)
    text = Replace(text, ChrW(160), " ")
    text = Replace(text, ChrW(12288), " ")
    text = Trim(text)
    NormalizeKey = LCase(text)
End Function

Function SafeHostText(methodName)
    On Error Resume Next
    If methodName = "GetHeaders" Then
        SafeHostText = Host.GetHeaders()
    ElseIf methodName = "GetRangeSummary" Then
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