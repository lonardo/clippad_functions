' 函数名: HostExcelFormulaValueSnapshot
' 描述: 将当前 Excel 区域复制到新工作表并仅保留显示值，形成公式转数值快照；适合报表发送前固化结果
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

    Main = HostExcelFormulaValueSnapshot(appObj)
End Function

Function HostExcelFormulaValueSnapshot(appObj)
    On Error Resume Next

    Dim ws, selectedRange, dataRange
    Set ws = appObj.ActiveSheet
    Set selectedRange = appObj.Selection
    If Err.Number <> 0 Or TypeName(ws) = "Empty" Or TypeName(selectedRange) = "Empty" Then
        Err.Clear
        HostExcelFormulaValueSnapshot = "{""ok"":false,""code"":""E_NO_SELECTION"",""message"":""请先选择 Excel 数据区域中的一个单元格""}"
        Exit Function
    End If

    Set dataRange = selectedRange.CurrentRegion
    If Err.Number <> 0 Or dataRange.Rows.Count < 1 Then
        Err.Clear
        HostExcelFormulaValueSnapshot = "{""ok"":false,""code"":""E_EMPTY_RANGE"",""message"":""当前区域为空，无法生成快照""}"
        Exit Function
    End If

    Dim formulaCount, planId, previewJson, rangeSummary
    formulaCount = CountFormulas(dataRange)
    rangeSummary = SafeHostText("GetRangeSummary")
    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_formula_value_snapshot", "{""scope"":""currentRegion""}", "office.excel.formulaSnapshot"
    previewJson = SafePreviewWritePlan()

    Dim outSheet
    Set outSheet = appObj.Worksheets.Add
    outSheet.Name = SafeSheetName(appObj, "Host数值快照")
    CopyValuesAndFormats dataRange, outSheet.Range("A1")
    SafeCloseWritePlan planId

    Dim summary
    summary = "Excel 公式转数值快照完成；来源=" & ws.Name & "!" & dataRange.Address(False, False) & _
        "；公式单元格=" & CStr(formulaCount) & "；输出表=" & outSheet.Name & vbCrLf & _
        "Host.GetRangeSummary: " & rangeSummary & vbCrLf & _
        "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostExcelFormulaValueSnapshot = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""formulaCount"":" & CStr(formulaCount) & _
        ",""rows"":" & CStr(dataRange.Rows.Count) & ",""columns"":" & CStr(dataRange.Columns.Count) & _
        ",""outputSheet"":""" & EscapeJson(outSheet.Name) & """}"

End Function

Function CountFormulas(dataRange)
    On Error Resume Next
    Dim cell, count
    count = 0
    For Each cell In dataRange.Cells
        If cell.HasFormula Then count = count + 1
        If Err.Number <> 0 Then Err.Clear
    Next
    CountFormulas = count
End Function

Sub CopyValuesAndFormats(sourceRange, targetCell)
    On Error Resume Next
    Dim outRange, r, c
    Set outRange = targetCell.Resize(sourceRange.Rows.Count, sourceRange.Columns.Count)
    For r = 1 To sourceRange.Rows.Count
        For c = 1 To sourceRange.Columns.Count
            outRange.Cells(r, c).Value = sourceRange.Cells(r, c).Text
        Next
    Next
    sourceRange.Copy
    outRange.PasteSpecial -4122
    outRange.Columns.AutoFit
    Err.Clear
End Sub

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
