' 函数名: HostExcelSummaryFormulaTemplate
' 描述: 根据当前 Excel 数据区域表头生成常用汇总公式模板，包括计数、求和、平均、最大和最小；适合快速搭建公式检查页
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

    Main = HostExcelSummaryFormulaTemplate(appObj)
End Function

Function HostExcelSummaryFormulaTemplate(appObj)
    On Error Resume Next

    Dim ws, selectedRange, dataRange
    Set ws = appObj.ActiveSheet
    Set selectedRange = appObj.Selection
    If Err.Number <> 0 Or TypeName(ws) = "Empty" Or TypeName(selectedRange) = "Empty" Then
        Err.Clear
        HostExcelSummaryFormulaTemplate = "{""ok"":false,""code"":""E_NO_SELECTION"",""message"":""请先选择 Excel 数据区域中的一个单元格""}"
        Exit Function
    End If

    Set dataRange = selectedRange.CurrentRegion
    If Err.Number <> 0 Or dataRange.Rows.Count < 2 Then
        Err.Clear
        HostExcelSummaryFormulaTemplate = "{""ok"":false,""code"":""E_SMALL_RANGE"",""message"":""当前区域至少需要标题行和一行数据""}"
        Exit Function
    End If

    Dim headersJson, budgetJson, planId, previewJson
    headersJson = SafeHostText("GetHeaders")
    budgetJson = SafeRunBudget(dataRange.Columns.Count)
    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_summary_formula_template", "{""scope"":""currentRegion""}", "office.excel.formulaTemplate"
    previewJson = SafePreviewWritePlan()

    Dim outSheet
    Set outSheet = appObj.Worksheets.Add
    outSheet.Name = SafeSheetName(appObj, "Host公式模板")
    BuildFormulaTemplate ws, dataRange, outSheet
    SafeCloseWritePlan planId

    Dim summary
    summary = "Excel 常用汇总公式模板生成完成；来源=" & ws.Name & "!" & dataRange.Address(False, False) & _
        "；字段数=" & CStr(dataRange.Columns.Count) & "；输出表=" & outSheet.Name & vbCrLf & _
        "Host.GetHeaders: " & headersJson & vbCrLf & _
        "Host.GetRunBudgetPlan: " & budgetJson & vbCrLf & _
        "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostExcelSummaryFormulaTemplate = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""columns"":" & CStr(dataRange.Columns.Count) & _
        ",""outputSheet"":""" & EscapeJson(outSheet.Name) & """}"

End Function

Sub BuildFormulaTemplate(sourceSheet, dataRange, outSheet)
    On Error Resume Next
    Dim c, sourceAddress, headerText, dataAddress, sheetName
    outSheet.Range("A1").Value = "字段"
    outSheet.Range("B1").Value = "非空计数"
    outSheet.Range("C1").Value = "求和"
    outSheet.Range("D1").Value = "平均"
    outSheet.Range("E1").Value = "最大"
    outSheet.Range("F1").Value = "最小"
    outSheet.Rows(1).Font.Bold = True
    sheetName = "'" & Replace(sourceSheet.Name, "'", "''") & "'!"

    For c = 1 To dataRange.Columns.Count
        headerText = CStr(dataRange.Cells(1, c).Text)
        If Len(Trim(headerText)) = 0 Then headerText = "字段" & CStr(c)
        sourceAddress = dataRange.Offset(1, c - 1).Resize(dataRange.Rows.Count - 1, 1).Address(False, False)
        dataAddress = sheetName & sourceAddress
        outSheet.Cells(c + 1, 1).Value = headerText
        outSheet.Cells(c + 1, 2).Formula = "=COUNTA(" & dataAddress & ")"
        outSheet.Cells(c + 1, 3).Formula = "=SUM(" & dataAddress & ")"
        outSheet.Cells(c + 1, 4).Formula = "=AVERAGE(" & dataAddress & ")"
        outSheet.Cells(c + 1, 5).Formula = "=MAX(" & dataAddress & ")"
        outSheet.Cells(c + 1, 6).Formula = "=MIN(" & dataAddress & ")"
    Next
    outSheet.Columns("A:F").AutoFit
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
    If methodName = "GetHeaders" Then
        SafeHostText = Host.GetHeaders()
    Else
        SafeHostText = ""
    End If
    If Err.Number <> 0 Then
        SafeHostText = ""
        Err.Clear
    End If
End Function

Function SafeRunBudget(columnCount)
    On Error Resume Next
    SafeRunBudget = Host.GetRunBudgetPlan("excel_summary_formula_template", columnCount, 200, 10, 120)
    If Err.Number <> 0 Then
        SafeRunBudget = ""
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
