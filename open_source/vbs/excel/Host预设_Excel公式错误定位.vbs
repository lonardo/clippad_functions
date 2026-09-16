' 函数名: HostExcelLocateFormulaIssues
' 描述: 扫描当前 Excel 区域中的公式错误、外部链接和易波动函数并高亮定位；适合公式表交付前审计
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

    Main = HostExcelLocateFormulaIssues(appObj)
End Function

Function HostExcelLocateFormulaIssues(appObj)
    On Error Resume Next

    Dim ws, selectedRange, dataRange
    Set ws = appObj.ActiveSheet
    Set selectedRange = appObj.Selection
    If Err.Number <> 0 Or TypeName(ws) = "Empty" Or TypeName(selectedRange) = "Empty" Then
        Err.Clear
        HostExcelLocateFormulaIssues = "{""ok"":false,""code"":""E_NO_SELECTION"",""message"":""请先选择 Excel 数据区域中的一个单元格""}"
        Exit Function
    End If

    Set dataRange = selectedRange.CurrentRegion
    If Err.Number <> 0 Or dataRange.Rows.Count < 1 Then
        Err.Clear
        HostExcelLocateFormulaIssues = "{""ok"":false,""code"":""E_EMPTY_RANGE"",""message"":""当前区域为空，无法审计公式""}"
        Exit Function
    End If

    Dim typeStats, formulaIssueSummary, preflightJson, planId, previewJson
    typeStats = SafeHostText("GetTypeStats")
    formulaIssueSummary = SafeHostText("GetExcelFormulaIssueSummary")
    preflightJson = SafePreflight()
    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_locate_formula_issues", "{""scope"":""currentRegion""}", "office.excel.formulaAudit"
    previewJson = SafePreviewWritePlan()

    Dim formulaCount, errorCount, externalCount, volatileCount, report
    report = BuildFormulaIssueReport(dataRange, formulaCount, errorCount, externalCount, volatileCount)
    SafeCloseWritePlan planId

    Dim summary
    summary = "Excel 公式问题定位完成；工作表=" & ws.Name & "；区域=" & dataRange.Address(False, False) & _
        "；公式=" & CStr(formulaCount) & "；错误=" & CStr(errorCount) & _
        "；外部链接=" & CStr(externalCount) & "；易波动函数=" & CStr(volatileCount) & vbCrLf & _
        report & vbCrLf & _
        "Host.GetTypeStats: " & typeStats & vbCrLf & _
        "Host.GetExcelFormulaIssueSummary: " & formulaIssueSummary & vbCrLf & _
        "Host.GetRunPreflightPlan: " & preflightJson & vbCrLf & _
        "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostExcelLocateFormulaIssues = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""formulaCount"":" & CStr(formulaCount) & _
        ",""errorCount"":" & CStr(errorCount) & ",""externalLinkCount"":" & CStr(externalCount) & _
        ",""volatileCount"":" & CStr(volatileCount) & "}"

End Function

Function BuildFormulaIssueReport(dataRange, ByRef formulaCount, ByRef errorCount, ByRef externalCount, ByRef volatileCount)
    On Error Resume Next
    Dim cell, formulaText, issueText, report, sampleCount
    formulaCount = 0
    errorCount = 0
    externalCount = 0
    volatileCount = 0
    sampleCount = 0
    report = "问题样例:"

    For Each cell In dataRange.Cells
        formulaText = ""
        If cell.HasFormula Then formulaText = CStr(cell.Formula)
        If Len(formulaText) > 0 Then
            formulaCount = formulaCount + 1
            issueText = ""
            If IsError(cell.Value) Then
                errorCount = errorCount + 1
                issueText = AppendIssue(issueText, "公式错误")
                cell.Interior.Color = RGB(255, 199, 206)
                cell.Font.Color = RGB(156, 0, 6)
            End If
            If FormulaHasExternalLink(formulaText) Then
                externalCount = externalCount + 1
                issueText = AppendIssue(issueText, "外部链接")
                If Not IsError(cell.Value) Then cell.Interior.Color = RGB(255, 235, 156)
            End If
            If FormulaHasVolatileFunction(formulaText) Then
                volatileCount = volatileCount + 1
                issueText = AppendIssue(issueText, "易波动函数")
                If Not IsError(cell.Value) Then cell.Interior.Color = RGB(218, 238, 243)
            End If
            If Len(issueText) > 0 Then
                sampleCount = sampleCount + 1
                If sampleCount <= 20 Then
                    report = report & vbCrLf & cell.Address(False, False) & "｜" & issueText & "｜" & formulaText
                End If
            End If
        End If
        If Err.Number <> 0 Then Err.Clear
    Next

    If sampleCount = 0 Then
        report = report & " 未发现公式错误、外部链接或易波动函数"
    ElseIf sampleCount > 20 Then
        report = report & vbCrLf & "另有 " & CStr(sampleCount - 20) & " 条问题未列出"
    End If
    BuildFormulaIssueReport = report
End Function

Function AppendIssue(currentText, itemText)
    If Len(currentText) = 0 Then
        AppendIssue = itemText
    Else
        AppendIssue = currentText & "," & itemText
    End If
End Function

Function FormulaHasExternalLink(formulaText)
    Dim text
    text = LCase(CStr(formulaText))
    FormulaHasExternalLink = (InStr(text, "[") > 0 Or InStr(text, "http://") > 0 Or InStr(text, "https://") > 0)
End Function

Function FormulaHasVolatileFunction(formulaText)
    Dim text
    text = UCase(CStr(formulaText))
    FormulaHasVolatileFunction = (InStr(text, "NOW(") > 0 Or InStr(text, "TODAY(") > 0 Or _
        InStr(text, "RAND(") > 0 Or InStr(text, "RANDBETWEEN(") > 0 Or _
        InStr(text, "OFFSET(") > 0 Or InStr(text, "INDIRECT(") > 0)
End Function

Function SafeHostText(methodName)
    On Error Resume Next
    If methodName = "GetTypeStats" Then
        SafeHostText = Host.GetTypeStats()
    ElseIf methodName = "GetExcelFormulaIssueSummary" Then
        SafeHostText = Host.GetExcelFormulaIssueSummary()
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
