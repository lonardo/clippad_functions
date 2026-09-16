' 函数名: HostExcelFormulaAuditSummary
' 描述: 扫描当前 Excel 区域公式单元格并导出审计摘要；验证 Host.GetTypeStats、GetRunPreflightPlan、ResolveOutputPlan 和剪贴板摘要
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

    Main = HostExcelFormulaAuditSummary(appObj)
End Function

Function HostExcelFormulaAuditSummary(appObj)
    On Error Resume Next

    Dim ws, selectedRange, dataRange
    Set ws = appObj.ActiveSheet
    Set selectedRange = appObj.Selection
    If Err.Number <> 0 Or TypeName(ws) = "Empty" Or TypeName(selectedRange) = "Empty" Then
        Err.Clear
        HostExcelFormulaAuditSummary = "{""ok"":false,""code"":""E_NO_SELECTION"",""message"":""请先选择 Excel 数据区域中的一个单元格""}"
        Exit Function
    End If

    Set dataRange = selectedRange.CurrentRegion
    If Err.Number <> 0 Or dataRange.Rows.Count < 1 Then
        Err.Clear
        HostExcelFormulaAuditSummary = "{""ok"":false,""code"":""E_EMPTY_RANGE"",""message"":""当前区域为空，无法审计公式""}"
        Exit Function
    End If

    Dim preflightJson, statsJson, auditText, formulaCount
    preflightJson = Host.GetRunPreflightPlan("Excel", True, True, False)
    statsJson = SafeHostText("GetTypeStats")
    formulaCount = 0
    auditText = BuildFormulaAudit(dataRange, formulaCount, statsJson, preflightJson)

    Dim outputJson, outputPath, written, summary
    outputJson = Host.ResolveOutputPlan("excel_formula_audit.txt", "avoid")
    outputPath = JsonStringValue(outputJson, "path")
    If Len(outputPath) = 0 Then outputPath = Host.ResolveTempPath("excel_formula_audit.txt")
    written = Host.WriteTextFile(outputPath, auditText, True)

    summary = "Excel 公式审计完成；工作表=" & CStr(ws.Name) & "；区域=" & dataRange.Address(False, False) & _
        "；公式单元格=" & CStr(formulaCount) & "；输出=" & outputPath
    Host.WriteClipboard summary & vbCrLf & auditText
    SafeWriteLog summary

    HostExcelFormulaAuditSummary = "{""ok"":" & JsonBool(written) & ",""message"":""" & EscapeJson(summary) & """,""formulaCount"":" & CStr(formulaCount) & _
        ",""outputPath"":""" & EscapeJson(outputPath) & """}"

End Function

Function BuildFormulaAudit(rng, ByRef formulaCount, statsJson, preflightJson)
    On Error Resume Next
    Dim text, cell, formulaText, valueText
    text = "Excel 公式审计摘要" & vbCrLf & _
        "Range: " & rng.Address(False, False) & vbCrLf & _
        "Host.GetTypeStats: " & statsJson & vbCrLf & _
        "Host.GetRunPreflightPlan: " & preflightJson & vbCrLf & vbCrLf

    formulaCount = 0
    For Each cell In rng.Cells
        formulaText = ""
        Err.Clear
        If cell.HasFormula Then formulaText = CStr(cell.Formula)
        Err.Clear
        If Len(formulaText) > 0 Then
            formulaCount = formulaCount + 1
            valueText = CStr(cell.Text)
            text = text & cell.Address(False, False) & vbTab & formulaText & vbTab & valueText & vbCrLf
        End If
    Next
    If formulaCount = 0 Then text = text & "(当前区域未检测到公式)" & vbCrLf
    BuildFormulaAudit = text
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

Function JsonStringValue(jsonText, key)
    Dim pattern, keyPos, colonPos, q1, q2
    pattern = """" & key & """"
    keyPos = InStr(1, jsonText, pattern, vbTextCompare)
    If keyPos <= 0 Then
        JsonStringValue = ""
        Exit Function
    End If
    colonPos = InStr(keyPos + Len(pattern), jsonText, ":")
    q1 = InStr(colonPos + 1, jsonText, """")
    q2 = InStr(q1 + 1, jsonText, """")
    If q1 <= 0 Or q2 <= q1 Then
        JsonStringValue = ""
    Else
        JsonStringValue = JsonUnescape(Mid(jsonText, q1 + 1, q2 - q1 - 1))
    End If
End Function

Function JsonUnescape(text)
    Dim t
    t = CStr(text)
    t = Replace(t, "\\", "\")
    t = Replace(t, "\""", """")
    t = Replace(t, "\r", vbCr)
    t = Replace(t, "\n", vbLf)
    t = Replace(t, "\t", vbTab)
    JsonUnescape = t
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
