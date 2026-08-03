' 函数名: HostExcelBlankErrorReport
' 描述: 统计当前 Excel 区域中的空白、错误值和公式单元格，生成检查报告并复制摘要；适合交付前数据体检
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

    Main = HostExcelBlankErrorReport(appObj)
End Function

Function HostExcelBlankErrorReport(appObj)
    On Error Resume Next

    Dim ws, selectedRange, dataRange
    Set ws = appObj.ActiveSheet
    Set selectedRange = appObj.Selection
    If Err.Number <> 0 Or TypeName(ws) = "Empty" Or TypeName(selectedRange) = "Empty" Then
        Err.Clear
        HostExcelBlankErrorReport = "{""ok"":false,""code"":""E_NO_SELECTION"",""message"":""请先选择 Excel 数据区域中的一个单元格""}"
        Exit Function
    End If

    Set dataRange = selectedRange.CurrentRegion
    If Err.Number <> 0 Or dataRange.Rows.Count < 1 Then
        Err.Clear
        HostExcelBlankErrorReport = "{""ok"":false,""code"":""E_EMPTY_RANGE"",""message"":""当前区域为空，无法生成报告""}"
        Exit Function
    End If

    Dim headersJson, hasBlankJson, hasFormulaJson
    headersJson = SafeHostText("GetHeaders")
    hasBlankJson = SafeHostText("HasBlank")
    hasFormulaJson = SafeHostText("HasFormula")

    Dim outSheet
    Set outSheet = appObj.Worksheets.Add
    outSheet.Name = SafeSheetName(appObj, "Host单元格体检")
    outSheet.Range("A1").Value = "类型"
    outSheet.Range("B1").Value = "地址"
    outSheet.Range("C1").Value = "显示值"
    outSheet.Range("D1").Value = "公式"

    Dim blankCount, errorCount, formulaCount, rowOut
    rowOut = 2
    ScanCells dataRange, outSheet, rowOut, blankCount, errorCount, formulaCount
    outSheet.Rows(1).Font.Bold = True
    outSheet.Columns("A:D").AutoFit

    Dim summary
    summary = "Excel 单元格体检完成；来源=" & ws.Name & "!" & dataRange.Address(False, False) & _
        "；空白=" & CStr(blankCount) & "；错误=" & CStr(errorCount) & "；公式=" & CStr(formulaCount) & vbCrLf & _
        "Host.GetHeaders: " & headersJson & vbCrLf & _
        "Host.HasBlank: " & hasBlankJson & vbCrLf & _
        "Host.HasFormula: " & hasFormulaJson
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostExcelBlankErrorReport = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""blankCount"":" & CStr(blankCount) & _
        ",""errorCount"":" & CStr(errorCount) & ",""formulaCount"":" & CStr(formulaCount) & _
        ",""outputSheet"":""" & EscapeJson(outSheet.Name) & """}"

End Function

Sub ScanCells(dataRange, outSheet, ByRef rowOut, ByRef blankCount, ByRef errorCount, ByRef formulaCount)
    On Error Resume Next
    Dim cell
    blankCount = 0
    errorCount = 0
    formulaCount = 0
    For Each cell In dataRange.Cells
        If Len(Trim(CStr(cell.Text))) = 0 Then
            blankCount = blankCount + 1
            WriteIssue outSheet, rowOut, "空白", cell.Address(False, False), "", ""
        End If
        If IsError(cell.Value) Then
            errorCount = errorCount + 1
            WriteIssue outSheet, rowOut, "错误", cell.Address(False, False), CStr(cell.Text), ""
        End If
        If cell.HasFormula Then
            formulaCount = formulaCount + 1
            WriteIssue outSheet, rowOut, "公式", cell.Address(False, False), CStr(cell.Text), CStr(cell.Formula)
        End If
        If Err.Number <> 0 Then Err.Clear
    Next
End Sub

Sub WriteIssue(outSheet, ByRef rowOut, issueType, addressText, displayText, formulaText)
    On Error Resume Next
    outSheet.Cells(rowOut, 1).Value = issueType
    outSheet.Cells(rowOut, 2).Value = addressText
    outSheet.Cells(rowOut, 3).Value = displayText
    outSheet.Cells(rowOut, 4).Value = formulaText
    rowOut = rowOut + 1
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
    Select Case methodName
        Case "GetHeaders"
            SafeHostText = Host.GetHeaders()
        Case "HasBlank"
            SafeHostText = CStr(Host.HasBlank())
        Case "HasFormula"
            SafeHostText = CStr(Host.HasFormula())
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
