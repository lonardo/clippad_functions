' 函数名: HostExcelHighlightOutliers
' 描述: 扫描当前 Excel 区域中的数字单元格并高亮明显偏离均值的异常值，同时复制异常摘要；适合质检、财务和运营表排查
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

    Main = HostExcelHighlightOutliers(appObj)
End Function

Function HostExcelHighlightOutliers(appObj)
    On Error Resume Next

    Dim ws, selectedRange, dataRange
    Set ws = appObj.ActiveSheet
    Set selectedRange = appObj.Selection
    If Err.Number <> 0 Or TypeName(ws) = "Empty" Or TypeName(selectedRange) = "Empty" Then
        Err.Clear
        HostExcelHighlightOutliers = "{""ok"":false,""code"":""E_NO_SELECTION"",""message"":""请先选择 Excel 数据区域中的一个单元格""}"
        Exit Function
    End If

    Set dataRange = selectedRange.CurrentRegion
    If Err.Number <> 0 Or dataRange.Rows.Count < 2 Then
        Err.Clear
        HostExcelHighlightOutliers = "{""ok"":false,""code"":""E_SMALL_RANGE"",""message"":""当前区域不足两行，无法检测异常值""}"
        Exit Function
    End If

    Dim thresholdText, threshold
    thresholdText = SafePrompt("请输入异常阈值倍数，默认 2 表示偏离均值超过 2 个标准差", "2")
    threshold = CDblSafe(thresholdText, 2)
    If threshold <= 0 Then threshold = 2

    Dim typeStats, planId, previewJson
    typeStats = SafeHostText("GetTypeStats")
    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_highlight_outliers", "{""scope"":""currentRegion"",""threshold"":" & NumberText(threshold) & "}", "office.excel.analyze"
    previewJson = SafePreviewWritePlan()

    Dim numericCount, meanValue, stdevValue
    numericCount = CountNumericCells(dataRange)
    If numericCount < 2 Then
        SafeCloseWritePlan planId
        HostExcelHighlightOutliers = "{""ok"":false,""code"":""E_NOT_ENOUGH_NUMBERS"",""message"":""当前区域数字单元格不足两个，无法检测异常值""}"
        Exit Function
    End If

    meanValue = AverageNumericCells(dataRange)
    stdevValue = StdevNumericCells(dataRange, meanValue, numericCount)

    Dim highlighted, report
    highlighted = HighlightOutlierCells(dataRange, meanValue, stdevValue, threshold, report)
    SafeCloseWritePlan planId

    Dim summary
    summary = "Excel 异常值检测完成；工作表=" & ws.Name & "；区域=" & dataRange.Address(False, False) & _
        "；数字单元格=" & CStr(numericCount) & "；高亮异常=" & CStr(highlighted) & _
        "；均值=" & FormatNumber(meanValue, 2) & "；标准差=" & FormatNumber(stdevValue, 2) & vbCrLf & _
        report & vbCrLf & _
        "Host.GetTypeStats: " & typeStats & vbCrLf & _
        "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostExcelHighlightOutliers = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""numericCount"":" & CStr(numericCount) & _
        ",""highlighted"":" & CStr(highlighted) & ",""mean"":" & NumberText(meanValue) & ",""stdev"":" & NumberText(stdevValue) & "}"

End Function

Function CountNumericCells(dataRange)
    On Error Resume Next
    Dim cell, count
    count = 0
    For Each cell In dataRange.Cells
        If IsNumeric(cell.Value) And Len(CStr(cell.Value)) > 0 Then count = count + 1
    Next
    CountNumericCells = count
End Function

Function AverageNumericCells(dataRange)
    On Error Resume Next
    Dim cell, sumValue, count
    sumValue = 0
    count = 0
    For Each cell In dataRange.Cells
        If IsNumeric(cell.Value) And Len(CStr(cell.Value)) > 0 Then
            sumValue = sumValue + CDbl(cell.Value)
            count = count + 1
        End If
    Next
    If count = 0 Then
        AverageNumericCells = 0
    Else
        AverageNumericCells = sumValue / count
    End If
End Function

Function StdevNumericCells(dataRange, meanValue, count)
    On Error Resume Next
    Dim cell, sumSquares, delta
    sumSquares = 0
    For Each cell In dataRange.Cells
        If IsNumeric(cell.Value) And Len(CStr(cell.Value)) > 0 Then
            delta = CDbl(cell.Value) - meanValue
            sumSquares = sumSquares + delta * delta
        End If
    Next
    If count <= 1 Then
        StdevNumericCells = 0
    Else
        StdevNumericCells = Sqr(sumSquares / (count - 1))
    End If
End Function

Function HighlightOutlierCells(dataRange, meanValue, stdevValue, threshold, ByRef report)
    On Error Resume Next
    Dim cell, count, value, score, maxItems
    count = 0
    maxItems = 12
    report = "异常样例:"
    If stdevValue <= 0 Then
        HighlightOutlierCells = 0
        report = report & " 标准差为 0，未发现异常"
        Exit Function
    End If

    For Each cell In dataRange.Cells
        If IsNumeric(cell.Value) And Len(CStr(cell.Value)) > 0 Then
            value = CDbl(cell.Value)
            score = Abs(value - meanValue) / stdevValue
            If score >= threshold Then
                cell.Interior.Color = RGB(255, 242, 204)
                cell.Font.Color = RGB(156, 87, 0)
                count = count + 1
                If count <= maxItems Then
                    report = report & vbCrLf & cell.Address(False, False) & "=" & CStr(cell.Value) & "，偏离=" & FormatNumber(score, 2)
                End If
            End If
        End If
    Next
    If count = 0 Then report = report & " 未发现超过阈值的数字"
    If count > maxItems Then report = report & vbCrLf & "另有 " & CStr(count - maxItems) & " 个异常未列出"
    HighlightOutlierCells = count
End Function

Function CDblSafe(value, fallback)
    On Error Resume Next
    CDblSafe = CDbl(value)
    If Err.Number <> 0 Then
        CDblSafe = fallback
        Err.Clear
    End If
End Function

Function NumberText(value)
    NumberText = Replace(CStr(value), ",", "")
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
