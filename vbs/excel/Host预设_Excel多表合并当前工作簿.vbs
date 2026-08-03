' 函数名: HostExcelMergeWorksheets
' 描述: 合并当前工作簿中多个工作表的 UsedRange 到一个汇总表，并保留来源工作表名称；适合月报、分店和分部门明细汇总
' 适用应用: Excel
' 搜索范围: 全文
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

    Main = HostExcelMergeWorksheets(appObj)
End Function

Function HostExcelMergeWorksheets(appObj)
    On Error Resume Next

    Dim wb
    Set wb = appObj.ActiveWorkbook
    If Err.Number <> 0 Or TypeName(wb) = "Empty" Or TypeName(wb) = "Nothing" Then
        Err.Clear
        HostExcelMergeWorksheets = "{""ok"":false,""code"":""E_NO_WORKBOOK"",""message"":""当前没有活动工作簿""}"
        Exit Function
    End If
    If wb.Worksheets.Count < 2 Then
        HostExcelMergeWorksheets = "{""ok"":false,""code"":""E_SINGLE_SHEET"",""message"":""当前工作簿少于两个工作表，无需合并""}"
        Exit Function
    End If

    Dim workbookName, budgetJson, planId, previewJson
    workbookName = SafeHostText("GetWorkbookName")
    budgetJson = SafeRunBudget(wb.Worksheets.Count)
    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_merge_worksheets", "{""scope"":""activeWorkbook""}", "office.excel.merge"
    previewJson = SafePreviewWritePlan()

    Dim outSheet
    Set outSheet = wb.Worksheets.Add
    outSheet.Name = SafeSheetName(wb.Application, "Host多表合并")
    outSheet.Cells(1, 1).Value = "来源工作表"

    Dim rowOut, mergedSheets, mergedRows
    rowOut = 2
    mergedSheets = 0
    mergedRows = 0
    MergeWorkbookSheets wb, outSheet, rowOut, mergedSheets, mergedRows
    outSheet.Rows(1).Font.Bold = True
    outSheet.Columns.AutoFit
    SafeCloseWritePlan planId

    Dim summary
    summary = "Excel 多表合并完成；工作簿=" & workbookName & "；合并工作表=" & CStr(mergedSheets) & _
        "；合并数据行=" & CStr(mergedRows) & "；输出表=" & outSheet.Name & vbCrLf & _
        "Host.GetRunBudgetPlan: " & budgetJson & vbCrLf & _
        "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostExcelMergeWorksheets = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""mergedSheets"":" & CStr(mergedSheets) & _
        ",""mergedRows"":" & CStr(mergedRows) & ",""outputSheet"":""" & EscapeJson(outSheet.Name) & """}"

End Function

Sub MergeWorkbookSheets(wb, outSheet, ByRef rowOut, ByRef mergedSheets, ByRef mergedRows)
    On Error Resume Next
    Dim ws, rng, r, c, colCount, headerWritten
    headerWritten = False
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, outSheet.Name, vbTextCompare) <> 0 Then
            Set rng = ws.UsedRange
            If Err.Number <> 0 Then
                Err.Clear
            ElseIf Not rng Is Nothing Then
                If rng.Rows.Count > 0 And rng.Columns.Count > 0 Then
                    colCount = rng.Columns.Count
                    If Not headerWritten Then
                        For c = 1 To colCount
                            outSheet.Cells(1, c + 1).Value = rng.Cells(1, c).Text
                        Next
                        headerWritten = True
                    End If
                    For r = 2 To rng.Rows.Count
                        If RowHasContent(rng.Rows(r)) Then
                            outSheet.Cells(rowOut, 1).Value = ws.Name
                            For c = 1 To colCount
                                outSheet.Cells(rowOut, c + 1).Value = rng.Cells(r, c).Value
                            Next
                            rowOut = rowOut + 1
                            mergedRows = mergedRows + 1
                        End If
                    Next
                    mergedSheets = mergedSheets + 1
                End If
            End If
        End If
    Next
End Sub

Function RowHasContent(rowRange)
    On Error Resume Next
    Dim cell
    RowHasContent = False
    For Each cell In rowRange.Cells
        If Len(Trim(CStr(cell.Text))) > 0 Then
            RowHasContent = True
            Exit Function
        End If
    Next
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
    If methodName = "GetWorkbookName" Then
        SafeHostText = Host.GetWorkbookName()
    Else
        SafeHostText = ""
    End If
    If Err.Number <> 0 Then
        SafeHostText = ""
        Err.Clear
    End If
End Function

Function SafeRunBudget(sheetCount)
    On Error Resume Next
    SafeRunBudget = Host.GetRunBudgetPlan("excel_merge_worksheets", sheetCount, 200, 10, 180)
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
