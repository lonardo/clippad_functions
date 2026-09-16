' 函数名: HostExcelBuildWorksheetIndex
' 描述: 在当前工作簿中新建可点击的工作表目录，便于多工作表报表快速跳转
' 适用应用: Excel
' 搜索范围: 无
' 搜索对象: 无

Option Explicit

Const xlCenter = -4108
Const xlSolid = 1

Function Main()
    On Error Resume Next

    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = "{""ok"":false,""code"":""E_NO_EXCEL_APP"",""message"":""未取得 Excel 应用，请在 Excel 中运行该预设""}"
        Exit Function
    End If

    Main = HostExcelBuildWorksheetIndex(appObj)
End Function

Function HostExcelBuildWorksheetIndex(appObj)
    On Error Resume Next

    Dim workbook, indexSheet, sheet, rowIndex, planId
    Set workbook = appObj.ActiveWorkbook
    If Err.Number <> 0 Or TypeName(workbook) = "Empty" Or TypeName(workbook) = "Nothing" Then
        Err.Clear
        HostExcelBuildWorksheetIndex = "{""ok"":false,""code"":""E_NO_WORKBOOK"",""message"":""当前没有活动工作簿""}"
        Exit Function
    End If

    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_build_worksheet_index", "{""workbook"":""" & EscapeJson(workbook.Name) & """}", "office.excel.navigation"
    Set indexSheet = workbook.Worksheets.Add(workbook.Worksheets(1))
    indexSheet.Name = SafeSheetName(workbook, "目录")
    indexSheet.Range("A1").Value = "工作表目录"
    indexSheet.Range("A1:B1").Merge
    indexSheet.Range("A1").Font.Bold = True
    indexSheet.Range("A1").Font.Size = 16
    indexSheet.Range("A1").HorizontalAlignment = xlCenter
    indexSheet.Range("A1").Interior.Color = RGB(221, 235, 247)
    indexSheet.Range("A3").Value = "序号"
    indexSheet.Range("B3").Value = "工作表"
    indexSheet.Range("A3:B3").Font.Bold = True
    indexSheet.Range("A3:B3").Interior.Color = RGB(217, 225, 242)

    rowIndex = 4
    For Each sheet In workbook.Worksheets
        If sheet.Name <> indexSheet.Name Then
            indexSheet.Cells(rowIndex, 1).Value = rowIndex - 3
            indexSheet.Cells(rowIndex, 2).Value = sheet.Name
            indexSheet.Hyperlinks.Add indexSheet.Cells(rowIndex, 2), "", "'" & Replace(sheet.Name, "'", "''") & "'!A1", "", sheet.Name
            rowIndex = rowIndex + 1
        End If
    Next
    indexSheet.Columns("A:B").AutoFit
    indexSheet.Range("A3:B" & CStr(rowIndex - 1)).Borders.LineStyle = xlSolid
    SafeCloseWritePlan planId

    Dim count, summary
    count = rowIndex - 4
    summary = "Excel 工作表目录已创建；目录页=" & indexSheet.Name & "；可跳转工作表=" & CStr(count)
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostExcelBuildWorksheetIndex = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""sheetCount"":" & CStr(count) & ",""indexSheet"":""" & EscapeJson(indexSheet.Name) & """}"
End Function

Function SafeSheetName(workbook, baseName)
    On Error Resume Next
    Dim candidate, suffix, exists, sheet
    suffix = 1
    Do
        If suffix = 1 Then
            candidate = baseName
        Else
            candidate = baseName & CStr(suffix)
        End If
        exists = False
        For Each sheet In workbook.Worksheets
            If StrComp(sheet.Name, candidate, vbTextCompare) = 0 Then exists = True
        Next
        If Not exists Then
            SafeSheetName = candidate
            Exit Function
        End If
        suffix = suffix + 1
    Loop
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
    EscapeJson = text
End Function
