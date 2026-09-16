' 函数名: HostExcelMergeWorkbooksPreviewCopy

' 描述: 选择文件夹内多个工作簿只读预检后，确认生成当前工作簿新表合并副本；带来源文件与工作表列，不改写源文件

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

        Main = FailureJson("E_NO_EXCEL_APP", "未取得 Excel 应用，请在 Excel 中运行该预设")

        Exit Function

    End If

    Main = HostExcelMergeWorkbooksPreviewCopy(appObj)

End Function

Function HostExcelMergeWorkbooksPreviewCopy(appObj)

    On Error Resume Next

    Dim activeWb

    Set activeWb = appObj.ActiveWorkbook

    If Err.Number <> 0 Or TypeName(activeWb) = "Empty" Or TypeName(activeWb) = "Nothing" Then

        Err.Clear

        HostExcelMergeWorkbooksPreviewCopy = FailureJson("E_NO_WORKBOOK", "当前没有活动工作簿")

        Exit Function

    End If

    Dim folderPath

    folderPath = Trim(SafeSelectFolder("选择包含待合并工作簿的文件夹（xls/xlsx/xlsm）"))

    If Len(folderPath) = 0 Then

        HostExcelMergeWorkbooksPreviewCopy = FailureJson("E_FOLDER_REQUIRED", "未选择文件夹；未修改工作簿")

        Exit Function

    End If

    Dim includeActiveText, includeActive

    includeActiveText = Trim(SafePrompt("是否包含当前活动工作簿工作表？是 / 否", "是"))

    includeActive = (StrComp(includeActiveText, "是", vbTextCompare) = 0 Or StrComp(includeActiveText, "Y", vbTextCompare) = 0 Or StrComp(includeActiveText, "YES", vbTextCompare) = 0)

    Dim paths(), pathCount, i

    pathCount = 0

    CollectPattern folderPath, "*.xls", paths, pathCount

    CollectPattern folderPath, "*.xlsx", paths, pathCount

    CollectPattern folderPath, "*.xlsm", paths, pathCount

    DeduplicatePaths paths, pathCount

    Dim activePathKey, filePath, fileName, skippedTemp, skippedActive

    activePathKey = ""

    On Error Resume Next

    If Len(CStr(activeWb.FullName)) > 0 Then activePathKey = NormalizePathKey(activeWb.FullName)

    Err.Clear

    Dim sourcePaths(), sourceNames(), sourceCount

    sourceCount = 0

    skippedTemp = 0

    skippedActive = 0

    For i = 0 To pathCount - 1

        filePath = paths(i)

        fileName = FileNameFromPath(filePath)

        If Left(fileName, 2) = "~$" Then

            skippedTemp = skippedTemp + 1

        ElseIf Len(activePathKey) > 0 And StrComp(NormalizePathKey(filePath), activePathKey, vbBinaryCompare) = 0 Then

            skippedActive = skippedActive + 1

        Else

            ReDim Preserve sourcePaths(sourceCount)

            ReDim Preserve sourceNames(sourceCount)

            sourcePaths(sourceCount) = filePath

            sourceNames(sourceCount) = fileName

            sourceCount = sourceCount + 1

        End If

    Next

    If sourceCount = 0 And Not includeActive Then

        HostExcelMergeWorkbooksPreviewCopy = FailureJson("E_NO_SOURCE", "未找到可合并的工作簿；未修改当前工作簿")

        Exit Function

    End If

    Dim previewFiles, previewSheets, previewRows, openedExternal, failedOpen

    Dim sampleText, sampleCount, wbObj, wsObj, rng, dataRows, isExternal

    previewFiles = 0

    previewSheets = 0

    previewRows = 0

    openedExternal = 0

    failedOpen = 0

    sampleText = ""

    sampleCount = 0

    Dim oldAlerts, oldUpdating

    oldAlerts = appObj.DisplayAlerts

    oldUpdating = appObj.ScreenUpdating

    appObj.DisplayAlerts = False

    appObj.ScreenUpdating = False

    If includeActive Then

        previewFiles = previewFiles + 1

        For Each wsObj In activeWb.Worksheets

            Set rng = Nothing

            Err.Clear

            Set rng = wsObj.UsedRange

            If Err.Number = 0 And Not rng Is Nothing Then

                If rng.Rows.Count >= 1 And rng.Columns.Count >= 1 Then

                    dataRows = CountContentRows(rng)

                    previewSheets = previewSheets + 1

                    previewRows = previewRows + dataRows

                    If sampleCount < 6 Then

                        sampleText = sampleText & "[当前]" & wsObj.Name & " 行=" & CStr(dataRows) & "; "

                        sampleCount = sampleCount + 1

                    End If

                End If

            End If

            Err.Clear

        Next

    End If

    For i = 0 To sourceCount - 1

        Set wbObj = Nothing

        Err.Clear

        Set wbObj = appObj.Workbooks.Open(sourcePaths(i), False, True)

        If Err.Number <> 0 Or TypeName(wbObj) = "Empty" Or TypeName(wbObj) = "Nothing" Then

            Err.Clear

            failedOpen = failedOpen + 1

        Else

            openedExternal = openedExternal + 1

            previewFiles = previewFiles + 1

            For Each wsObj In wbObj.Worksheets

                Set rng = Nothing

                Err.Clear

                Set rng = wsObj.UsedRange

                If Err.Number = 0 And Not rng Is Nothing Then

                    If rng.Rows.Count >= 1 And rng.Columns.Count >= 1 Then

                        dataRows = CountContentRows(rng)

                        previewSheets = previewSheets + 1

                        previewRows = previewRows + dataRows

                        If sampleCount < 6 Then

                            sampleText = sampleText & sourceNames(i) & "/" & wsObj.Name & " 行=" & CStr(dataRows) & "; "

                            sampleCount = sampleCount + 1

                        End If

                    End If

                End If

                Err.Clear

            Next

            wbObj.Close False

            Err.Clear

        End If

    Next

    appObj.DisplayAlerts = oldAlerts

    appObj.ScreenUpdating = oldUpdating

    If previewSheets = 0 Then

        HostExcelMergeWorkbooksPreviewCopy = FailureJson("E_NO_SHEETS", "预检未发现可合并工作表；源文件未改写")

        Exit Function

    End If

    Dim budgetJson, planId, planPreview, previewText

    budgetJson = SafeRunBudget(previewSheets)

    previewText = "Excel 多工作簿合并预检（尚未写入）" & vbCrLf & _
        "文件夹=" & folderPath & "；外部文件候选=" & CStr(sourceCount) & "；成功打开=" & CStr(openedExternal) & "；打开失败=" & CStr(failedOpen) & vbCrLf & _
        "包含当前簿=" & YesNo(includeActive) & "；预检文件=" & CStr(previewFiles) & "；工作表=" & CStr(previewSheets) & "；数据行约=" & CStr(previewRows) & vbCrLf & _
        "跳过临时文件=" & CStr(skippedTemp) & "；跳过当前簿路径重复=" & CStr(skippedActive) & vbCrLf & _
        "样例=" & sampleText & vbCrLf & _
        "输出=当前工作簿新表（来源文件/来源工作表列）；源文件只读打开且不会被保存改写" & vbCrLf & _
        "Host.GetRunBudgetPlan: " & budgetJson

    planId = SafeBeginWritePlan()

    SafeRecordWrite "excel_merge_workbooks_preview_copy", "{""folder"":""" & EscapeJson(folderPath) & """,""files"":" & CStr(previewFiles) & ",""sheets"":" & CStr(previewSheets) & ",""rows"":" & CStr(previewRows) & "}", "office.excel.mergeWorkbooks"

    planPreview = SafePreviewWritePlan()

    SafeRollbackWritePlan planId

    If StrComp(Trim(SafePrompt(previewText & vbCrLf & vbCrLf & "确认生成多工作簿合并副本，请输入：生成", "")), "生成", vbTextCompare) <> 0 Then

        HostExcelMergeWorkbooksPreviewCopy = FailureJson("E_CONFIRM_REQUIRED", "未输入“生成”，已取消且未修改工作簿")

        Exit Function

    End If

    Dim outputName, outputSheet, headerWritten, colCount, c, r, rowOut, mergedSheets, mergedRows, sourceLabel

    outputName = UniqueSheetName(activeWb, "多簿合并")

    Set outputSheet = activeWb.Worksheets.Add

    outputSheet.Name = outputName

    If Err.Number <> 0 Then

        Err.Clear

        SafeDeleteSheet appObj, outputSheet

        HostExcelMergeWorkbooksPreviewCopy = FailureJson("E_OUTPUT_SHEET", "无法创建合并结果工作表")

        Exit Function

    End If

    outputSheet.Cells(1, 1).Value = "来源文件"

    outputSheet.Cells(1, 2).Value = "来源工作表"

    headerWritten = False

    colCount = 0

    rowOut = 2

    mergedSheets = 0

    mergedRows = 0

    oldAlerts = appObj.DisplayAlerts

    oldUpdating = appObj.ScreenUpdating

    appObj.DisplayAlerts = False

    appObj.ScreenUpdating = False

    If includeActive Then

        MergeOneWorkbook activeWb, "[当前工作簿]", outputSheet, headerWritten, colCount, rowOut, mergedSheets, mergedRows

    End If

    For i = 0 To sourceCount - 1

        Set wbObj = Nothing

        Err.Clear

        Set wbObj = appObj.Workbooks.Open(sourcePaths(i), False, True)

        If Err.Number = 0 And Not (TypeName(wbObj) = "Empty" Or TypeName(wbObj) = "Nothing") Then

            MergeOneWorkbook wbObj, sourceNames(i), outputSheet, headerWritten, colCount, rowOut, mergedSheets, mergedRows

            wbObj.Close False

        End If

        Err.Clear

    Next

    outputSheet.Rows(1).Font.Bold = True

    outputSheet.Columns.AutoFit

    appObj.DisplayAlerts = oldAlerts

    appObj.ScreenUpdating = oldUpdating

    If Err.Number <> 0 Then

        Dim writeError

        writeError = Err.Description

        Err.Clear

        SafeDeleteSheet appObj, outputSheet

        HostExcelMergeWorkbooksPreviewCopy = FailureJson("E_MERGE_WRITE", "合并写入失败，已删除未完成工作表：" & writeError)

        Exit Function

    End If

    Dim summary

    summary = "Excel 多工作簿合并副本已生成；工作表=" & outputName & "；合并工作表=" & CStr(mergedSheets) & _
        "；合并数据行=" & CStr(mergedRows) & "；源文件未保存改写；当前原有工作表未改内容"

    Host.WriteClipboard summary

    SafeWriteLog summary

    HostExcelMergeWorkbooksPreviewCopy = "{""ok"":true,""sourceUnchanged"":true,""outputSheet"":""" & EscapeJson(outputName) & _
        """,""mergedSheets"":" & CStr(mergedSheets) & ",""mergedRows"":" & CStr(mergedRows) & _
        ",""previewFiles"":" & CStr(previewFiles) & ",""failedOpen"":" & CStr(failedOpen) & _
        ",""writePlanPreview"":""" & EscapeJson(planPreview) & """,""message"":""" & EscapeJson(summary) & """}"

End Function

Sub MergeOneWorkbook(wbObj, sourceFileLabel, outputSheet, ByRef headerWritten, ByRef colCount, ByRef rowOut, ByRef mergedSheets, ByRef mergedRows)

    On Error Resume Next

    Dim wsObj, rng, r, c

    For Each wsObj In wbObj.Worksheets

        Set rng = Nothing

        Err.Clear

        Set rng = wsObj.UsedRange

        If Err.Number = 0 And Not rng Is Nothing Then

            If rng.Rows.Count >= 1 And rng.Columns.Count >= 1 Then

                If Not headerWritten Then

                    colCount = rng.Columns.Count

                    For c = 1 To colCount

                        outputSheet.Cells(1, c + 2).Value = rng.Cells(1, c).Text

                    Next

                    headerWritten = True

                End If

                For r = 2 To rng.Rows.Count

                    If RowHasContent(rng.Rows(r)) Then

                        outputSheet.Cells(rowOut, 1).Value = sourceFileLabel

                        outputSheet.Cells(rowOut, 2).Value = wsObj.Name

                        For c = 1 To colCount

                            If c <= rng.Columns.Count Then

                                outputSheet.Cells(rowOut, c + 2).Value = rng.Cells(r, c).Value

                            End If

                        Next

                        rowOut = rowOut + 1

                        mergedRows = mergedRows + 1

                    End If

                Next

                mergedSheets = mergedSheets + 1

            End If

        End If

        Err.Clear

    Next

End Sub

Function CountContentRows(rng)

    Dim r, n

    n = 0

    If rng.Rows.Count < 2 Then

        CountContentRows = 0

        Exit Function

    End If

    For r = 2 To rng.Rows.Count

        If RowHasContent(rng.Rows(r)) Then n = n + 1

    Next

    CountContentRows = n

End Function

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

Sub DeduplicatePaths(ByRef paths, ByRef pathCount)

    Dim i, j, keepCount, kept()

    keepCount = 0

    For i = 0 To pathCount - 1

        Dim dup

        dup = False

        For j = 0 To keepCount - 1

            If StrComp(NormalizePathKey(paths(i)), NormalizePathKey(kept(j)), vbBinaryCompare) = 0 Then

                dup = True

                Exit For

            End If

        Next

        If Not dup Then

            ReDim Preserve kept(keepCount)

            kept(keepCount) = paths(i)

            keepCount = keepCount + 1

        End If

    Next

    pathCount = keepCount

    If keepCount = 0 Then

        Erase paths

    Else

        ReDim paths(keepCount - 1)

        For i = 0 To keepCount - 1

            paths(i) = kept(i)

        Next

    End If

End Sub

Function YesNo(value)

    If CBool(value) Then YesNo = "是" Else YesNo = "否"

End Function

Function SafeSelectFolder(promptText)

    On Error Resume Next

    SafeSelectFolder = Host.SelectFolder(promptText)

    If Err.Number <> 0 Then SafeSelectFolder = ""

    Err.Clear

End Function

Sub CollectPattern(folderPath, pattern, ByRef paths, ByRef pathCount)

    On Error Resume Next

    Dim payload

    payload = Host.EnumerateFiles(folderPath, pattern, False)

    If Err.Number <> 0 Then

        Err.Clear

        Exit Sub

    End If

    CollectFilesFromPayload payload, paths, pathCount

End Sub

Sub CollectFilesFromPayload(payload, ByRef paths, ByRef pathCount)

    Dim filesPos, arrayStart, pos, nextPath

    filesPos = InStr(1, payload, """files"":[", vbTextCompare)

    If filesPos <= 0 Then Exit Sub

    arrayStart = InStr(filesPos, payload, "[")

    If arrayStart <= 0 Then Exit Sub

    pos = arrayStart + 1

    Do

        nextPath = NextJsonString(payload, pos)

        If Len(nextPath) = 0 Then Exit Do

        ReDim Preserve paths(pathCount)

        paths(pathCount) = JsonUnescape(nextPath)

        pathCount = pathCount + 1

    Loop

End Sub

Function NextJsonString(text, ByRef pos)

    Dim q1, q2, ch, escaped, raw

    q1 = InStr(pos, text, """")

    If q1 <= 0 Then

        NextJsonString = ""

        Exit Function

    End If

    If InStr(pos, text, "]") > 0 And InStr(pos, text, "]") < q1 Then

        NextJsonString = ""

        Exit Function

    End If

    q2 = q1 + 1

    escaped = False

    Do While q2 <= Len(text)

        ch = Mid(text, q2, 1)

        If ch = "\" And Not escaped Then

            escaped = True

        ElseIf ch = """" And Not escaped Then

            Exit Do

        Else

            escaped = False

        End If

        q2 = q2 + 1

    Loop

    If q2 > Len(text) Then

        NextJsonString = ""

        Exit Function

    End If

    raw = Mid(text, q1 + 1, q2 - q1 - 1)

    pos = q2 + 1

    NextJsonString = raw

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

Function FileNameFromPath(pathText)

    Dim p1, p2, p

    p1 = InStrRev(pathText, "\")

    p2 = InStrRev(pathText, "/")

    p = p1

    If p2 > p Then p = p2

    If p > 0 Then

        FileNameFromPath = Mid(pathText, p + 1)

    Else

        FileNameFromPath = pathText

    End If

End Function

Function NormalizePathKey(pathText)

    Dim text

    text = LCase(Trim(CStr(pathText)))

    text = Replace(text, "/", "\")

    NormalizePathKey = text

End Function

Function UniqueSheetName(workbook, baseName)

    Dim candidate, suffix

    candidate = Left(baseName, 31)

    suffix = 1

    Do While SheetExists(workbook, candidate)

        suffix = suffix + 1

        candidate = Left(baseName, 27) & "_" & CStr(suffix)

    Loop

    UniqueSheetName = candidate

End Function

Function SheetExists(workbook, sheetName)

    Dim sheetObj

    SheetExists = False

    For Each sheetObj In workbook.Worksheets

        If StrComp(CStr(sheetObj.Name), CStr(sheetName), vbTextCompare) = 0 Then

            SheetExists = True

            Exit Function

        End If

    Next

End Function

Sub SafeDeleteSheet(appObj, sheetObj)

    On Error Resume Next

    Dim oldAlerts

    oldAlerts = appObj.DisplayAlerts

    appObj.DisplayAlerts = False

    sheetObj.Delete

    appObj.DisplayAlerts = oldAlerts

    Err.Clear

End Sub

Function SafeRunBudget(itemCount)

    On Error Resume Next

    SafeRunBudget = Host.GetRunBudgetPlan("excel_merge_workbooks", itemCount, 200, 10, 180)

    If Err.Number <> 0 Then SafeRunBudget = ""

    Err.Clear

End Function

Function SafePrompt(message, defaultValue)

    On Error Resume Next

    SafePrompt = Host.Prompt(message, defaultValue)

    If Err.Number <> 0 Then SafePrompt = defaultValue

    Err.Clear

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

Sub SafeRollbackWritePlan(planId)

    On Error Resume Next

    Host.RollbackWritePlan planId

    Err.Clear

End Sub

Sub SafeWriteLog(message)

    On Error Resume Next

    Host.WriteLog message

    Err.Clear

End Sub

Function FailureJson(code, message)

    FailureJson = "{""ok"":false,""code"":""" & EscapeJson(code) & """,""message"":""" & EscapeJson(message) & """}"

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
