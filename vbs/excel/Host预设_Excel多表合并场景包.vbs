' 函数名: HostExcelMultiSheetMergePack
' 描述: 多表合并场景包：支持当前工作簿或文件夹批处理；先做表头同构/空表/打开失败预检，确认后合并到新表并输出异常清单与处理摘要，不改源表
' 适用应用: Excel
' 搜索范围: 全文
' 搜索对象: 无
' 作用范围: 全文
' 输出类型: 表

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
    Main = HostExcelMultiSheetMergePack(appObj)
End Function

Function HostExcelMultiSheetMergePack(appObj)
    On Error Resume Next
    Dim workbook, activeName, outputMode, sourceMode
    Set workbook = appObj.ActiveWorkbook
    If Err.Number <> 0 Or TypeName(workbook) = "Empty" Or TypeName(workbook) = "Nothing" Then
        Err.Clear
        HostExcelMultiSheetMergePack = FailureJson("E_NO_WORKBOOK", "当前没有活动工作簿")
        Exit Function
    End If
    activeName = workbook.ActiveSheet.Name
    sourceMode = NormalizeMergeSourceMode(SafePrompt("合并来源：当前工作簿 或 文件夹（默认 当前工作簿）", "当前工作簿"))
    outputMode = NormalizeOutputMode(SafePrompt("输出模式：NewSheet 或 NewWorkbook（默认 NewSheet）", "NewSheet"))

    Dim filePaths(), fileLabels(), fileIsExternal(), fileCount
    Dim failLines, warnLines, sampleLines
    Dim sheetCount, dataRows, emptySheets, openFailed, headerMismatch, skippedTemp
    Dim canonicalHeaders, canonicalColCount, folderPath, includeActive
    fileCount = 0: sheetCount = 0: dataRows = 0
    emptySheets = 0: openFailed = 0: headerMismatch = 0: skippedTemp = 0
    Dim headerStrategy, headerStrategyLabel, skippedMismatchSheets
    headerStrategy = "name_align": skippedMismatchSheets = 0
    headerStrategyLabel = HeaderStrategyLabel(headerStrategy)
    failLines = "": warnLines = "": sampleLines = ""
    canonicalHeaders = Array(): canonicalColCount = 0: folderPath = "": includeActive = False

    Dim oldAlerts, oldUpdating
    oldAlerts = appObj.DisplayAlerts
    oldUpdating = appObj.ScreenUpdating
    appObj.DisplayAlerts = False
    appObj.ScreenUpdating = False

    If sourceMode = "folder" Then
        folderPath = Trim(SafeSelectFolder("选择包含待合并工作簿的文件夹（xls/xlsx/xlsm）"))
        If Len(folderPath) = 0 Then
            appObj.DisplayAlerts = oldAlerts
            appObj.ScreenUpdating = oldUpdating
            HostExcelMultiSheetMergePack = FailureJson("E_FOLDER_REQUIRED", "未选择文件夹；未修改工作簿")
            Exit Function
        End If
        includeActive = IsYesText(SafePrompt("是否包含当前活动工作簿？是 / 否", "是"))
        Dim paths(), pathCount, i, filePath, fileName, activePathKey, wbObj
        pathCount = 0
        CollectPattern folderPath, "*.xls", paths, pathCount
        CollectPattern folderPath, "*.xlsx", paths, pathCount
        CollectPattern folderPath, "*.xlsm", paths, pathCount
        DeduplicatePaths paths, pathCount
        activePathKey = ""
        If Len(CStr(workbook.FullName)) > 0 Then activePathKey = NormalizePathKey(workbook.FullName)
        Err.Clear
        If includeActive Then
            AppendMergeFile filePaths, fileLabels, fileIsExternal, fileCount, "", "[当前工作簿]", False
            InspectWorkbookSources workbook, "[当前工作簿]", sheetCount, dataRows, emptySheets, sampleLines, canonicalHeaders, canonicalColCount, headerMismatch, warnLines
        End If
        For i = 0 To pathCount - 1
            filePath = paths(i)
            fileName = FileNameFromPath(filePath)
            If Left(fileName, 2) = "~$" Then
                skippedTemp = skippedTemp + 1
            ElseIf Len(activePathKey) > 0 And StrComp(NormalizePathKey(filePath), activePathKey, vbBinaryCompare) = 0 Then
                ' skip duplicate active workbook path
            Else
                Set wbObj = Nothing
                Err.Clear
                Set wbObj = appObj.Workbooks.Open(filePath, False, True)
                If Err.Number <> 0 Or TypeName(wbObj) = "Empty" Or TypeName(wbObj) = "Nothing" Then
                    Err.Clear
                    openFailed = openFailed + 1
                    If Len(failLines) > 0 Then failLines = failLines & vbCrLf
                    failLines = failLines & "打开失败" & vbTab & fileName & vbTab & "-" & vbTab & "无法只读打开"
                Else
                    AppendMergeFile filePaths, fileLabels, fileIsExternal, fileCount, filePath, fileName, True
                    InspectWorkbookSources wbObj, fileName, sheetCount, dataRows, emptySheets, sampleLines, canonicalHeaders, canonicalColCount, headerMismatch, warnLines
                    wbObj.Close False
                    Err.Clear
                End If
            End If
        Next
    Else
        If workbook.Worksheets.Count < 2 Then
            appObj.DisplayAlerts = oldAlerts
            appObj.ScreenUpdating = oldUpdating
            HostExcelMultiSheetMergePack = FailureJson("E_SINGLE_SHEET", "当前工作簿少于两个工作表，无需合并")
            Exit Function
        End If
        folderPath = "[当前工作簿]"
        AppendMergeFile filePaths, fileLabels, fileIsExternal, fileCount, "", "[当前工作簿]", False
        InspectWorkbookSources workbook, "[当前工作簿]", sheetCount, dataRows, emptySheets, sampleLines, canonicalHeaders, canonicalColCount, headerMismatch, warnLines
    End If

    appObj.DisplayAlerts = oldAlerts
    appObj.ScreenUpdating = oldUpdating

    If sheetCount = 0 Then
        HostExcelMultiSheetMergePack = FailureJson("E_NO_SOURCE", "预检后没有可合并的数据表；未修改工作簿")
        Exit Function
    End If

    If headerMismatch > 0 Then
        headerStrategy = NormalizeHeaderStrategy(SafePrompt("表头不一致处理：1按位置合并 / 2按名称对齐 / 3跳过不一致表（默认 2）", "2"))
    Else
        headerStrategy = "name_align"
    End If
    headerStrategyLabel = HeaderStrategyLabel(headerStrategy)

    Dim previewText, planId, planPreview
    previewText = "Excel 多表合并场景包预览（尚未写入）" & vbCrLf & _
        "命令ID：excel.multi_sheet_merge_pack" & vbCrLf & _
        "来源模式=" & sourceMode & "；路径=" & folderPath & vbCrLf & _
        "文件数=" & CStr(fileCount) & "；可合并表=" & CStr(sheetCount) & "；预估数据行=" & CStr(dataRows) & vbCrLf & _
        "空表=" & CStr(emptySheets) & "；打开失败=" & CStr(openFailed) & "；表头不一致=" & CStr(headerMismatch) & "；跳过临时文件=" & CStr(skippedTemp) & vbCrLf & _
        "标准列数=" & CStr(canonicalColCount) & "；输出模式=" & outputMode & vbCrLf & _
        "表头策略=" & headerStrategyLabel & "；跳过表=" & CStr(skippedMismatchSheets) & vbCrLf & _
        "样例表（最多 8 个）：" & vbCrLf & sampleLines & vbCrLf
    If Len(warnLines) > 0 Then previewText = previewText & "表头/结构警告：" & vbCrLf & warnLines & vbCrLf
    If Len(failLines) > 0 Then previewText = previewText & "失败项：" & vbCrLf & failLines & vbCrLf
    previewText = previewText & "将生成：合并结果（含来源文件/来源工作表）"
    If headerMismatch > 0 Or openFailed > 0 Or emptySheets > 0 Or Len(failLines) > 0 Or Len(warnLines) > 0 Then previewText = previewText & " + 异常清单"
    previewText = previewText & " + 处理摘要；sourceUnchanged=true"

    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_multi_sheet_merge_pack", "{""sourceMode"":""" & EscapeJson(sourceMode) & """,""files"":" & CStr(fileCount) & ",""sheets"":" & CStr(sheetCount) & ",""rows"":" & CStr(dataRows) & ",""headerMismatch"":" & CStr(headerMismatch) & ",""headerStrategy"":""" & EscapeJson(headerStrategy) & """,""skippedMismatchSheets"":" & CStr(skippedMismatchSheets) & ",""outputMode"":""" & EscapeJson(outputMode) & """}", "office.excel.multiSheetMergePack"
    planPreview = SafePreviewWritePlan()
    SafeRollbackWritePlan planId

    If Not SafeConfirmStep(previewText & vbCrLf & vbCrLf & "确认执行多表合并？", "excel.multi_sheet_merge_pack") Then
        HostExcelMultiSheetMergePack = FailureJson("E_CONFIRM_REQUIRED", "用户取消或未确认，未修改工作簿")
        Exit Function
    End If

    Dim createdSheets(), createdCount, mergeSheet, issueSheet, summarySheet
    Dim mergedSheets, mergedRows, writeFailLines
    createdCount = 0
    ReDim createdSheets(6)
    mergedSheets = 0
    mergedRows = 0
    writeFailLines = ""
    Set issueSheet = Nothing

    oldAlerts = appObj.DisplayAlerts
    oldUpdating = appObj.ScreenUpdating
    appObj.DisplayAlerts = False
    appObj.ScreenUpdating = False

    Set mergeSheet = CreateOutputSheet(appObj, activeName, "多表合并", outputMode)
    If mergeSheet Is Nothing Then
        appObj.DisplayAlerts = oldAlerts
        appObj.ScreenUpdating = oldUpdating
        HostExcelMultiSheetMergePack = FailureJson("E_OUTPUT_SHEET", "无法创建多表合并输出工作表")
        Exit Function
    End If
    createdCount = createdCount + 1: Set createdSheets(createdCount) = mergeSheet

    Dim headerWritten, colCount, rowOut, j
    headerWritten = False
    colCount = canonicalColCount
    rowOut = 2
    mergeSheet.Cells(1, 1).Value = "来源文件"
    mergeSheet.Cells(1, 2).Value = "来源工作表"

    For j = 0 To fileCount - 1
        If fileIsExternal(j) Then
            Set wbObj = Nothing
            Err.Clear
            Set wbObj = appObj.Workbooks.Open(filePaths(j), False, True)
            If Err.Number <> 0 Or TypeName(wbObj) = "Empty" Or TypeName(wbObj) = "Nothing" Then
                Err.Clear
                openFailed = openFailed + 1
                If Len(writeFailLines) > 0 Then writeFailLines = writeFailLines & vbCrLf
                writeFailLines = writeFailLines & "写入时打开失败" & vbTab & fileLabels(j) & vbTab & "-" & vbTab & "无法只读打开"
            Else
                MergeWorkbookToSheet wbObj, fileLabels(j), mergeSheet, headerWritten, colCount, rowOut, mergedSheets, mergedRows, canonicalHeaders, warnLines, headerStrategy, skippedMismatchSheets
                wbObj.Close False
                Err.Clear
            End If
        Else
            MergeWorkbookToSheet workbook, fileLabels(j), mergeSheet, headerWritten, colCount, rowOut, mergedSheets, mergedRows, canonicalHeaders, warnLines, headerStrategy, skippedMismatchSheets
        End If
    Next

    appObj.DisplayAlerts = oldAlerts
    appObj.ScreenUpdating = oldUpdating

    If mergedRows = 0 Then
        RollbackCreatedSheets appObj, createdSheets, createdCount
        HostExcelMultiSheetMergePack = FailureJson("E_MERGE_WRITE", "多表合并写入失败或无数据行，已删除未完成输出")
        Exit Function
    End If

    mergeSheet.Rows(1).Font.Bold = True
    mergeSheet.Columns.AutoFit

    If headerMismatch > 0 Or openFailed > 0 Or emptySheets > 0 Or skippedMismatchSheets > 0 Or Len(failLines) > 0 Or Len(warnLines) > 0 Or Len(writeFailLines) > 0 Then
        Set issueSheet = CreateOutputSheet(appObj, activeName, "异常清单", outputMode)
        If issueSheet Is Nothing Then
            RollbackCreatedSheets appObj, createdSheets, createdCount
            HostExcelMultiSheetMergePack = FailureJson("E_OUTPUT_SHEET", "无法创建异常清单工作表")
            Exit Function
        End If
        createdCount = createdCount + 1: Set createdSheets(createdCount) = issueSheet
        If Len(writeFailLines) > 0 Then
            If Len(failLines) > 0 Then failLines = failLines & vbCrLf
            failLines = failLines & writeFailLines
        End If
        WriteMergeIssueSheet issueSheet, warnLines, failLines, emptySheets, openFailed, headerMismatch
    End If

    Set summarySheet = CreateOutputSheet(appObj, activeName, "处理摘要", outputMode)
    If summarySheet Is Nothing Then
        RollbackCreatedSheets appObj, createdSheets, createdCount
        HostExcelMultiSheetMergePack = FailureJson("E_OUTPUT_SHEET", "无法创建处理摘要工作表")
        Exit Function
    End If
    createdCount = createdCount + 1: Set createdSheets(createdCount) = summarySheet

    Dim summaryLines, impactJson, summary, issueSheetName
    issueSheetName = ""
    If Not issueSheet Is Nothing Then issueSheetName = issueSheet.Name
    summaryLines = "场景名=多表合并场景包" & vbCrLf & _
        "命令ID=excel.multi_sheet_merge_pack" & vbCrLf & _
        "源表名称=" & activeName & vbCrLf & _
        "来源模式=" & sourceMode & vbCrLf & _
        "来源路径=" & folderPath & vbCrLf & _
        "输出模式=" & outputMode & vbCrLf & _
        "输出位置=" & mergeSheet.Name & IIf(Len(issueSheetName) > 0, "," & issueSheetName, "") & "," & summarySheet.Name & vbCrLf & _
        "处理前行数=" & CStr(dataRows) & vbCrLf & _
        "处理后行数=" & CStr(mergedRows) & vbCrLf & _
        "参数=files=" & CStr(fileCount) & ";mergeableSheets=" & CStr(sheetCount) & ";columns=" & CStr(colCount) & vbCrLf & _
        "合并工作表数=" & CStr(mergedSheets) & vbCrLf & _
        "合并数据行=" & CStr(mergedRows) & vbCrLf & _
        "空表=" & CStr(emptySheets) & vbCrLf & _
        "打开失败=" & CStr(openFailed) & vbCrLf & _
        "表头不一致=" & CStr(headerMismatch) & vbCrLf & _
        "表头策略=" & headerStrategyLabel & vbCrLf & _
        "跳过表=" & CStr(skippedMismatchSheets) & vbCrLf & _
        "执行时间=" & Now & vbCrLf & _
        "结果=成功"
    impactJson = SafeHostText("ExcelWriteImpactSummary", summaryLines, summarySheet.Name)
    If Not ExtractJsonBoolean(impactJson, "ok") Then WriteSummaryFallback summarySheet, summaryLines

    summary = "Excel 多表合并完成：输出=" & mergeSheet.Name & "；合并表=" & CStr(mergedSheets) & "；行=" & CStr(mergedRows) & "；表头策略=" & headerStrategyLabel & "；跳过表=" & CStr(skippedMismatchSheets) & "；sourceUnchanged=true"
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostExcelMultiSheetMergePack = "{""ok"":true,""sourceUnchanged"":true,""commandId"":""excel.multi_sheet_merge_pack"",""outputSheet"":""" & EscapeJson(mergeSheet.Name) & """,""summarySheet"":""" & EscapeJson(summarySheet.Name) & """,""issueSheet"":""" & EscapeJson(issueSheetName) & """,""mergedSheets"":" & CStr(mergedSheets) & ",""mergedRows"":" & CStr(mergedRows) & ",""headerMismatch"":" & CStr(headerMismatch) & ",""headerStrategy"":""" & EscapeJson(headerStrategy) & """,""headerStrategyLabel"":""" & EscapeJson(headerStrategyLabel) & """,""skippedMismatchSheets"":" & CStr(skippedMismatchSheets) & ",""openFailed"":" & CStr(openFailed) & ",""writePlanPreview"":""" & EscapeJson(planPreview) & """,""message"":""" & EscapeJson(summary) & """}"
End Function

Function NormalizeMergeSourceMode(text)
    Dim t
    t = LCase(Trim(CStr(text)))
    If t = "folder" Or t = "dir" Or t = "目录" Or t = "文件夹" Or t = "多工作簿" Then
        NormalizeMergeSourceMode = "folder"
    Else
        NormalizeMergeSourceMode = "workbook"
    End If
End Function

Function IsYesText(text)
    Dim t
    t = LCase(Trim(CStr(text)))
    IsYesText = (t = "是" Or t = "y" Or t = "yes" Or t = "true" Or t = "1")
End Function

Function CountContentRows(rng)
    Dim r, total
    total = 0
    If rng Is Nothing Then
        CountContentRows = 0
        Exit Function
    End If
    For r = 2 To rng.Rows.Count
        If RowHasContent(rng.Rows(r)) Then total = total + 1
    Next
    CountContentRows = total
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
    filesPos = InStr(1, CStr(payload), """files"":[", vbTextCompare)
    If filesPos <= 0 Then
        filesPos = InStr(1, CStr(payload), "[")
        If filesPos <= 0 Then Exit Sub
        arrayStart = filesPos
    Else
        arrayStart = InStr(filesPos, CStr(payload), "[")
    End If
    If arrayStart <= 0 Then Exit Sub
    pos = arrayStart + 1
    Do
        nextPath = NextJsonString(CStr(payload), pos)
        If Len(nextPath) = 0 Then Exit Do
        ReDim Preserve paths(pathCount)
        paths(pathCount) = JsonUnescape(nextPath)
        pathCount = pathCount + 1
    Loop
End Sub

Function NextJsonString(text, ByRef pos)
    Dim startPos, i, ch, buf, bs
    NextJsonString = ""
    bs = Chr(92)
    startPos = InStr(pos, text, """")
    If startPos <= 0 Then Exit Function
    buf = ""
    i = startPos + 1
    Do While i <= Len(text)
        ch = Mid(text, i, 1)
        If ch = bs Then
            If i < Len(text) Then
                buf = buf & Mid(text, i + 1, 1)
                i = i + 2
            Else
                Exit Do
            End If
        ElseIf ch = """" Then
            pos = i + 1
            NextJsonString = buf
            Exit Function
        Else
            buf = buf & ch
            i = i + 1
        End If
    Loop
End Function

Function JsonUnescape(text)
    Dim t, bs
    bs = Chr(92)
    t = CStr(text)
    t = Replace(t, bs & bs, bs)
    t = Replace(t, bs & Chr(34), Chr(34))
    t = Replace(t, bs & "/", "/")
    JsonUnescape = t
End Function

Sub DeduplicatePaths(ByRef paths, ByRef pathCount)
    Dim i, j, keepCount, kept(), dup
    keepCount = 0
    For i = 0 To pathCount - 1
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

Function NormalizePathKey(pathText)
    Dim t
    t = LCase(Trim(CStr(pathText)))
    t = Replace(t, "/", "\")
    NormalizePathKey = t
End Function

Function FileNameFromPath(pathText)
    Dim t, p
    t = Replace(CStr(pathText), "/", "\")
    p = InStrRev(t, "\")
    If p > 0 Then FileNameFromPath = Mid(t, p + 1) Else FileNameFromPath = t
End Function

Sub AppendMergeFile(ByRef filePaths, ByRef fileLabels, ByRef fileIsExternal, ByRef fileCount, pathText, labelText, isExternal)
    ReDim Preserve filePaths(fileCount)
    ReDim Preserve fileLabels(fileCount)
    ReDim Preserve fileIsExternal(fileCount)
    filePaths(fileCount) = pathText
    fileLabels(fileCount) = labelText
    fileIsExternal(fileCount) = isExternal
    fileCount = fileCount + 1
End Sub

Function CaptureHeaders(rng)
    Dim c, arr()
    ReDim arr(rng.Columns.Count - 1)
    For c = 1 To rng.Columns.Count
        arr(c - 1) = NormalizeHeaderName(rng.Cells(1, c).Text)
    Next
    CaptureHeaders = arr
End Function

Function NormalizeHeaderName(value)
    Dim text
    text = NormalizeCellText(value)
    text = Replace(text, " ", "")
    text = Replace(text, "_", "")
    text = Replace(text, "-", "")
    NormalizeHeaderName = LCase(text)
End Function

Function NormalizeHeaderStrategy(value)
    Dim t
    t = LCase(Trim(CStr(value)))
    If t = "1" Or t = "position" Or t = "pos" Or InStr(1, t, "按位置", vbTextCompare) > 0 Then
        NormalizeHeaderStrategy = "position"
    ElseIf t = "3" Or t = "skip" Or InStr(1, t, "跳过", vbTextCompare) > 0 Then
        NormalizeHeaderStrategy = "skip"
    Else
        NormalizeHeaderStrategy = "name_align"
    End If
End Function

Function HeaderStrategyLabel(value)
    Dim m
    m = NormalizeHeaderStrategy(value)
    If m = "position" Then
        HeaderStrategyLabel = "按位置合并"
    ElseIf m = "skip" Then
        HeaderStrategyLabel = "跳过不一致表"
    Else
        HeaderStrategyLabel = "按名称对齐"
    End If
End Function

Function BuildHeaderIndex(headers)
    Dim i, key, buf
    buf = ""
    If IsArray(headers) Then
        For i = 0 To UBound(headers)
            key = LCase(NormalizeCellText(headers(i)))
            If Len(key) = 0 Then key = "#" & CStr(i + 1)
            If Len(buf) > 0 Then buf = buf & vbLf
            buf = buf & key & "=" & CStr(i + 1)
        Next
    End If
    BuildHeaderIndex = buf
End Function

Function FindHeaderIndex(indexText, headerName, fallbackPos)
    Dim lines, i, lineText, key, prefix
    FindHeaderIndex = 0
    key = LCase(NormalizeCellText(headerName))
    If Len(key) = 0 Then
        FindHeaderIndex = fallbackPos
        Exit Function
    End If
    prefix = key & "="
    lines = Split(CStr(indexText), vbLf)
    For i = 0 To UBound(lines)
        lineText = Trim(lines(i))
        If LCase(Left(lineText, Len(prefix))) = prefix Then
            FindHeaderIndex = CLng(Mid(lineText, Len(prefix) + 1))
            Exit Function
        End If
    Next
    FindHeaderIndex = 0
End Function


Function CompareHeaders(canonicalHeaders, headers)
    Dim i, leftN, rightN, missing
    CompareHeaders = ""
    leftN = UBound(canonicalHeaders) + 1
    rightN = UBound(headers) + 1
    If leftN <> rightN Then
        CompareHeaders = "列数不同 标准=" & CStr(leftN) & " 当前=" & CStr(rightN)
        Exit Function
    End If
    missing = ""
    For i = 0 To UBound(canonicalHeaders)
        If StrComp(CStr(canonicalHeaders(i)), CStr(headers(i)), vbBinaryCompare) <> 0 Then
            If Len(missing) > 0 Then missing = missing & ";"
            missing = missing & "第" & CStr(i + 1) & "列[" & CStr(canonicalHeaders(i)) & "!=" & CStr(headers(i)) & "]"
        End If
    Next
    CompareHeaders = missing
End Function

Sub InspectWorkbookSources(wbObj, fileLabel, ByRef sheetCount, ByRef dataRows, ByRef emptySheets, ByRef sampleLines, ByRef canonicalHeaders, ByRef canonicalColCount, ByRef headerMismatch, ByRef warnLines)
    On Error Resume Next
    Dim ws, rng, rowCount, headers, mismatchNote
    For Each ws In wbObj.Worksheets
        Set rng = Nothing
        Err.Clear
        Set rng = ws.UsedRange
        If Err.Number <> 0 Or rng Is Nothing Then
            Err.Clear
            emptySheets = emptySheets + 1
            If Len(warnLines) > 0 Then warnLines = warnLines & vbCrLf
            warnLines = warnLines & "空表/无UsedRange" & vbTab & fileLabel & vbTab & ws.Name & vbTab & "无可用区域"
        ElseIf rng.Rows.Count <= 1 Or rng.Columns.Count < 1 Then
            emptySheets = emptySheets + 1
            If Len(warnLines) > 0 Then warnLines = warnLines & vbCrLf
            warnLines = warnLines & "空表" & vbTab & fileLabel & vbTab & ws.Name & vbTab & "仅有表头或无数据"
        Else
            rowCount = CountContentRows(rng)
            If rowCount <= 0 Then
                emptySheets = emptySheets + 1
                If Len(warnLines) > 0 Then warnLines = warnLines & vbCrLf
                warnLines = warnLines & "空表" & vbTab & fileLabel & vbTab & ws.Name & vbTab & "无内容行"
            Else
                headers = CaptureHeaders(rng)
                If canonicalColCount = 0 Then
                    canonicalHeaders = headers
                    canonicalColCount = UBound(headers) + 1
                Else
                    mismatchNote = CompareHeaders(canonicalHeaders, headers)
                    If Len(mismatchNote) > 0 Then
                        headerMismatch = headerMismatch + 1
                        If Len(warnLines) > 0 Then warnLines = warnLines & vbCrLf
                        warnLines = warnLines & "表头不一致" & vbTab & fileLabel & vbTab & ws.Name & vbTab & mismatchNote
                    End If
                End If
                sheetCount = sheetCount + 1
                dataRows = dataRows + rowCount
                If sheetCount <= 8 Then
                    If Len(sampleLines) > 0 Then sampleLines = sampleLines & vbCrLf
                    sampleLines = sampleLines & "- " & fileLabel & " / " & ws.Name & " : " & CStr(rowCount) & " 行；列=" & CStr(rng.Columns.Count)
                End If
            End If
        End If
    Next
End Sub

Sub MergeWorkbookToSheet(wbObj, sourceFileLabel, outputSheet, ByRef headerWritten, ByRef colCount, ByRef rowOut, ByRef mergedSheets, ByRef mergedRows, ByRef canonicalHeaders, ByRef warnLines, headerStrategy, ByRef skippedMismatchSheets)
    On Error Resume Next
    Dim wsObj, rng, r, c, headers, note, strategy, srcIndex, srcCol, canonIndexText, headerName
    strategy = NormalizeHeaderStrategy(headerStrategy)
    canonIndexText = BuildHeaderIndex(canonicalHeaders)
    For Each wsObj In wbObj.Worksheets
        Set rng = Nothing
        Err.Clear
        Set rng = wsObj.UsedRange
        If Err.Number = 0 And Not rng Is Nothing Then
            If rng.Rows.Count >= 2 And rng.Columns.Count >= 1 Then
                If CountContentRows(rng) > 0 Then
                    headers = CaptureHeaders(rng)
                    note = ""
                    If headerWritten Then
                        note = CompareHeaders(canonicalHeaders, headers)
                    End If
                    If headerWritten And Len(note) > 0 And strategy = "skip" Then
                        skippedMismatchSheets = skippedMismatchSheets + 1
                        If Len(warnLines) > 0 Then warnLines = warnLines & vbCrLf
                        warnLines = warnLines & "跳过表头不一致" & vbTab & sourceFileLabel & vbTab & wsObj.Name & vbTab & note
                    Else
                        If Not headerWritten Then
                            If colCount <= 0 Then
                                If IsArray(canonicalHeaders) Then
                                    colCount = UBound(canonicalHeaders) + 1
                                Else
                                    colCount = rng.Columns.Count
                                End If
                            End If
                            For c = 1 To colCount
                                If IsArray(canonicalHeaders) And c - 1 <= UBound(canonicalHeaders) Then
                                    outputSheet.Cells(1, c + 2).Value = canonicalHeaders(c - 1)
                                ElseIf c <= rng.Columns.Count Then
                                    outputSheet.Cells(1, c + 2).Value = rng.Cells(1, c).Text
                                Else
                                    outputSheet.Cells(1, c + 2).Value = "列" & CStr(c)
                                End If
                            Next
                            headerWritten = True
                            canonIndexText = BuildHeaderIndex(canonicalHeaders)
                        ElseIf Len(note) > 0 Then
                            If Len(warnLines) > 0 Then warnLines = warnLines & vbCrLf
                            If strategy = "name_align" Then
                                warnLines = warnLines & "按名称对齐合并" & vbTab & sourceFileLabel & vbTab & wsObj.Name & vbTab & note
                            Else
                                warnLines = warnLines & "按位置兼容合并" & vbTab & sourceFileLabel & vbTab & wsObj.Name & vbTab & note
                            End If
                        End If
                        srcIndex = BuildHeaderIndex(headers)
                        For r = 2 To rng.Rows.Count
                            If RowHasContent(rng.Rows(r)) Then
                                outputSheet.Cells(rowOut, 1).Value = sourceFileLabel
                                outputSheet.Cells(rowOut, 2).Value = wsObj.Name
                                For c = 1 To colCount
                                    srcCol = 0
                                    If strategy = "name_align" And IsArray(canonicalHeaders) And c - 1 <= UBound(canonicalHeaders) Then
                                        headerName = canonicalHeaders(c - 1)
                                        srcCol = FindHeaderIndex(srcIndex, headerName, 0)
                                    Else
                                        srcCol = c
                                    End If
                                    If srcCol > 0 And srcCol <= rng.Columns.Count Then
                                        outputSheet.Cells(rowOut, c + 2).Value = rng.Cells(r, srcCol).Value
                                    End If
                                Next
                                rowOut = rowOut + 1
                                mergedRows = mergedRows + 1
                            End If
                        Next
                        mergedSheets = mergedSheets + 1
                    End If
                End If
            End If
        End If
        Err.Clear
    Next
End Sub

Sub WriteMergeIssueSheet(sheetObj, warnLines, failLines, emptySheets, openFailed, headerMismatch)
    Dim lines, i, parts, outRow, allText
    sheetObj.Cells(1, 1).Value = "类型"
    sheetObj.Cells(1, 2).Value = "来源文件"
    sheetObj.Cells(1, 3).Value = "工作表"
    sheetObj.Cells(1, 4).Value = "说明"
    sheetObj.Cells(1, 5).Value = "处理状态"
    outRow = 1
    allText = ""
    If Len(warnLines) > 0 Then allText = warnLines
    If Len(failLines) > 0 Then
        If Len(allText) > 0 Then allText = allText & vbCrLf
        allText = allText & failLines
    End If
    If Len(allText) > 0 Then
        lines = Split(allText, vbCrLf)
        For i = 0 To UBound(lines)
            If Len(Trim(lines(i))) > 0 Then
                parts = Split(lines(i), vbTab)
                outRow = outRow + 1
                If UBound(parts) >= 0 Then sheetObj.Cells(outRow, 1).Value = parts(0)
                If UBound(parts) >= 1 Then sheetObj.Cells(outRow, 2).Value = parts(1)
                If UBound(parts) >= 2 Then sheetObj.Cells(outRow, 3).Value = parts(2)
                If UBound(parts) >= 3 Then sheetObj.Cells(outRow, 4).Value = parts(3) Else sheetObj.Cells(outRow, 4).Value = lines(i)
                sheetObj.Cells(outRow, 5).Value = "待处理"
            End If
        Next
    End If
    outRow = outRow + 1
    sheetObj.Cells(outRow, 1).Value = "汇总"
    sheetObj.Cells(outRow, 4).Value = "空表=" & CStr(emptySheets) & "；打开失败=" & CStr(openFailed) & "；表头不一致=" & CStr(headerMismatch)
    sheetObj.Cells(outRow, 5).Value = "待复核"
    sheetObj.Rows(1).Font.Bold = True
    sheetObj.Columns.AutoFit
End Sub

Function IsNothing(value)

    IsNothing = (TypeName(value) = "Nothing" Or TypeName(value) = "Empty")

End Function

Function ValuesEqual(sourceValue, targetValue)

    ValuesEqual = (StrComp(NormalizeCellText(sourceValue), NormalizeCellText(targetValue), vbTextCompare) = 0)

End Function

Function NormalizeExportFormat(value)

    Dim t

    t = LCase(Trim(CStr(value)))

    If t = "tsv" Or t = "tab" Or t = "制表符" Then

        NormalizeExportFormat = "tsv"

    Else

        NormalizeExportFormat = "csv"

    End If

End Function

Function FormatFileStamp(value)

    Dim d

    d = CDate(value)

    FormatFileStamp = Year(d) & Right("0" & Month(d), 2) & Right("0" & Day(d), 2) & "_" & Right("0" & Hour(d), 2) & Right("0" & Minute(d), 2) & Right("0" & Second(d), 2)

End Function

Function NormalizeLogic(value)

    Dim t

    t = UCase(Trim(CStr(value)))

    If t = "OR" Or t = "或" Or t = "或者" Then

        NormalizeLogic = "OR"

    Else

        NormalizeLogic = "AND"

    End If

End Function

Function NormalizeMatchMode(value)

    Dim t

    t = LCase(Trim(CStr(value)))

    If t = "contains" Or t = "包含" Or t = "like" Or t = "fuzzy" Then

        NormalizeMatchMode = "contains"

    Else

        NormalizeMatchMode = "exact"

    End If

End Function

Function MatchModeLabel(value)

    If NormalizeMatchMode(value) = "contains" Then

        MatchModeLabel = "包含"

    Else

        MatchModeLabel = "精确"

    End If

End Function

Function IsMatch(cellText, keyword, matchMode)

    Dim leftText, rightText

    leftText = NormalizeCellText(cellText)

    rightText = NormalizeCellText(keyword)

    If Len(rightText) = 0 Then

        IsMatch = (Len(leftText) = 0)

        Exit Function

    End If

    If NormalizeMatchMode(matchMode) = "contains" Then

        IsMatch = (InStr(1, leftText, rightText, vbTextCompare) > 0)

    Else

        IsMatch = (StrComp(leftText, rightText, vbTextCompare) = 0)

    End If

End Function

Function NormalizeMaskKind(value)

    Dim t

    t = LCase(Trim(CStr(value)))

    If t = "phone" Or t = "mobile" Or t = "手机" Or t = "手机号" Then

        NormalizeMaskKind = "phone"

    ElseIf t = "id" Or t = "idcard" Or t = "身份证" Then

        NormalizeMaskKind = "idcard"

    ElseIf t = "email" Or t = "mail" Or t = "邮箱" Then

        NormalizeMaskKind = "email"

    Else

        NormalizeMaskKind = "auto"

    End If

End Function

Function MaskKindLabel(value)

    Dim k

    k = NormalizeMaskKind(value)

    If k = "phone" Then

        MaskKindLabel = "手机号"

    ElseIf k = "idcard" Then

        MaskKindLabel = "身份证"

    ElseIf k = "email" Then

        MaskKindLabel = "邮箱"

    Else

        MaskKindLabel = "自动"

    End If

End Function

Function DigitsOnly(value)

    Dim i, ch, buf

    buf = ""

    For i = 1 To Len(CStr(value))

        ch = Mid(CStr(value), i, 1)

        If ch >= "0" And ch <= "9" Then buf = buf & ch

    Next

    DigitsOnly = buf

End Function

Function LooksLikeIdCard(value)

    Dim digits, text

    text = UCase(Trim(CStr(value)))

    digits = DigitsOnly(text)

    LooksLikeIdCard = (Len(digits) = 15 Or Len(digits) = 18 Or (Len(text) = 18 And Right(text, 1) = "X"))

End Function

Function MaskSensitiveValue(value, maskKind)

    Dim text, digits, kind, atPos, localPart, domainPart

    MaskSensitiveValue = ""

    text = NormalizeCellText(value)

    If Len(text) = 0 Then Exit Function

    kind = NormalizeMaskKind(maskKind)

    digits = DigitsOnly(text)

    If kind = "auto" Then

        If InStr(text, "@") > 0 Then

            kind = "email"

        ElseIf LooksLikeIdCard(text) Then

            kind = "idcard"

        ElseIf Len(digits) = 11 And Left(digits, 1) = "1" Then

            kind = "phone"

        Else

            Exit Function

        End If

    End If

    If kind = "phone" Then

        If Len(digits) < 7 Then Exit Function

        MaskSensitiveValue = Left(digits, 3) & "****" & Right(digits, 4)

    ElseIf kind = "idcard" Then

        If Len(digits) < 8 Then Exit Function

        If Len(digits) >= 18 Then

            MaskSensitiveValue = Left(digits, 4) & "**********" & Right(digits, 4)

        Else

            MaskSensitiveValue = Left(digits, 3) & "*********" & Right(digits, 3)

        End If

    ElseIf kind = "email" Then

        atPos = InStr(text, "@")

        If atPos <= 1 Then Exit Function

        localPart = Left(text, atPos - 1)

        domainPart = Mid(text, atPos)

        If Len(localPart) <= 1 Then

            MaskSensitiveValue = "*" & domainPart

        Else

            MaskSensitiveValue = Left(localPart, 1) & "***" & domainPart

        End If

    End If

End Function

Function RowHasContent(rowRange)

    Dim cell

    RowHasContent = False

    For Each cell In rowRange.Cells

        If Len(NormalizeCellText(cell.Text)) > 0 Then

            RowHasContent = True

            Exit Function

        End If

    Next

End Function

Function IIf(condition, trueValue, falseValue)

    If condition Then

        IIf = trueValue

    Else

        IIf = falseValue

    End If

End Function

Function ExtractJsonBoolean(jsonText, keyName)

    Dim marker, keyPos, colonPos, valueText

    ExtractJsonBoolean = False

    marker = """" & CStr(keyName) & """"

    keyPos = InStr(1, CStr(jsonText), marker, vbTextCompare)

    If keyPos <= 0 Then Exit Function

    colonPos = InStr(keyPos + Len(marker), CStr(jsonText), ":")

    If colonPos <= 0 Then Exit Function

    valueText = LCase(LTrim(Mid(CStr(jsonText), colonPos + 1)))

    ExtractJsonBoolean = (Left(valueText, 4) = "true")

End Function

Function ExtractJsonString(jsonText, keyName)

    Dim marker, keyPos, colonPos, quotePos, i, ch, nextCh, result

    ExtractJsonString = ""

    marker = """" & CStr(keyName) & """"

    keyPos = InStr(1, CStr(jsonText), marker, vbTextCompare)

    If keyPos <= 0 Then Exit Function

    colonPos = InStr(keyPos + Len(marker), CStr(jsonText), ":")

    If colonPos <= 0 Then Exit Function

    quotePos = InStr(colonPos + 1, CStr(jsonText), """")

    If quotePos <= 0 Then Exit Function

    result = ""

    i = quotePos + 1

    Do While i <= Len(jsonText)

        ch = Mid(jsonText, i, 1)

        If ch = """" Then Exit Do

        If ch = "\" And i < Len(jsonText) Then

            nextCh = Mid(jsonText, i + 1, 1)

            result = result & nextCh

            i = i + 2

        Else

            result = result & ch

            i = i + 1

        End If

    Loop

    ExtractJsonString = result

End Function

Function ExtractJsonNumber(jsonText, keyName)

    Dim marker, keyPos, colonPos, valueText, ch, i, buf

    ExtractJsonNumber = 0

    marker = """" & CStr(keyName) & """"

    keyPos = InStr(1, CStr(jsonText), marker, vbTextCompare)

    If keyPos <= 0 Then Exit Function

    colonPos = InStr(keyPos + Len(marker), CStr(jsonText), ":")

    If colonPos <= 0 Then Exit Function

    valueText = LTrim(Mid(CStr(jsonText), colonPos + 1))

    buf = ""

    For i = 1 To Len(valueText)

        ch = Mid(valueText, i, 1)

        If (ch >= "0" And ch <= "9") Or ch = "-" Or ch = "." Then

            buf = buf & ch

        ElseIf Len(buf) > 0 Then

            Exit For

        End If

    Next

    If Len(buf) > 0 Then ExtractJsonNumber = CDbl(buf)

End Function

Function ParseJsonIntArray(jsonText, keyName)

    Dim marker, keyPos, bracketStart, bracketEnd, body, parts, i, n, values()

    ParseJsonIntArray = Array()

    marker = """" & CStr(keyName) & """"

    keyPos = InStr(1, CStr(jsonText), marker, vbTextCompare)

    If keyPos <= 0 Then Exit Function

    bracketStart = InStr(keyPos, CStr(jsonText), "[")

    If bracketStart <= 0 Then Exit Function

    bracketEnd = InStr(bracketStart + 1, CStr(jsonText), "]")

    If bracketEnd <= bracketStart Then Exit Function

    body = Trim(Mid(CStr(jsonText), bracketStart + 1, bracketEnd - bracketStart - 1))

    If Len(body) = 0 Then Exit Function

    parts = Split(body, ",")

    n = -1

    ReDim values(UBound(parts))

    For i = 0 To UBound(parts)

        If Len(Trim(parts(i))) > 0 Then

            n = n + 1

            values(n) = CLng(Trim(parts(i)))

        End If

    Next

    If n >= 0 Then

        ReDim Preserve values(n)

        ParseJsonIntArray = values

    End If

End Function

Function FailureJson(code, message)
    FailureJson = "{""ok"":false,""code"":""" & EscapeJson(code) & """,""message"":""" & EscapeJson(message) & """,""sourceUnchanged"":true}"
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

Function SafePrompt(message, defaultValue)

    On Error Resume Next

    SafePrompt = Host.Prompt(message, defaultValue)

    If Err.Number <> 0 Then

        SafePrompt = CStr(defaultValue)

        Err.Clear

    End If

End Function

Function SafeConfirmStep(message, stepInfo)

    On Error Resume Next

    Dim confirmed, answer

    confirmed = Host.ConfirmStep(message, stepInfo)

    If Err.Number = 0 Then

        SafeConfirmStep = CBool(confirmed)

        Exit Function

    End If

    Err.Clear

    answer = Trim(SafePrompt(message & vbCrLf & vbCrLf & "确认继续请输入：确认", ""))

    SafeConfirmStep = (StrComp(answer, "确认", vbTextCompare) = 0)

End Function

Function SafeBeginWritePlan()

    On Error Resume Next

    SafeBeginWritePlan = Host.BeginWritePlan("")

    If Err.Number <> 0 Then SafeBeginWritePlan = ""

    Err.Clear

End Function

Sub SafeRecordWrite(actionId, paramsJson, capabilityId)

    On Error Resume Next

    Host.RecordWrite actionId, paramsJson, "office", "excel.sheet.delete", "{}", capabilityId

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

Function SafeHostText(methodName, arg1, arg2, arg3, arg4)

    On Error Resume Next

    SafeHostText = ""

    If methodName = "ExcelGetHeaderMap" Then

        SafeHostText = Host.ExcelGetHeaderMap(arg1)

    ElseIf methodName = "ExcelCreateOutputSheet" Then

        SafeHostText = Host.ExcelCreateOutputSheet(arg1, arg2, arg3, arg4)

    ElseIf methodName = "ExcelWriteImpactSummary" Then

        SafeHostText = Host.ExcelWriteImpactSummary(arg1, arg2)

    ElseIf methodName = "ExcelPreviewRowSamples" Then

        SafeHostText = Host.ExcelPreviewRowSamples(arg1, arg2)

    ElseIf methodName = "ExcelDedupePlan" Then

        SafeHostText = Host.ExcelDedupePlan(arg1, arg2)

    ElseIf methodName = "ExcelValidatePlan" Then

        SafeHostText = Host.ExcelValidatePlan(arg1)

    End If

    If Err.Number <> 0 Then SafeHostText = ""

    Err.Clear

End Function

Function NormalizeCellText(value)

    Dim text

    text = CStr(value)

    text = Replace(text, ChrW(160), " ")

    text = Replace(text, ChrW(12288), " ")

    NormalizeCellText = Trim(text)

End Function

Function NormalizeKey(value)

    NormalizeKey = LCase(NormalizeCellText(value))

End Function

Function FindHeaderColumn(sourceRange, fieldName)

    Dim colIndex

    FindHeaderColumn = 0

    For colIndex = 1 To sourceRange.Columns.Count

        If StrComp(NormalizeCellText(sourceRange.Cells(1, colIndex).Text), NormalizeCellText(fieldName), vbTextCompare) = 0 Then

            FindHeaderColumn = colIndex

            Exit Function

        End If

    Next

End Function

Function SplitCsvFields(text)

    Dim cleaned, parts, i, n, values()

    cleaned = Replace(Replace(Trim(CStr(text)), "，", ","), "；", ",")

    If Len(cleaned) = 0 Then

        SplitCsvFields = Array()

        Exit Function

    End If

    parts = Split(cleaned, ",")

    n = -1

    ReDim values(UBound(parts))

    For i = 0 To UBound(parts)

        If Len(Trim(parts(i))) > 0 Then

            n = n + 1

            values(n) = Trim(parts(i))

        End If

    Next

    If n < 0 Then

        SplitCsvFields = Array()

    Else

        ReDim Preserve values(n)

        SplitCsvFields = values

    End If

End Function

Function ResolveHeaderColumns(sourceRange, fieldNames)

    Dim i, colIndex, n, cols()

    ResolveHeaderColumns = Array()

    If Not IsArray(fieldNames) Then Exit Function

    On Error Resume Next

    If UBound(fieldNames) < 0 Then Exit Function

    Err.Clear

    n = -1

    ReDim cols(UBound(fieldNames))

    For i = 0 To UBound(fieldNames)

        colIndex = FindHeaderColumn(sourceRange, fieldNames(i))

        If colIndex <= 0 Then

            ResolveHeaderColumns = Array()

            Exit Function

        End If

        n = n + 1

        cols(n) = colIndex

    Next

    If n >= 0 Then

        ReDim Preserve cols(n)

        ResolveHeaderColumns = cols

    End If

End Function

Function UniqueSheetName(workbook, baseName)

    Dim candidate, index, exists, ws

    index = 1

    Do

        If index = 1 Then

            candidate = Left(CStr(baseName), 31)

        Else

            candidate = Left(CStr(baseName), 28) & CStr(index)

        End If

        exists = False

        For Each ws In workbook.Worksheets

            If StrComp(ws.Name, candidate, vbTextCompare) = 0 Then exists = True

        Next

        If Not exists Then

            UniqueSheetName = candidate

            Exit Function

        End If

        index = index + 1

    Loop

End Function

Function BuildTimestamp()

    Dim d

    d = Now

    BuildTimestamp = Year(d) & Right("0" & Month(d), 2) & Right("0" & Day(d), 2) & "_" & Right("0" & Hour(d), 2) & Right("0" & Minute(d), 2) & Right("0" & Second(d), 2)

End Function

Function BuildOutputSheetName(sourceSheetName, functionName)

    Dim baseName

    baseName = CStr(sourceSheetName) & "_" & CStr(functionName) & "_" & BuildTimestamp()

    baseName = Replace(baseName, "[", "(")

    baseName = Replace(baseName, "]", ")")

    baseName = Replace(baseName, ":", "-")

    baseName = Replace(baseName, "\\", "-")

    baseName = Replace(baseName, "/", "-")

    baseName = Replace(baseName, "?", "")

    baseName = Replace(baseName, "*", "")

    If Len(baseName) > 31 Then baseName = Left(baseName, 31)

    BuildOutputSheetName = baseName

End Function

Function CreateOutputSheet(appObj, sourceSheetName, functionName, outputMode)

    On Error Resume Next

    Dim hostJson, sheetName, workbook, sheetObj, modeText

    Set CreateOutputSheet = Nothing

    modeText = NormalizeOutputMode(outputMode)

    hostJson = SafeHostText("ExcelCreateOutputSheet", functionName, modeText, sourceSheetName, True)

    sheetName = ExtractJsonString(hostJson, "sheetName")

    If Len(sheetName) = 0 Then sheetName = ExtractJsonString(hostJson, "name")

    If Len(sheetName) > 0 Then

        Set sheetObj = Nothing

        Set sheetObj = appObj.ActiveWorkbook.Worksheets(sheetName)

        If Not sheetObj Is Nothing Then

            Set CreateOutputSheet = sheetObj

            Exit Function

        End If

    End If

    Set workbook = appObj.ActiveWorkbook

    If modeText = "NewWorkbook" Then

        Set workbook = appObj.Workbooks.Add

    End If

    sheetName = UniqueSheetName(workbook, BuildOutputSheetName(sourceSheetName, functionName))

    Set sheetObj = workbook.Worksheets.Add

    sheetObj.Name = sheetName

    If Err.Number <> 0 Then

        Err.Clear

        Set CreateOutputSheet = Nothing

        Exit Function

    End If

    Set CreateOutputSheet = sheetObj

End Function

Sub SafeDeleteSheet(appObj, sheetObj)

    On Error Resume Next

    Dim prev

    If sheetObj Is Nothing Then Exit Sub

    prev = appObj.DisplayAlerts

    appObj.DisplayAlerts = False

    sheetObj.Delete

    appObj.DisplayAlerts = prev

    Err.Clear

End Sub

Function NormalizeOutputMode(text)

    Dim t

    t = LCase(Trim(CStr(text)))

    If t = "newworkbook" Or t = "workbook" Or t = "新工作簿" Then

        NormalizeOutputMode = "NewWorkbook"

    Else

        NormalizeOutputMode = "NewSheet"

    End If

End Function

Sub WriteSummaryFallback(sheetObj, summaryLines)

    Dim lines, i

    lines = Split(CStr(summaryLines), vbCrLf)

    sheetObj.Cells(1, 1).Value = "项目"

    sheetObj.Cells(1, 2).Value = "内容"

    For i = 0 To UBound(lines)

        sheetObj.Cells(i + 2, 1).Value = "行" & CStr(i + 1)

        sheetObj.Cells(i + 2, 2).Value = lines(i)

    Next

    sheetObj.Rows(1).Font.Bold = True

    sheetObj.Columns.AutoFit

End Sub

Sub RollbackCreatedSheets(appObj, createdSheets, createdCount)

    Dim i

    For i = createdCount To 1 Step -1

        SafeDeleteSheet appObj, createdSheets(i)

        Set createdSheets(i) = Nothing

    Next

End Sub

Function IsBlankCell(value)

    IsBlankCell = (Len(NormalizeCellText(value)) = 0)

End Function

Function FindWorksheet(workbook, sheetName)

    On Error Resume Next

    Dim ws

    Set FindWorksheet = Nothing

    For Each ws In workbook.Worksheets

        If StrComp(ws.Name, Trim(CStr(sheetName)), vbTextCompare) = 0 Then

            Set FindWorksheet = ws

            Exit Function

        End If

    Next

End Function

Function FindOtherSheetName(workbook, currentName)

    Dim ws

    FindOtherSheetName = ""

    For Each ws In workbook.Worksheets

        If StrComp(ws.Name, currentName, vbTextCompare) <> 0 Then

            FindOtherSheetName = ws.Name

            Exit Function

        End If

    Next

End Function

Function FindKeyRow(dataRange, keyColumn, keyText)

    Dim rowIndex

    FindKeyRow = 0

    For rowIndex = 2 To dataRange.Rows.Count

        If NormalizeKey(dataRange.Cells(rowIndex, keyColumn).Text) = keyText Then

            FindKeyRow = rowIndex

            Exit Function

        End If

    Next

End Function

Function CountKeyOccurrences(dataRange, keyColumn, keyText)

    Dim rowIndex, total

    total = 0

    For rowIndex = 2 To dataRange.Rows.Count

        If NormalizeKey(dataRange.Cells(rowIndex, keyColumn).Text) = keyText Then total = total + 1

    Next

    CountKeyOccurrences = total

End Function

Function TargetAbsoluteRow(targetRange, relativeRow)

    If relativeRow > 0 Then TargetAbsoluteRow = targetRange.Row + relativeRow - 1 Else TargetAbsoluteRow = 0

End Function
