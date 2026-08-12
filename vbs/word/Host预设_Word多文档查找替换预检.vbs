' 函数名: HostWordMultiDocFindReplacePreflight

' 描述: 选择文件夹批量只读预检 Word 文档查找命中后，确认词“替换”才逐个打开写回并保存；输出文件级命中摘要，无确认不写

' 适用应用: Word

' 搜索范围: 指定范围

' 搜索对象: 文本

Option Explicit

Const wdFindContinue = 1

Const wdReplaceAll = 2

Const wdMainTextStory = 1

Function Main()

    On Error Resume Next

    Dim appObj

    Set appObj = Host.GetApplication()

    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then

        Err.Clear

        Main = FailureJson("E_NO_WORD_APP", "未取得 Word 应用，请在 Word 中运行该预设")

        Exit Function

    End If

    Main = HostWordMultiDocFindReplacePreflight(appObj)

End Function

Function HostWordMultiDocFindReplacePreflight(appObj)

    On Error Resume Next

    Dim folderPath

    folderPath = Trim(SafeSelectFolder("选择包含待查找替换 Word 文档的文件夹（doc/docx/docm）"))

    If Len(folderPath) = 0 Then

        HostWordMultiDocFindReplacePreflight = FailureJson("E_FOLDER_REQUIRED", "未选择文件夹；未修改任何文档")

        Exit Function

    End If

    Dim findText, replaceText

    findText = SafePrompt("请输入要查找的文本", "")

    If Len(Trim(findText)) = 0 Then

        HostWordMultiDocFindReplacePreflight = FailureJson("E_EMPTY_FIND", "查找文本为空；未修改任何文档")

        Exit Function

    End If

    replaceText = SafePrompt("请输入替换后的文本（可留空表示删除）", "")

    Dim paths(), pathCount, i

    pathCount = 0

    CollectPattern folderPath, "*.doc", paths, pathCount

    CollectPattern folderPath, "*.docx", paths, pathCount

    CollectPattern folderPath, "*.docm", paths, pathCount

    Dim filePath, fileName, sourcePaths(), sourceNames(), sourceCount, skippedTemp

    sourceCount = 0

    skippedTemp = 0

    For i = 0 To pathCount - 1

        filePath = paths(i)

        fileName = FileNameFromPath(filePath)

        If Left(fileName, 2) = "~$" Then

            skippedTemp = skippedTemp + 1

        Else

            ReDim Preserve sourcePaths(sourceCount)

            ReDim Preserve sourceNames(sourceCount)

            sourcePaths(sourceCount) = filePath

            sourceNames(sourceCount) = fileName

            sourceCount = sourceCount + 1

        End If

    Next

    If sourceCount = 0 Then

        HostWordMultiDocFindReplacePreflight = FailureJson("E_NO_FILES", "文件夹内未找到可处理的 Word 文档")

        Exit Function

    End If

    Dim hitCounts(), openFlags(), sampleLines, totalHits, openedCount, failedOpen, hitFiles

    ReDim hitCounts(sourceCount - 1)

    ReDim openFlags(sourceCount - 1)

    sampleLines = ""

    totalHits = 0

    openedCount = 0

    failedOpen = 0

    hitFiles = 0

    Dim doc, hitCount, lineCount

    lineCount = 0

    For i = 0 To sourceCount - 1

        Set doc = Nothing

        Err.Clear

        Set doc = appObj.Documents.Open(sourcePaths(i), False, True)

        If Err.Number <> 0 Or TypeName(doc) = "Empty" Or TypeName(doc) = "Nothing" Then

            Err.Clear

            openFlags(i) = False

            hitCounts(i) = 0

            failedOpen = failedOpen + 1

        Else

            openFlags(i) = True

            openedCount = openedCount + 1

            hitCount = CountOccurrences(SafeDocText(doc), findText)

            hitCounts(i) = hitCount

            totalHits = totalHits + hitCount

            If hitCount > 0 Then hitFiles = hitFiles + 1

            If lineCount < 8 Then

                If Len(sampleLines) > 0 Then sampleLines = sampleLines & vbCrLf

                sampleLines = sampleLines & sourceNames(i) & " => 命中 " & CStr(hitCount)

                lineCount = lineCount + 1

            End If

            doc.Close False

            Err.Clear

        End If

    Next

    Dim budgetJson, planId, planPreview, previewText

    budgetJson = SafeRunBudget(sourceCount)

    previewText = "Word 多文档查找替换预检（尚未写入）" & vbCrLf & _
        "文件夹=" & folderPath & vbCrLf & _
        "查找=" & findText & "；替换为=" & replaceText & vbCrLf & _
        "候选文件=" & CStr(sourceCount) & "；只读打开成功=" & CStr(openedCount) & "；打开失败=" & CStr(failedOpen) & "；跳过临时=" & CStr(skippedTemp) & vbCrLf & _
        "命中文件=" & CStr(hitFiles) & "；命中合计=" & CStr(totalHits) & vbCrLf & _
        "样例：" & vbCrLf & sampleLines & vbCrLf & _
        "确认前不会写回任何文档" & vbCrLf & _
        "Host.GetRunBudgetPlan: " & budgetJson

    planId = SafeBeginWritePlan()

    SafeRecordWrite "word_multi_doc_find_replace_preflight", "{""folder"":""" & EscapeJson(folderPath) & """,""files"":" & CStr(sourceCount) & ",""hits"":" & CStr(totalHits) & "}", "office.word.multiDocFindReplace"

    planPreview = SafePreviewWritePlan()

    SafeRollbackWritePlan planId

    If StrComp(Trim(SafePrompt(previewText & vbCrLf & vbCrLf & "确认执行多文档替换，请输入：替换", "")), "替换", vbTextCompare) <> 0 Then

        HostWordMultiDocFindReplacePreflight = FailureJson("E_CONFIRM_REQUIRED", "未输入“替换”，已取消且未修改任何文档")

        Exit Function

    End If

    Dim replacedFiles, replacedHits, failedWrite, writeSummary, changed

    replacedFiles = 0

    replacedHits = 0

    failedWrite = 0

    writeSummary = ""

    For i = 0 To sourceCount - 1

        If openFlags(i) And hitCounts(i) > 0 Then

            Set doc = Nothing

            Err.Clear

            Set doc = appObj.Documents.Open(sourcePaths(i), False, False)

            If Err.Number <> 0 Or TypeName(doc) = "Empty" Or TypeName(doc) = "Nothing" Then

                Err.Clear

                failedWrite = failedWrite + 1

            Else

                changed = ReplaceInDocument(doc, findText, replaceText)

                If changed Then

                    doc.Save

                    If Err.Number <> 0 Then

                        Err.Clear

                        failedWrite = failedWrite + 1

                    Else

                        replacedFiles = replacedFiles + 1

                        replacedHits = replacedHits + hitCounts(i)

                        If Len(writeSummary) > 0 Then writeSummary = writeSummary & "; "

                        writeSummary = writeSummary & sourceNames(i) & "(" & CStr(hitCounts(i)) & ")"

                    End If

                End If

                doc.Close False

                Err.Clear

            End If

        End If

    Next

    Dim summary

    summary = "Word 多文档查找替换完成；替换文件=" & CStr(replacedFiles) & "；估算命中写回=" & CStr(replacedHits) & "；写失败=" & CStr(failedWrite) & "；预检命中文件=" & CStr(hitFiles)

    If Len(writeSummary) > 0 Then summary = summary & "；明细=" & writeSummary

    Host.WriteClipboard summary

    SafeWriteLog summary

    HostWordMultiDocFindReplacePreflight = "{""ok"":true,""sourceUnchangedBeforeConfirm"":true,""folder"":""" & EscapeJson(folderPath) & _
        """,""candidateFiles"":" & CStr(sourceCount) & ",""openedReadonly"":" & CStr(openedCount) & ",""failedOpen"":" & CStr(failedOpen) & _
        ",""hitFiles"":" & CStr(hitFiles) & ",""totalHits"":" & CStr(totalHits) & ",""replacedFiles"":" & CStr(replacedFiles) & _
        ",""replacedHits"":" & CStr(replacedHits) & ",""failedWrite"":" & CStr(failedWrite) & ",""writePlanPreview"":""" & EscapeJson(planPreview) & _
        """,""message"":""" & EscapeJson(summary) & """}"

End Function

Function SafeDocText(doc)

    On Error Resume Next

    SafeDocText = ""

    SafeDocText = CStr(doc.Content.Text)

    If Err.Number <> 0 Then

        SafeDocText = ""

        Err.Clear

    End If

End Function

Function CountOccurrences(text, needle)

    Dim count, startAt, pos

    count = 0

    startAt = 1

    If Len(needle) = 0 Then

        CountOccurrences = 0

        Exit Function

    End If

    Do

        pos = InStr(startAt, text, needle, vbTextCompare)

        If pos <= 0 Then Exit Do

        count = count + 1

        startAt = pos + Len(needle)

    Loop

    CountOccurrences = count

End Function

Function ReplaceInDocument(doc, findText, replaceText)

    On Error Resume Next

    Dim changed, rng, res

    changed = False

    Set rng = doc.StoryRanges(wdMainTextStory)

    Do While Not (TypeName(rng) = "Empty" Or TypeName(rng) = "Nothing")

        res = rng.Find.Execute(findText, False, False, False, False, False, True, wdFindContinue, False, replaceText, wdReplaceAll)

        If res Then changed = True

        Set rng = rng.NextStoryRange

        If Err.Number <> 0 Then

            Err.Clear

            Exit Do

        End If

    Loop

    ReplaceInDocument = changed

End Function

Function SafeRunBudget(itemCount)

    On Error Resume Next

    SafeRunBudget = Host.GetRunBudgetPlan("word_multi_doc_find_replace", itemCount, 200, 10, 180)

    If Err.Number <> 0 Then SafeRunBudget = "{}"

    Err.Clear

End Function

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

Function JsonUnescape(value)

    Dim text, i, ch, outText

    text = CStr(value)

    outText = ""

    i = 1

    Do While i <= Len(text)

        ch = Mid(text, i, 1)

        If ch = "\" And i < Len(text) Then

            i = i + 1

            ch = Mid(text, i, 1)

            If ch = "n" Then

                outText = outText & vbLf

            ElseIf ch = "r" Then

                outText = outText & vbCr

            ElseIf ch = "t" Then

                outText = outText & vbTab

            Else

                outText = outText & ch

            End If

        Else

            outText = outText & ch

        End If

        i = i + 1

    Loop

    JsonUnescape = outText

End Function

Function FileNameFromPath(pathValue)

    Dim text, p

    text = CStr(pathValue)

    p = InStrRev(text, "\")

    If p <= 0 Then p = InStrRev(text, "/")

    If p > 0 Then FileNameFromPath = Mid(text, p + 1) Else FileNameFromPath = text

End Function

Function BaseNameFromFileName(fileName)

    Dim text, p

    text = CStr(fileName)

    p = InStrRev(text, ".")

    If p > 1 Then BaseNameFromFileName = Left(text, p - 1) Else BaseNameFromFileName = text

End Function

Function UniqueOutputPath(folderPath, baseName, extension)

    Dim candidate, idx, fullPath

    candidate = Host.SanitizeFileName(baseName, "part")

    If Len(candidate) = 0 Then candidate = "part"

    fullPath = Host.CombinePath(folderPath, candidate & extension)

    idx = 1

    Do While Host.PathExists(fullPath)

        idx = idx + 1

        fullPath = Host.CombinePath(folderPath, candidate & "_" & CStr(idx) & extension)

    Loop

    UniqueOutputPath = fullPath

End Function
