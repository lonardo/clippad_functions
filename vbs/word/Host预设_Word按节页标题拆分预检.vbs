' 函数名: HostWordSplitBySectionPageHeadingPreflight

' 描述: 按节、分页符或标题1预览分段数量与样例后，确认词“生成”导出多个 docx 副本到所选文件夹；原文档不拆坏

' 适用应用: Word

' 搜索范围: 全文

' 搜索对象: 无

Option Explicit

Const wdGoToPage = 1

Const wdGoToAbsolute = 1

Const wdStatisticPages = 2

Const wdFormatXMLDocument = 12

Function Main()

    On Error Resume Next

    Dim appObj

    Set appObj = Host.GetApplication()

    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then

        Err.Clear

        Main = FailureJson("E_NO_WORD_APP", "未取得 Word 应用，请在 Word 中运行该预设")

        Exit Function

    End If

    Main = HostWordSplitBySectionPageHeadingPreflight(appObj)

End Function

Function HostWordSplitBySectionPageHeadingPreflight(appObj)

    On Error Resume Next

    Dim doc

    Set doc = appObj.ActiveDocument

    If Err.Number <> 0 Or TypeName(doc) = "Empty" Or TypeName(doc) = "Nothing" Then

        Err.Clear

        HostWordSplitBySectionPageHeadingPreflight = FailureJson("E_NO_DOCUMENT", "当前没有活动 Word 文档")

        Exit Function

    End If

    Dim modeText, modeCode

    modeText = Trim(SafePrompt("拆分方式：节 / 页 / 标题1", "节"))

    modeCode = NormalizeSplitMode(modeText)

    If Len(modeCode) = 0 Then

        HostWordSplitBySectionPageHeadingPreflight = FailureJson("E_MODE_INVALID", "拆分方式仅支持：节 / 页 / 标题1")

        Exit Function

    End If

    Dim partCount, sampleLines, starts(), ends(), labels(), okBuild

    okBuild = BuildSplitPlan(doc, modeCode, partCount, starts, ends, labels, sampleLines)

    If Not okBuild Then

        HostWordSplitBySectionPageHeadingPreflight = FailureJson("E_SPLIT_PLAN", "无法构建拆分预览：" & sampleLines)

        Exit Function

    End If

    If partCount <= 0 Then

        HostWordSplitBySectionPageHeadingPreflight = FailureJson("E_NO_PARTS", "没有可拆分的段落/节/页")

        Exit Function

    End If

    Dim baseName

    baseName = "未命名文档"

    On Error Resume Next

    If Len(CStr(doc.Name)) > 0 Then baseName = BaseNameNoExt(CStr(doc.Name))

    Err.Clear

    Dim planId, planPreview, previewText

    previewText = "Word 按" & ModeLabel(modeCode) & "拆分预检（尚未写入）" & vbCrLf & _
        "原文档=" & baseName & "；分段数=" & CStr(partCount) & "；输出=所选文件夹多个 docx 副本" & vbCrLf & _
        "原文档不会被拆坏或覆盖" & vbCrLf & _
        "样例：" & vbCrLf & sampleLines

    planId = SafeBeginWritePlan()

    SafeRecordWrite "word_split_section_page_heading_preflight", "{""mode"":""" & EscapeJson(modeCode) & """,""parts"":" & CStr(partCount) & "}", "office.word.splitCopy"

    planPreview = SafePreviewWritePlan()

    SafeRollbackWritePlan planId

    If StrComp(Trim(SafePrompt(previewText & vbCrLf & vbCrLf & "确认导出拆分副本，请输入：生成", "")), "生成", vbTextCompare) <> 0 Then

        HostWordSplitBySectionPageHeadingPreflight = FailureJson("E_CONFIRM_REQUIRED", "未输入“生成”，已取消且未导出文件")

        Exit Function

    End If

    Dim outputFolder

    outputFolder = Trim(SafeSelectFolder("选择拆分副本输出文件夹"))

    If Len(outputFolder) = 0 Then

        HostWordSplitBySectionPageHeadingPreflight = FailureJson("E_FOLDER_REQUIRED", "未选择输出文件夹；未导出文件")

        Exit Function

    End If

    Dim i, okCount, failCount, outPaths, newDoc, rng, outPath, nameTry, saveOk

    okCount = 0

    failCount = 0

    outPaths = ""

    For i = 0 To partCount - 1

        Set rng = Nothing

        Err.Clear

        Set rng = doc.Range(starts(i), ends(i))

        If Err.Number <> 0 Or TypeName(rng) = "Empty" Or TypeName(rng) = "Nothing" Then

            Err.Clear

            failCount = failCount + 1

        Else

            Set newDoc = Nothing

            Err.Clear

            Set newDoc = appObj.Documents.Add

            If Err.Number <> 0 Or TypeName(newDoc) = "Empty" Or TypeName(newDoc) = "Nothing" Then

                Err.Clear

                failCount = failCount + 1

            Else

                Err.Clear

                newDoc.Range(0, 0).FormattedText = rng.FormattedText

                If Err.Number <> 0 Then

                    Err.Clear

                    newDoc.Range(0, 0).Text = rng.Text

                End If

                nameTry = baseName & "_" & labels(i)

                outPath = UniqueOutputPath(outputFolder, nameTry, ".docx")

                saveOk = False

                Err.Clear

                newDoc.SaveAs2 outPath, wdFormatXMLDocument

                If Err.Number = 0 Then

                    saveOk = True

                Else

                    Err.Clear

                    newDoc.SaveAs outPath, wdFormatXMLDocument

                    If Err.Number = 0 Then saveOk = True Else Err.Clear

                End If

                newDoc.Close False

                Err.Clear

                If saveOk Then

                    okCount = okCount + 1

                    If Len(outPaths) > 0 Then outPaths = outPaths & "; "

                    outPaths = outPaths & FileNameOnly(outPath)

                Else

                    failCount = failCount + 1

                End If

            End If

        End If

    Next

    Dim summary

    summary = "Word 拆分副本导出完成；方式=" & ModeLabel(modeCode) & "；成功=" & CStr(okCount) & "；失败=" & CStr(failCount) & "；输出文件夹=" & outputFolder & "；原文档未修改"

    If Len(outPaths) > 0 Then summary = summary & "；文件=" & outPaths

    Host.WriteClipboard summary

    SafeWriteLog summary

    HostWordSplitBySectionPageHeadingPreflight = "{""ok"":true,""sourceUnchanged"":true,""mode"":""" & EscapeJson(modeCode) & _
        """,""partCount"":" & CStr(partCount) & ",""exported"":" & CStr(okCount) & ",""failed"":" & CStr(failCount) & _
        ",""outputFolder"":""" & EscapeJson(outputFolder) & """,""writePlanPreview"":""" & EscapeJson(planPreview) & _
        """,""message"":""" & EscapeJson(summary) & """}"

End Function

Function BuildSplitPlan(doc, modeCode, ByRef partCount, ByRef starts, ByRef ends, ByRef labels, ByRef sampleLines)

    On Error Resume Next

    partCount = 0

    sampleLines = ""

    BuildSplitPlan = False

    If modeCode = "section" Then

        BuildSplitPlan = BuildSectionPlan(doc, partCount, starts, ends, labels, sampleLines)

    ElseIf modeCode = "page" Then

        BuildSplitPlan = BuildPagePlan(doc, partCount, starts, ends, labels, sampleLines)

    ElseIf modeCode = "heading1" Then

        BuildSplitPlan = BuildHeading1Plan(doc, partCount, starts, ends, labels, sampleLines)

    End If

End Function

Function BuildSectionPlan(doc, ByRef partCount, ByRef starts, ByRef ends, ByRef labels, ByRef sampleLines)

    On Error Resume Next

    Dim i, secCount, rng, lineCount, labelText

    secCount = doc.Sections.Count

    If Err.Number <> 0 Or secCount <= 0 Then

        sampleLines = "无法读取分节信息"

        BuildSectionPlan = False

        Exit Function

    End If

    If secCount <= 1 Then

        sampleLines = "文档没有分节符可拆分（仅 1 节）"

        BuildSectionPlan = False

        Exit Function

    End If

    ReDim starts(secCount - 1)

    ReDim ends(secCount - 1)

    ReDim labels(secCount - 1)

    partCount = secCount

    lineCount = 0

    sampleLines = ""

    For i = 1 To secCount

        Set rng = doc.Sections(i).Range

        starts(i - 1) = rng.Start

        ends(i - 1) = rng.End

        labelText = "第" & CStr(i) & "节"

        labels(i - 1) = labelText

        If lineCount < 8 Then

            If Len(sampleLines) > 0 Then sampleLines = sampleLines & vbCrLf

            sampleLines = sampleLines & labelText & "；预览=" & Left(CleanPreviewText(rng.Text), 40)

            lineCount = lineCount + 1

        End If

    Next

    BuildSectionPlan = True

End Function

Function BuildPagePlan(doc, ByRef partCount, ByRef starts, ByRef ends, ByRef labels, ByRef sampleLines)

    On Error Resume Next

    Dim totalPages, i, r1, r2, startPos, endPos, lineCount, labelText

    totalPages = 0

    Err.Clear

    totalPages = doc.ComputeStatistics(wdStatisticPages)

    If Err.Number <> 0 Or totalPages <= 0 Then

        sampleLines = "无法计算页数"

        BuildPagePlan = False

        Exit Function

    End If

    ReDim starts(totalPages - 1)

    ReDim ends(totalPages - 1)

    ReDim labels(totalPages - 1)

    partCount = totalPages

    lineCount = 0

    sampleLines = ""

    For i = 1 To totalPages

        Set r1 = Nothing

        Err.Clear

        Set r1 = doc.GoTo(wdGoToPage, wdGoToAbsolute, i)

        If Err.Number <> 0 Or TypeName(r1) = "Empty" Or TypeName(r1) = "Nothing" Then

            sampleLines = "无法定位第 " & CStr(i) & " 页"

            BuildPagePlan = False

            Exit Function

        End If

        startPos = r1.Start

        If i < totalPages Then

            Set r2 = Nothing

            Err.Clear

            Set r2 = doc.GoTo(wdGoToPage, wdGoToAbsolute, i + 1)

            If Err.Number = 0 And Not (TypeName(r2) = "Empty" Or TypeName(r2) = "Nothing") Then

                endPos = r2.Start

            Else

                endPos = doc.Content.End

                Err.Clear

            End If

        Else

            endPos = doc.Content.End

        End If

        starts(i - 1) = startPos

        ends(i - 1) = endPos

        labelText = "第" & CStr(i) & "页"

        labels(i - 1) = labelText

        If lineCount < 8 Then

            If Len(sampleLines) > 0 Then sampleLines = sampleLines & vbCrLf

            sampleLines = sampleLines & labelText & "；预览=" & Left(CleanPreviewText(doc.Range(startPos, endPos).Text), 40)

            lineCount = lineCount + 1

        End If

    Next

    BuildPagePlan = True

End Function

Function BuildHeading1Plan(doc, ByRef partCount, ByRef starts, ByRef ends, ByRef labels, ByRef sampleLines)

    On Error Resume Next

    Dim para, styleName, count, i, lineCount, labelText, textValue, positions(), titles()

    count = 0

    For Each para In doc.Paragraphs

        styleName = ""

        Err.Clear

        styleName = CStr(para.Style)

        If Err.Number <> 0 Then

            Err.Clear

            styleName = ""

        End If

        If IsHeading1Style(styleName) Then

            ReDim Preserve positions(count)

            ReDim Preserve titles(count)

            positions(count) = para.Range.Start

            textValue = CleanPreviewText(para.Range.Text)

            If Len(textValue) = 0 Then textValue = "标题" & CStr(count + 1)

            titles(count) = textValue

            count = count + 1

        End If

    Next

    If count = 0 Then

        sampleLines = "未找到标题1样式段落"

        BuildHeading1Plan = False

        Exit Function

    End If

    ReDim starts(count - 1)

    ReDim ends(count - 1)

    ReDim labels(count - 1)

    partCount = count

    lineCount = 0

    sampleLines = ""

    For i = 0 To count - 1

        starts(i) = positions(i)

        If i < count - 1 Then

            ends(i) = positions(i + 1)

        Else

            ends(i) = doc.Content.End

        End If

        labelText = "标题1_" & CStr(i + 1) & "_" & Left(titles(i), 20)

        labels(i) = labelText

        If lineCount < 8 Then

            If Len(sampleLines) > 0 Then sampleLines = sampleLines & vbCrLf

            sampleLines = sampleLines & "分段" & CStr(i + 1) & "=" & titles(i)

            lineCount = lineCount + 1

        End If

    Next

    BuildHeading1Plan = True

End Function

Function IsHeading1Style(styleName)

    Dim text

    text = LCase(Trim(CStr(styleName)))

    If text = "heading 1" Or text = "标题 1" Or text = "标题1" Then

        IsHeading1Style = True

    ElseIf InStr(1, text, "heading 1", vbTextCompare) > 0 Then

        IsHeading1Style = True

    ElseIf InStr(1, text, "标题 1", vbTextCompare) > 0 Or InStr(1, text, "标题1", vbTextCompare) > 0 Then

        IsHeading1Style = True

    Else

        IsHeading1Style = False

    End If

End Function

Function NormalizeSplitMode(value)

    Dim text

    text = LCase(Trim(CStr(value)))

    If text = "节" Or text = "section" Or text = "分节" Then

        NormalizeSplitMode = "section"

    ElseIf text = "页" Or text = "page" Or text = "分页符" Or text = "分页" Then

        NormalizeSplitMode = "page"

    ElseIf text = "标题1" Or text = "标题 1" Or text = "heading1" Or text = "h1" Then

        NormalizeSplitMode = "heading1"

    Else

        NormalizeSplitMode = ""

    End If

End Function

Function ModeLabel(modeCode)

    If modeCode = "section" Then

        ModeLabel = "节"

    ElseIf modeCode = "page" Then

        ModeLabel = "页"

    ElseIf modeCode = "heading1" Then

        ModeLabel = "标题1"

    Else

        ModeLabel = modeCode

    End If

End Function

Function CleanPreviewText(value)

    Dim text

    text = CStr(value)

    text = Replace(text, vbCr, " ")

    text = Replace(text, vbLf, " ")

    text = Replace(text, vbTab, " ")

    text = Replace(text, ChrW(7), "")

    text = Trim(text)

    CleanPreviewText = text

End Function

Function BaseNameNoExt(fileName)

    Dim text, p

    text = CStr(fileName)

    p = InStrRev(text, ".")

    If p > 1 Then BaseNameNoExt = Left(text, p - 1) Else BaseNameNoExt = text

End Function

Function FileNameOnly(pathValue)

    Dim text, p

    text = CStr(pathValue)

    p = InStrRev(text, "\")

    If p <= 0 Then p = InStrRev(text, "/")

    If p > 0 Then FileNameOnly = Mid(text, p + 1) Else FileNameOnly = text

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
