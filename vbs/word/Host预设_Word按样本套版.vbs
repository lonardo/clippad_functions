' 函数名: HostWordApplySampleFormatting
' 描述: 选择 Word 样本文档后，先预览标题、层级标题、正文、表格、页面及页眉页脚格式的影响范围；仅对已保存且无未保存修改的文档创建备份并经二次确认后写入，失败时重载原文件
' 适用应用: Word
' 搜索范围: 全文
' 搜索对象: 无

Option Explicit

Const wdNoProtection = -1
Const wdOutlineLevelBodyText = 10

Function Main()
    On Error Resume Next

    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = "{""ok"":false,""code"":""E_NO_WORD_APP"",""message"":""未取得 Word 应用，请在 Word 中运行该预设""}"
        Exit Function
    End If

    Main = HostWordApplySampleFormatting(appObj)
End Function

Function HostWordApplySampleFormatting(appObj)
    On Error Resume Next

    Dim targetDoc
    Set targetDoc = appObj.ActiveDocument
    If Err.Number <> 0 Or TypeName(targetDoc) = "Empty" Or TypeName(targetDoc) = "Nothing" Then
        Err.Clear
        HostWordApplySampleFormatting = FailureJson("E_NO_DOCUMENT", "当前没有活动 Word 文档")
        Exit Function
    End If

    If IsDocumentProtected(targetDoc) Then
        HostWordApplySampleFormatting = FailureJson("E_DOCUMENT_PROTECTED", "当前文档受保护，已停止套版且未修改文档")
        Exit Function
    End If
    If IsTrackingRevisions(targetDoc) Then
        HostWordApplySampleFormatting = FailureJson("E_REVISIONS_ACTIVE", "当前文档正在记录修订，已停止套版；请先确认修订处理策略")
        Exit Function
    End If

    Dim targetSelectionStart, targetSelectionEnd
    targetSelectionStart = -1
    targetSelectionEnd = -1
    CaptureTargetSelection appObj, targetDoc, targetSelectionStart, targetSelectionEnd

    Dim selectedPath, samplePath, profileId, profileVersion, profileDisplayName, profileText, sourceJson, runtimeJson, guardrailsJson
    selectedPath = Trim(SafeSelectFile("选择规范 Word 样本文档或已确认的 JSON 格式档案", "Word 文档或格式档案|*.doc;*.docx;*.docm;*.json"))
    If Len(selectedPath) = 0 Then
        HostWordApplySampleFormatting = FailureJson("E_SAMPLE_REQUIRED", "未选择样本文档；未修改当前文档")
        Exit Function
    End If
    samplePath = selectedPath
    profileId = ""
    profileVersion = ""
    profileDisplayName = ""
    If IsJsonProfilePath(selectedPath) Then
        profileText = SafeReadTextFile(selectedPath)
        If Len(profileText) = 0 Then
            HostWordApplySampleFormatting = FailureJson("E_PROFILE_READ", "无法读取格式档案：" & selectedPath)
            Exit Function
        End If
        sourceJson = ExtractJsonObject(profileText, "source")
        runtimeJson = ExtractJsonObject(profileText, "runtime")
        guardrailsJson = ExtractJsonObject(profileText, "guardrails")
        profileId = ExtractJsonString(profileText, "id")
        profileVersion = ExtractJsonString(profileText, "profileVersion")
        profileDisplayName = ExtractJsonString(profileText, "displayName")
        If LCase(ExtractJsonString(profileText, "status")) <> "user_confirmed" Then
            HostWordApplySampleFormatting = FailureJson("E_PROFILE_UNCONFIRMED", "格式档案尚未确认来源、版本和适用范围，已拒绝套版")
            Exit Function
        End If
        If Len(sourceJson) = 0 Or Len(runtimeJson) = 0 Or Len(guardrailsJson) = 0 Or Len(profileId) = 0 Or Len(profileVersion) = 0 Then
            HostWordApplySampleFormatting = FailureJson("E_PROFILE_INVALID", "格式档案不符合运行时契约，缺少来源、样本、保护项或身份信息")
            Exit Function
        End If
        If Not ExtractJsonBoolean(sourceJson, "confirmedByUser") Then
            HostWordApplySampleFormatting = FailureJson("E_PROFILE_UNCONFIRMED", "格式档案尚未确认来源、版本和适用范围，已拒绝套版")
            Exit Function
        End If
        If Not ExtractJsonBoolean(guardrailsJson, "preserveBusinessContent") Or Not ExtractJsonBoolean(guardrailsJson, "requiresPreview") Or Not ExtractJsonBoolean(guardrailsJson, "requiresConfirmation") Then
            HostWordApplySampleFormatting = FailureJson("E_PROFILE_GUARDRAIL", "格式档案未声明业务内容保护、预览和确认门禁，已拒绝套版")
            Exit Function
        End If
        samplePath = ExtractJsonString(runtimeJson, "samplePath")
        If Len(samplePath) = 0 Or Not SafePathExists(samplePath) Then
            HostWordApplySampleFormatting = FailureJson("E_PROFILE_SAMPLE_MISSING", "格式档案引用的样本文档不存在或路径已失效")
            Exit Function
        End If
    End If

    Dim sampleDoc
    Set sampleDoc = Nothing
    Err.Clear
    Set sampleDoc = appObj.Documents.Open(samplePath, False, True)
    If Err.Number <> 0 Or TypeName(sampleDoc) = "Empty" Or TypeName(sampleDoc) = "Nothing" Then
        Dim openError
        openError = Err.Description
        Err.Clear
        HostWordApplySampleFormatting = FailureJson("E_SAMPLE_OPEN", "无法以只读方式打开样本文档：" & openError)
        Exit Function
    End If

    Dim scopeMode, selectedRules, runMode
    runMode = NormalizeRunMode(SafePrompt("请选择运行模式：预览 或 应用。默认只预览，不修改文档", "预览"))
    scopeMode = NormalizeScopeMode(SafePrompt("请选择处理范围：全文 或 选区", "全文"))
    selectedRules = NormalizeRules(SafePrompt("选择格式规则，可输入：标题,层级标题,正文,表格,页面,页眉页脚；默认全部", "标题,层级标题,正文,表格,页面,页眉页脚"))

    Dim scopeRange
    Set scopeRange = GetTargetRange(targetDoc, scopeMode, targetSelectionStart, targetSelectionEnd)
    If TypeName(scopeRange) = "Empty" Or TypeName(scopeRange) = "Nothing" Then
        SafeCloseDocument sampleDoc
        HostWordApplySampleFormatting = FailureJson("E_SCOPE_UNAVAILABLE", "无法取得目标范围；未修改当前文档")
        Exit Function
    End If
    If scopeMode = "selection" And SafeLong(scopeRange.End) <= SafeLong(scopeRange.Start) Then
        SafeCloseDocument sampleDoc
        HostWordApplySampleFormatting = FailureJson("E_EMPTY_SELECTION", "当前选区为空；未修改当前文档")
        Exit Function
    End If

    Dim sampleTitle, sampleHeading, sampleBody, sampleTable, sampleSection
    Set sampleTitle = FindFirstNonEmptyParagraph(sampleDoc.Content)
    Set sampleHeading = FindFirstHeadingParagraph(sampleDoc.Content)
    Set sampleBody = FindFirstBodyParagraph(sampleDoc.Content, sampleTitle)
    Set sampleTable = FindFirstTable(sampleDoc)
    Set sampleSection = FindFirstSection(sampleDoc)

    Dim titleCount, headingCount, bodyCount, tableCount, sectionCount, unavailableRules
    unavailableRules = ""
    CountImpact scopeRange, titleCount, headingCount, bodyCount, tableCount
    sectionCount = CountTargetSections(targetDoc, scopeMode)
    If HasRule(selectedRules, "标题") And IsNothing(sampleTitle) Then unavailableRules = AppendItem(unavailableRules, "标题")
    If HasRule(selectedRules, "层级标题") And IsNothing(sampleHeading) Then unavailableRules = AppendItem(unavailableRules, "层级标题")
    If HasRule(selectedRules, "正文") And IsNothing(sampleBody) Then unavailableRules = AppendItem(unavailableRules, "正文")
    If HasRule(selectedRules, "表格") And IsNothing(sampleTable) Then unavailableRules = AppendItem(unavailableRules, "表格")
    If (HasRule(selectedRules, "页面") Or HasRule(selectedRules, "页眉页脚")) And IsNothing(sampleSection) Then unavailableRules = AppendItem(unavailableRules, "页面/页眉页脚")

    Dim previewText, planId, planPreview, planValidation
    previewText = BuildPreviewText(sampleDoc.Name, scopeMode, selectedRules, titleCount, headingCount, bodyCount, tableCount, sectionCount, unavailableRules)
    If Len(profileId) > 0 Then
        previewText = "格式档案：" & profileDisplayName & "（" & profileId & "，版本 " & profileVersion & "）" & vbCrLf & previewText
    End If
    planId = SafeBeginWritePlan()
    SafeRecordWrite "word_sample_format_transfer", BuildPlanParams(samplePath, scopeMode, selectedRules, titleCount, headingCount, bodyCount, tableCount, sectionCount), "office.word.format.sampleTransfer"
    planPreview = SafePreviewWritePlan()
    planValidation = SafeValidateWritePlan(planId)
    SafeRollbackWritePlan planId

    If runMode <> "apply" Then
        Host.WriteClipboard previewText
        SafeWriteLog previewText
        SafeCloseDocument sampleDoc
        HostWordApplySampleFormatting = SuccessPreviewJson(samplePath, profileId, profileVersion, scopeMode, selectedRules, titleCount, headingCount, bodyCount, tableCount, sectionCount, unavailableRules, planPreview, planValidation, previewText)
        Exit Function
    End If

    Dim confirmation
    confirmation = Trim(SafePrompt(previewText & vbCrLf & vbCrLf & "确认仅修改格式并保留业务文字，请输入：应用", ""))
    If StrComp(confirmation, "应用", vbTextCompare) <> 0 Then
        SafeCloseDocument sampleDoc
        HostWordApplySampleFormatting = FailureJson("E_CONFIRM_REQUIRED", "未输入“应用”，已取消且未修改文档")
        Exit Function
    End If

    Dim targetPath
    targetPath = SafeDocumentPath(targetDoc)
    If Len(targetPath) = 0 Or Not IsDocumentClean(targetDoc) Then
        SafeCloseDocument sampleDoc
        HostWordApplySampleFormatting = FailureJson("E_SAVE_REQUIRED", "为保证可恢复，请先保存当前文档及已有修改，再重新运行套版")
        Exit Function
    End If

    Dim backupPath
    backupPath = SafeCreateBackup("word_sample_format_transfer")
    If Len(backupPath) = 0 Then
        SafeCloseDocument sampleDoc
        HostWordApplySampleFormatting = FailureJson("E_BACKUP_REQUIRED", "未能创建执行前备份，已停止且未修改文档")
        Exit Function
    End If

    Dim originalContent
    originalContent = CStr(targetDoc.Content.Text)
    On Error Resume Next
    targetDoc.Activate
    Err.Clear

    Dim changedTitle, changedHeading, changedBody, changedTables, changedSections, failureCount
    changedTitle = 0
    changedHeading = 0
    changedBody = 0
    changedTables = 0
    changedSections = 0
    failureCount = 0

    ApplyParagraphRules scopeRange, sampleTitle, sampleHeading, sampleBody, selectedRules, changedTitle, changedHeading, changedBody, failureCount
    If HasRule(selectedRules, "表格") And Not IsNothing(sampleTable) Then
        ApplyTableRules scopeRange, sampleTable, changedTables, failureCount
    End If
    If HasRule(selectedRules, "页面") Or HasRule(selectedRules, "页眉页脚") Then
        ApplySectionRules targetDoc, sampleSection, scopeMode, scopeRange, selectedRules, changedSections, failureCount
    End If

    Dim contentPreserved
    contentPreserved = (CStr(targetDoc.Content.Text) = originalContent)
    If failureCount > 0 Or Not contentPreserved Then
        SafeCloseDocument sampleDoc
        Dim reloadOk
        reloadOk = ReloadOriginalDocument(appObj, targetDoc, targetPath)
        HostWordApplySampleFormatting = FailureWithRollbackJson("E_APPLY_ROLLED_BACK", "套版出现失败或检测到业务文字变化，已放弃未保存结果并重新加载原文件", reloadOk, failureCount, contentPreserved, backupPath)
        Exit Function
    End If

    SafeCloseDocument sampleDoc

    Dim summary
    summary = "Word 按样本套版完成；范围=" & ScopeLabel(scopeMode) & "；标题=" & CStr(changedTitle) & _
        "；层级标题=" & CStr(changedHeading) & "；正文=" & CStr(changedBody) & _
        "；表格=" & CStr(changedTables) & "；节=" & CStr(changedSections) & _
        "；业务文字保持=true；本次结果尚未保存，可关闭不保存或使用备份恢复"
    If Len(backupPath) > 0 Then summary = summary & "；备份=" & backupPath
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostWordApplySampleFormatting = "{""ok"":true,""applied"":true,""contentPreserved"":true,""profileId"":""" & EscapeJson(profileId) & _
        """,""profileVersion"":""" & EscapeJson(profileVersion) & """,""scope"":""" & scopeMode & _
        """,""changed"":{""title"":" & CStr(changedTitle) & ",""heading"":" & CStr(changedHeading) & _
        ",""body"":" & CStr(changedBody) & ",""tables"":" & CStr(changedTables) & ",""sections"":" & CStr(changedSections) & _
        "},""backupPath"":""" & EscapeJson(backupPath) & """,""message"":""" & EscapeJson(summary) & """}"
End Function

Sub ApplyParagraphRules(scopeRange, sampleTitle, sampleHeading, sampleBody, selectedRules, ByRef changedTitle, ByRef changedHeading, ByRef changedBody, ByRef failureCount)
    Dim para, textValue, nonEmptyIndex
    nonEmptyIndex = 0
    For Each para In scopeRange.Paragraphs
        textValue = CleanParagraphText(para.Range.Text)
        If Len(textValue) > 0 Then
            nonEmptyIndex = nonEmptyIndex + 1
            If nonEmptyIndex = 1 And HasRule(selectedRules, "标题") And Not IsNothing(sampleTitle) Then
                If CopyParagraphFormat(sampleTitle, para) Then changedTitle = changedTitle + 1 Else failureCount = failureCount + 1
            ElseIf IsHeadingParagraph(para) And HasRule(selectedRules, "层级标题") And Not IsNothing(sampleHeading) Then
                If CopyParagraphFormat(sampleHeading, para) Then changedHeading = changedHeading + 1 Else failureCount = failureCount + 1
            ElseIf HasRule(selectedRules, "正文") And Not IsNothing(sampleBody) Then
                If CopyParagraphFormat(sampleBody, para) Then changedBody = changedBody + 1 Else failureCount = failureCount + 1
            End If
        End If
    Next
End Sub

Sub ApplyTableRules(scopeRange, sampleTable, ByRef changedTables, ByRef failureCount)
    Dim tbl
    For Each tbl In scopeRange.Tables
        If CopyTableFormat(sampleTable, tbl) Then changedTables = changedTables + 1 Else failureCount = failureCount + 1
    Next
End Sub

Sub ApplySectionRules(targetDoc, sampleSection, scopeMode, scopeRange, selectedRules, ByRef changedSections, ByRef failureCount)
    If IsNothing(sampleSection) Then Exit Sub
    Dim sectionObj
    For Each sectionObj In targetDoc.Sections
        If scopeMode <> "selection" Or IsRangeInSection(scopeRange, sectionObj.Range) Then
            If CopySectionFormat(sampleSection, sectionObj, selectedRules) Then changedSections = changedSections + 1 Else failureCount = failureCount + 1
            If scopeMode = "selection" Then Exit For
        End If
    Next
End Sub

Function IsRangeInSection(scopeRange, sectionRange)
    On Error Resume Next
    IsRangeInSection = (CLng(scopeRange.Start) >= CLng(sectionRange.Start) And CLng(scopeRange.Start) <= CLng(sectionRange.End))
    If Err.Number <> 0 Then IsRangeInSection = False
    Err.Clear
End Function

Function CopyParagraphFormat(sourcePara, targetPara)
    On Error Resume Next
    Err.Clear
    targetPara.Range.Font.NameFarEast = sourcePara.Range.Font.NameFarEast
    targetPara.Range.Font.Name = sourcePara.Range.Font.Name
    targetPara.Range.Font.Size = sourcePara.Range.Font.Size
    targetPara.Range.Font.Bold = sourcePara.Range.Font.Bold
    targetPara.Range.Font.Italic = sourcePara.Range.Font.Italic
    targetPara.Range.Font.Underline = sourcePara.Range.Font.Underline
    targetPara.Range.Font.Color = sourcePara.Range.Font.Color
    targetPara.Format.Alignment = sourcePara.Format.Alignment
    targetPara.Format.LeftIndent = sourcePara.Format.LeftIndent
    targetPara.Format.RightIndent = sourcePara.Format.RightIndent
    targetPara.Format.FirstLineIndent = sourcePara.Format.FirstLineIndent
    targetPara.Format.SpaceBefore = sourcePara.Format.SpaceBefore
    targetPara.Format.SpaceAfter = sourcePara.Format.SpaceAfter
    targetPara.Format.LineSpacingRule = sourcePara.Format.LineSpacingRule
    targetPara.Format.LineSpacing = sourcePara.Format.LineSpacing
    targetPara.Format.KeepWithNext = sourcePara.Format.KeepWithNext
    targetPara.Format.KeepTogether = sourcePara.Format.KeepTogether
    targetPara.Format.PageBreakBefore = sourcePara.Format.PageBreakBefore
    CopyParagraphFormat = (Err.Number = 0)
    Err.Clear
End Function

Function CopyTableFormat(sourceTable, targetTable)
    On Error Resume Next
    Dim copied
    copied = 0
    Err.Clear
    targetTable.AllowAutoFit = sourceTable.AllowAutoFit
    If Err.Number = 0 Then copied = copied + 1
    Err.Clear
    targetTable.Rows.Alignment = sourceTable.Rows.Alignment
    If Err.Number = 0 Then copied = copied + 1
    Err.Clear
    targetTable.Range.Font.NameFarEast = sourceTable.Range.Font.NameFarEast
    If Err.Number = 0 Then copied = copied + 1
    Err.Clear
    targetTable.Range.Font.Name = sourceTable.Range.Font.Name
    If Err.Number = 0 Then copied = copied + 1
    Err.Clear
    targetTable.Range.Font.Size = sourceTable.Range.Font.Size
    If Err.Number = 0 Then copied = copied + 1
    Err.Clear
    targetTable.Range.Font.Bold = sourceTable.Range.Font.Bold
    If Err.Number = 0 Then copied = copied + 1
    Err.Clear
    CopyOneBorder sourceTable, targetTable, -1
    CopyOneBorder sourceTable, targetTable, -2
    CopyOneBorder sourceTable, targetTable, -3
    CopyOneBorder sourceTable, targetTable, -4
    CopyOneBorder sourceTable, targetTable, -5
    CopyOneBorder sourceTable, targetTable, -6
    CopyTableFormat = (copied > 0)
    Err.Clear
End Function

Sub CopyOneBorder(sourceTable, targetTable, borderIndex)
    On Error Resume Next
    Err.Clear
    targetTable.Borders(borderIndex).LineStyle = sourceTable.Borders(borderIndex).LineStyle
    Err.Clear
    targetTable.Borders(borderIndex).LineWidth = sourceTable.Borders(borderIndex).LineWidth
    Err.Clear
    targetTable.Borders(borderIndex).Color = sourceTable.Borders(borderIndex).Color
    Err.Clear
End Sub

Function CopySectionFormat(sourceSection, targetSection, selectedRules)
    On Error Resume Next
    Dim copied
    copied = 0
    Err.Clear
    If HasRule(selectedRules, "页面") Then
        targetSection.PageSetup.TopMargin = sourceSection.PageSetup.TopMargin
        If Err.Number = 0 Then copied = copied + 1
        Err.Clear
        targetSection.PageSetup.BottomMargin = sourceSection.PageSetup.BottomMargin
        If Err.Number = 0 Then copied = copied + 1
        Err.Clear
        targetSection.PageSetup.LeftMargin = sourceSection.PageSetup.LeftMargin
        If Err.Number = 0 Then copied = copied + 1
        Err.Clear
        targetSection.PageSetup.RightMargin = sourceSection.PageSetup.RightMargin
        If Err.Number = 0 Then copied = copied + 1
        Err.Clear
        targetSection.PageSetup.Orientation = sourceSection.PageSetup.Orientation
        If Err.Number = 0 Then copied = copied + 1
        Err.Clear
    End If
    If HasRule(selectedRules, "页眉页脚") Then
        If CopyRangeFormat(sourceSection.Headers(1).Range, targetSection.Headers(1).Range) Then copied = copied + 1
        If CopyRangeFormat(sourceSection.Footers(1).Range, targetSection.Footers(1).Range) Then copied = copied + 1
    End If
    CopySectionFormat = (copied > 0)
    Err.Clear
End Function

Function CopyRangeFormat(sourceRange, targetRange)
    On Error Resume Next
    Dim copied
    copied = 0
    Err.Clear
    targetRange.Font.NameFarEast = sourceRange.Font.NameFarEast
    If Err.Number = 0 Then copied = copied + 1
    Err.Clear
    targetRange.Font.Name = sourceRange.Font.Name
    If Err.Number = 0 Then copied = copied + 1
    Err.Clear
    targetRange.Font.Size = sourceRange.Font.Size
    If Err.Number = 0 Then copied = copied + 1
    Err.Clear
    targetRange.Font.Bold = sourceRange.Font.Bold
    If Err.Number = 0 Then copied = copied + 1
    Err.Clear
    targetRange.ParagraphFormat.Alignment = sourceRange.ParagraphFormat.Alignment
    If Err.Number = 0 Then copied = copied + 1
    Err.Clear
    targetRange.ParagraphFormat.SpaceBefore = sourceRange.ParagraphFormat.SpaceBefore
    If Err.Number = 0 Then copied = copied + 1
    Err.Clear
    targetRange.ParagraphFormat.SpaceAfter = sourceRange.ParagraphFormat.SpaceAfter
    If Err.Number = 0 Then copied = copied + 1
    CopyRangeFormat = (copied > 0)
    Err.Clear
End Function

Sub CountImpact(scopeRange, ByRef titleCount, ByRef headingCount, ByRef bodyCount, ByRef tableCount)
    Dim para, textValue, nonEmptyIndex
    titleCount = 0
    headingCount = 0
    bodyCount = 0
    tableCount = SafeCount(scopeRange.Tables.Count)
    nonEmptyIndex = 0
    For Each para In scopeRange.Paragraphs
        textValue = CleanParagraphText(para.Range.Text)
        If Len(textValue) > 0 Then
            nonEmptyIndex = nonEmptyIndex + 1
            If nonEmptyIndex = 1 Then
                titleCount = titleCount + 1
            ElseIf IsHeadingParagraph(para) Then
                headingCount = headingCount + 1
            Else
                bodyCount = bodyCount + 1
            End If
        End If
    Next
End Sub

Function FindFirstNonEmptyParagraph(sourceRange)
    Dim para
    Set FindFirstNonEmptyParagraph = Nothing
    For Each para In sourceRange.Paragraphs
        If Len(CleanParagraphText(para.Range.Text)) > 0 Then
            Set FindFirstNonEmptyParagraph = para
            Exit Function
        End If
    Next
End Function

Function FindFirstHeadingParagraph(sourceRange)
    Dim para
    Set FindFirstHeadingParagraph = Nothing
    For Each para In sourceRange.Paragraphs
        If Len(CleanParagraphText(para.Range.Text)) > 0 And IsHeadingParagraph(para) Then
            Set FindFirstHeadingParagraph = para
            Exit Function
        End If
    Next
End Function

Function FindFirstBodyParagraph(sourceRange, titlePara)
    Dim para
    Set FindFirstBodyParagraph = Nothing
    For Each para In sourceRange.Paragraphs
        If Len(CleanParagraphText(para.Range.Text)) > 0 Then
            If IsNothing(titlePara) Or SafeLong(para.Range.Start) <> SafeLong(titlePara.Range.Start) Then
                If Not IsHeadingParagraph(para) Then
                    Set FindFirstBodyParagraph = para
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Function FindFirstTable(doc)
    Set FindFirstTable = Nothing
    If SafeCount(doc.Tables.Count) > 0 Then Set FindFirstTable = doc.Tables(1)
End Function

Function FindFirstSection(doc)
    Set FindFirstSection = Nothing
    If SafeCount(doc.Sections.Count) > 0 Then Set FindFirstSection = doc.Sections(1)
End Function

Sub CaptureTargetSelection(appObj, targetDoc, ByRef selectionStart, ByRef selectionEnd)
    On Error Resume Next
    selectionStart = CLng(appObj.Selection.Range.Start)
    selectionEnd = CLng(appObj.Selection.Range.End)
    If Err.Number <> 0 Then
        selectionStart = -1
        selectionEnd = -1
        Err.Clear
    End If
End Sub

Function GetTargetRange(targetDoc, scopeMode, selectionStart, selectionEnd)
    On Error Resume Next
    Set GetTargetRange = Nothing
    If scopeMode = "selection" Then
        If CLng(selectionStart) >= 0 And CLng(selectionEnd) >= CLng(selectionStart) Then
            Set GetTargetRange = targetDoc.Range(CLng(selectionStart), CLng(selectionEnd)).Duplicate
        End If
    Else
        Set GetTargetRange = targetDoc.Content.Duplicate
    End If
    If Err.Number <> 0 Then
        Set GetTargetRange = Nothing
        Err.Clear
    End If
End Function

Function IsHeadingParagraph(para)
    On Error Resume Next
    IsHeadingParagraph = False
    If CLng(para.OutlineLevel) <> wdOutlineLevelBodyText Then
        IsHeadingParagraph = True
    ElseIf InStr(1, CStr(para.Style), "标题", vbTextCompare) > 0 Or InStr(1, CStr(para.Style), "Heading", vbTextCompare) > 0 Then
        IsHeadingParagraph = True
    End If
    If Err.Number <> 0 Then
        IsHeadingParagraph = False
        Err.Clear
    End If
End Function

Function BuildPreviewText(sampleName, scopeMode, selectedRules, titleCount, headingCount, bodyCount, tableCount, sectionCount, unavailableRules)
    Dim text
    text = "Word 按样本套版差异预览（尚未写入）" & vbCrLf & _
        "样本：" & sampleName & vbCrLf & _
        "范围：" & ScopeLabel(scopeMode) & vbCrLf & _
        "规则：" & selectedRules & vbCrLf & _
        "预计影响：标题=" & CStr(titleCount) & "，层级标题=" & CStr(headingCount) & _
        "，正文=" & CStr(bodyCount) & "，表格=" & CStr(tableCount) & "，节=" & CStr(sectionCount)
    If Len(unavailableRules) > 0 Then text = text & vbCrLf & "不可映射：样本中未找到 " & unavailableRules
    text = text & vbCrLf & "内容保护：仅复制格式属性，不复制样本文字、页眉页脚文字或页码字段。"
    BuildPreviewText = text
End Function

Function BuildPlanParams(samplePath, scopeMode, selectedRules, titleCount, headingCount, bodyCount, tableCount, sectionCount)
    BuildPlanParams = "{""samplePath"":""" & EscapeJson(samplePath) & """,""scope"":""" & scopeMode & _
        """,""rules"":""" & EscapeJson(selectedRules) & """,""impact"":{""title"":" & CStr(titleCount) & _
        ",""heading"":" & CStr(headingCount) & ",""body"":" & CStr(bodyCount) & ",""tables"":" & CStr(tableCount) & _
        ",""sections"":" & CStr(sectionCount) & "}}"
End Function

Function SuccessPreviewJson(samplePath, profileId, profileVersion, scopeMode, selectedRules, titleCount, headingCount, bodyCount, tableCount, sectionCount, unavailableRules, planPreview, planValidation, previewText)
    SuccessPreviewJson = "{""ok"":true,""applied"":false,""previewOnly"":true,""requiresConfirmation"":true,""samplePath"":""" & EscapeJson(samplePath) & _
        """,""profileId"":""" & EscapeJson(profileId) & """,""profileVersion"":""" & EscapeJson(profileVersion) & _
        """,""scope"":""" & scopeMode & """,""rules"":""" & EscapeJson(selectedRules) & _
        """,""impact"":{""title"":" & CStr(titleCount) & ",""heading"":" & CStr(headingCount) & _
        ",""body"":" & CStr(bodyCount) & ",""tables"":" & CStr(tableCount) & ",""sections"":" & CStr(sectionCount) & _
        "},""unavailableRules"":""" & EscapeJson(unavailableRules) & """,""writePlanPreview"":""" & EscapeJson(planPreview) & _
        """,""writePlanValidation"":""" & EscapeJson(planValidation) & """,""message"":""" & EscapeJson(previewText) & """}"
End Function

Function FailureJson(code, message)
    FailureJson = "{""ok"":false,""code"":""" & EscapeJson(code) & """,""message"":""" & EscapeJson(message) & """}"
End Function

Function FailureWithRollbackJson(code, message, rollbackOk, failureCount, contentPreserved, backupPath)
    FailureWithRollbackJson = "{""ok"":false,""code"":""" & EscapeJson(code) & """,""rollbackAttempted"":true,""rollbackOk"":" & JsonBool(rollbackOk) & _
        ",""failureCount"":" & CStr(failureCount) & ",""contentPreservedBeforeRollback"":" & JsonBool(contentPreserved) & _
        ",""backupPath"":""" & EscapeJson(backupPath) & """,""message"":""" & EscapeJson(message) & """}"
End Function

Function NormalizeRunMode(value)
    Dim text
    text = LCase(Trim(CStr(value)))
    If text = "应用" Or text = "apply" Then NormalizeRunMode = "apply" Else NormalizeRunMode = "preview"
End Function

Function NormalizeScopeMode(value)
    Dim text
    text = LCase(Trim(CStr(value)))
    If text = "选区" Or text = "selection" Then NormalizeScopeMode = "selection" Else NormalizeScopeMode = "document"
End Function

Function NormalizeRules(value)
    Dim text
    text = Replace(Trim(CStr(value)), "，", ",")
    text = Replace(text, "、", ",")
    If Not ContainsKnownRule(text) Then text = "标题,层级标题,正文,表格,页面,页眉页脚"
    NormalizeRules = text
End Function

Function ContainsKnownRule(value)
    ContainsKnownRule = (InStr(value, "标题") > 0 Or InStr(value, "正文") > 0 Or InStr(value, "表格") > 0 Or InStr(value, "页面") > 0 Or InStr(value, "页眉页脚") > 0)
End Function

Function HasRule(selectedRules, ruleName)
    HasRule = (InStr(1, "," & selectedRules & ",", "," & ruleName & ",", vbTextCompare) > 0)
End Function

Function ScopeLabel(scopeMode)
    If scopeMode = "selection" Then ScopeLabel = "当前选区" Else ScopeLabel = "全文"
End Function

Function CountTargetSections(targetDoc, scopeMode)
    If scopeMode = "selection" Then CountTargetSections = 1 Else CountTargetSections = SafeCount(targetDoc.Sections.Count)
End Function

Function IsNothing(value)
    IsNothing = (TypeName(value) = "Nothing" Or TypeName(value) = "Empty")
End Function

Function IsDocumentProtected(doc)
    On Error Resume Next
    IsDocumentProtected = (CLng(doc.ProtectionType) <> wdNoProtection)
    If Err.Number <> 0 Then
        IsDocumentProtected = False
        Err.Clear
    End If
End Function

Function CleanParagraphText(value)
    Dim text
    text = CStr(value)
    text = Replace(text, vbCr, "")
    text = Replace(text, vbLf, "")
    text = Replace(text, ChrW(7), "")
    text = Replace(text, ChrW(160), " ")
    text = Replace(text, ChrW(12288), " ")
    CleanParagraphText = Trim(text)
End Function

Function AppendItem(currentValue, newValue)
    If Len(currentValue) = 0 Then AppendItem = newValue Else AppendItem = currentValue & "," & newValue
End Function

Function SafeDocumentPath(doc)
    On Error Resume Next
    SafeDocumentPath = CStr(doc.FullName)
    If Err.Number <> 0 Then SafeDocumentPath = ""
    Err.Clear
End Function

Function IsDocumentClean(doc)
    On Error Resume Next
    IsDocumentClean = CBool(doc.Saved)
    If Err.Number <> 0 Then IsDocumentClean = False
    Err.Clear
End Function

Function IsTrackingRevisions(doc)
    On Error Resume Next
    IsTrackingRevisions = CBool(doc.TrackRevisions)
    If Err.Number <> 0 Then IsTrackingRevisions = False
    Err.Clear
End Function

Function ReloadOriginalDocument(appObj, targetDoc, targetPath)
    On Error Resume Next
    Err.Clear
    targetDoc.Close False
    Dim reopenedDoc
    Set reopenedDoc = appObj.Documents.Open(targetPath)
    ReloadOriginalDocument = (Err.Number = 0 And Not IsNothing(reopenedDoc))
    Err.Clear
End Function

Function SafeSelectFile(title, filter)
    On Error Resume Next
    SafeSelectFile = Host.SelectFile(title, filter)
    If Err.Number <> 0 Then SafeSelectFile = ""
    Err.Clear
End Function

Function SafeReadTextFile(path)
    On Error Resume Next
    SafeReadTextFile = Host.ReadTextFile(path, 1048576)
    If Err.Number <> 0 Then SafeReadTextFile = ""
    Err.Clear
End Function

Function SafePathExists(path)
    On Error Resume Next
    SafePathExists = CBool(Host.PathExists(path))
    If Err.Number <> 0 Then SafePathExists = False
    Err.Clear
End Function

Function IsJsonProfilePath(path)
    IsJsonProfilePath = (LCase(Right(Trim(CStr(path)), 5)) = ".json")
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

Function ExtractJsonObject(jsonText, keyName)
    Dim marker, keyPos, colonPos, objectPos, i, ch, nextCh, depth, inString
    ExtractJsonObject = ""
    marker = """" & CStr(keyName) & """"
    keyPos = InStr(1, CStr(jsonText), marker, vbTextCompare)
    If keyPos <= 0 Then Exit Function
    colonPos = InStr(keyPos + Len(marker), CStr(jsonText), ":")
    If colonPos <= 0 Then Exit Function
    objectPos = InStr(colonPos + 1, CStr(jsonText), "{")
    If objectPos <= 0 Then Exit Function
    depth = 0
    inString = False
    i = objectPos
    Do While i <= Len(jsonText)
        ch = Mid(jsonText, i, 1)
        If inString Then
            If ch = "\" And i < Len(jsonText) Then
                i = i + 2
            ElseIf ch = """" Then
                inString = False
                i = i + 1
            Else
                i = i + 1
            End If
        Else
            If ch = """" Then
                inString = True
            ElseIf ch = "{" Then
                depth = depth + 1
            ElseIf ch = "}" Then
                depth = depth - 1
                If depth = 0 Then
                    ExtractJsonObject = Mid(jsonText, objectPos, i - objectPos + 1)
                    Exit Function
                End If
            End If
            i = i + 1
        End If
    Loop
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
            If nextCh = "n" Then
                result = result & vbLf
            ElseIf nextCh = "r" Then
                result = result & vbCr
            ElseIf nextCh = "t" Then
                result = result & vbTab
            Else
                result = result & nextCh
            End If
            i = i + 2
        Else
            result = result & ch
            i = i + 1
        End If
    Loop
    ExtractJsonString = result
End Function

Function SafePrompt(message, defaultValue)
    On Error Resume Next
    SafePrompt = Host.Prompt(message, defaultValue)
    If Err.Number <> 0 Then SafePrompt = defaultValue
    Err.Clear
End Function

Function SafeCreateBackup(label)
    On Error Resume Next
    SafeCreateBackup = Host.CreateBackup(label)
    If Err.Number <> 0 Then SafeCreateBackup = ""
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
    Host.RecordWrite actionId, paramsJson, "office", "word.reload_from_backup", "{""strategy"":""backup_restore""}", capabilityId
    Err.Clear
End Sub

Function SafePreviewWritePlan()
    On Error Resume Next
    SafePreviewWritePlan = Host.PreviewWritePlan()
    If Err.Number <> 0 Then SafePreviewWritePlan = ""
    Err.Clear
End Function

Function SafeValidateWritePlan(planId)
    On Error Resume Next
    SafeValidateWritePlan = Host.ValidateWritePlan(planId, True)
    If Err.Number <> 0 Then SafeValidateWritePlan = ""
    Err.Clear
End Function

Sub SafeRollbackWritePlan(planId)
    On Error Resume Next
    Host.RollbackWritePlan planId
    Err.Clear
End Sub

Sub SafeCloseDocument(doc)
    On Error Resume Next
    doc.Close False
    Err.Clear
End Sub

Sub SafeWriteLog(message)
    On Error Resume Next
    Host.WriteLog message
    Err.Clear
End Sub

Function SafeCount(value)
    On Error Resume Next
    SafeCount = CLng(value)
    If Err.Number <> 0 Then SafeCount = 0
    Err.Clear
End Function

Function SafeLong(value)
    On Error Resume Next
    SafeLong = CLng(value)
    If Err.Number <> 0 Then SafeLong = 0
    Err.Clear
End Function

Function JsonBool(value)
    If value Then JsonBool = "true" Else JsonBool = "false"
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
