' 函数名: HostPptApplySampleFooterPageNumber
' 描述: 默认预览；确认后按样本逐页应用页脚和页码可见性，不复制页脚文字、日期、母版、主题或业务内容
' 适用应用: PowerPoint
' 搜索范围: 全文
' 搜索对象: 无

Option Explicit

Const msoTrue = -1
Const msoFalse = 0
Const msoPlaceholder = 14
Const ppPlaceholderSlideNumber = 13
Const ppPlaceholderHeader = 14
Const ppPlaceholderFooter = 15
Const ppPlaceholderDate = 16

Function Main()
    On Error Resume Next
    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = FailureJson("E_NO_PPT_APP", "未取得 PowerPoint 应用，请在 PowerPoint 中运行该预设")
        Exit Function
    End If
    Main = HostPptApplySampleFooterPageNumber(appObj)
End Function

Function HostPptApplySampleFooterPageNumber(appObj)
    On Error Resume Next
    Dim targetPres
    Set targetPres = appObj.ActivePresentation
    If Err.Number <> 0 Or TypeName(targetPres) = "Empty" Or TypeName(targetPres) = "Nothing" Then
        Err.Clear
        HostPptApplySampleFooterPageNumber = FailureJson("E_NO_PRESENTATION", "当前没有活动演示文稿")
        Exit Function
    End If

    Dim samplePath, runMode, selectedRules
    samplePath = Trim(SafeSelectFile("选择规范 PowerPoint 样本（只迁移页脚/页码可见性，不复制页脚文字）", "PowerPoint 演示文稿|*.ppt;*.pptx;*.pptm;*.ppsx"))
    If Len(samplePath) = 0 Then
        HostPptApplySampleFooterPageNumber = FailureJson("E_SAMPLE_REQUIRED", "未选择样本演示文稿；未修改当前演示文稿")
        Exit Function
    End If
    If IsSamePresentation(targetPres, samplePath) Then
        HostPptApplySampleFooterPageNumber = FailureJson("E_SAMPLE_IS_TARGET", "样本不能与当前目标演示文稿相同")
        Exit Function
    End If
    runMode = LCase(Trim(SafePrompt("运行方式：输入“预览”仅查看影响，输入“应用”进入二次确认", "预览")))
    selectedRules = NormalizeRules(SafePrompt("选择要迁移的页面规则（逗号分隔）：页脚可见性、页码可见性", "页脚可见性,页码可见性"))
    If Len(selectedRules) = 0 Then selectedRules = "页脚可见性,页码可见性"
    If Not HasAnySupportedRule(selectedRules) Then
        HostPptApplySampleFooterPageNumber = FailureJson("E_RULES_REQUIRED", "未选择有效页面规则；可选：页脚可见性、页码可见性")
        Exit Function
    End If

    Dim samplePres
    Set samplePres = Nothing
    Err.Clear
    Set samplePres = appObj.Presentations.Open(samplePath, True, False, False)
    If Err.Number <> 0 Or TypeName(samplePres) = "Empty" Or TypeName(samplePres) = "Nothing" Then
        Dim openError
        openError = Err.Description
        Err.Clear
        HostPptApplySampleFooterPageNumber = FailureJson("E_SAMPLE_OPEN", "无法以只读方式打开样本演示文稿：" & openError)
        Exit Function
    End If

    Dim sameSize, pairedSlides, footerChanges, pageNumberChanges, unmatchedSlides, unavailableItems
    sameSize = SameSlideSize(samplePres, targetPres)
    CountVisibilityImpact samplePres, targetPres, selectedRules, pairedSlides, footerChanges, pageNumberChanges, unmatchedSlides, unavailableItems

    Dim previewText, planId, planPreview, planValidation
    previewText = "PPT 页脚/页码按样本应用预览；样本=" & samplePres.Name & "；目标=" & targetPres.Name & _
        "；规则=" & selectedRules & "；页面尺寸一致=" & BoolLabel(sameSize) & "；配对页=" & CStr(pairedSlides) & _
        "；页脚可见性变化=" & CStr(footerChanges) & "；页码可见性变化=" & CStr(pageNumberChanges) & _
        "；未配对页=" & CStr(unmatchedSlides) & "；不可读项=" & CStr(unavailableItems) & vbCrLf & _
        "不会复制页脚文字、日期、母版、主题、背景、页面或业务内容；成功结果不会自动保存。"
    planId = SafeBeginWritePlan()
    SafeRecordWrite "ppt_sample_footer_page_number_transfer", BuildPlanParams(samplePath, selectedRules, pairedSlides, footerChanges, pageNumberChanges, unmatchedSlides, unavailableItems), "office.ppt.layout.footerPageNumberTransfer"
    planPreview = SafePreviewWritePlan()
    planValidation = SafeValidateWritePlan(planId)
    SafeRollbackWritePlan planId

    If runMode <> "apply" And runMode <> "应用" Then
        Host.WriteClipboard previewText
        SafeWriteLog previewText
        SafeClosePresentation samplePres
        HostPptApplySampleFooterPageNumber = PreviewJson(samplePath, selectedRules, sameSize, pairedSlides, footerChanges, pageNumberChanges, unmatchedSlides, unavailableItems, planPreview, planValidation, previewText)
        Exit Function
    End If

    If Not sameSize Then
        SafeClosePresentation samplePres
        HostPptApplySampleFooterPageNumber = FailureJson("E_SLIDE_SIZE_MISMATCH", "样本与目标页面尺寸不同，已停止且未修改演示文稿")
        Exit Function
    End If

    Dim confirmation
    confirmation = Trim(SafePrompt(previewText & vbCrLf & vbCrLf & "确认仅修改页脚/页码可见性且不复制页脚文字，请输入：应用", ""))
    If StrComp(confirmation, "应用", vbTextCompare) <> 0 Then
        SafeClosePresentation samplePres
        HostPptApplySampleFooterPageNumber = FailureJson("E_CONFIRM_REQUIRED", "未输入“应用”，已取消且未修改演示文稿")
        Exit Function
    End If

    Dim targetPath
    targetPath = SafePresentationPath(targetPres)
    If Len(targetPath) = 0 Or Not IsPresentationClean(targetPres) Then
        SafeClosePresentation samplePres
        HostPptApplySampleFooterPageNumber = FailureJson("E_SAVE_REQUIRED", "为保证可恢复，请先保存当前演示文稿及已有修改，再重新运行")
        Exit Function
    End If

    Dim backupPath
    backupPath = SafeCreateBackup("ppt_sample_footer_page_number_transfer")
    If Len(backupPath) = 0 Then
        SafeClosePresentation samplePres
        HostPptApplySampleFooterPageNumber = FailureJson("E_BACKUP_REQUIRED", "未能创建执行前备份，已停止且未修改演示文稿")
        Exit Function
    End If

    Dim businessBefore, footerTextBefore
    businessBefore = BusinessTextFingerprint(targetPres)
    footerTextBefore = FooterTextFingerprint(targetPres)

    Dim appliedFooter, appliedPageNumber, failureCount
    appliedFooter = 0
    appliedPageNumber = 0
    failureCount = 0
    ApplyVisibilityRules samplePres, targetPres, selectedRules, appliedFooter, appliedPageNumber, failureCount

    Dim contentPreserved, footerTextPreserved
    contentPreserved = (BusinessTextFingerprint(targetPres) = businessBefore)
    footerTextPreserved = (FooterTextFingerprint(targetPres) = footerTextBefore)
    If failureCount > 0 Or Not contentPreserved Or Not footerTextPreserved Then
        SafeClosePresentation samplePres
        Dim reloadOk
        reloadOk = DiscardAndReloadPresentation(appObj, targetPres, targetPath)
        HostPptApplySampleFooterPageNumber = RollbackJson("E_APPLY_ROLLED_BACK", "写入失败或检测到文字变化，已放弃未保存结果并重新加载原文件", reloadOk, contentPreserved, footerTextPreserved, failureCount, backupPath)
        Exit Function
    End If

    SafeClosePresentation samplePres
    Dim summary
    summary = "PPT 页脚/页码按样本应用完成；页脚可见性=" & CStr(appliedFooter) & _
        "；页码可见性=" & CStr(appliedPageNumber) & _
        "；业务文字保持=true；页脚文字保持=true；结果尚未保存"
    If Len(backupPath) > 0 Then summary = summary & "；备份=" & backupPath
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostPptApplySampleFooterPageNumber = "{""ok"":true,""applied"":true,""contentPreserved"":true,""footerTextPreserved"":true,""changed"":{""footerVisibility"":" & _
        CStr(appliedFooter) & ",""slideNumberVisibility"":" & CStr(appliedPageNumber) & "},""backupPath"":""" & EscapeJson(backupPath) & _
        """,""message"":""" & EscapeJson(summary) & """}"
End Function

Sub CountVisibilityImpact(samplePres, targetPres, selectedRules, ByRef pairedSlides, ByRef footerChanges, ByRef pageNumberChanges, ByRef unmatchedSlides, ByRef unavailableItems)
    Dim i, sampleValue, targetValue, sampleOk, targetOk
    pairedSlides = Minimum(SafeCount(samplePres.Slides.Count), SafeCount(targetPres.Slides.Count))
    footerChanges = 0
    pageNumberChanges = 0
    unmatchedSlides = Abs(SafeCount(samplePres.Slides.Count) - SafeCount(targetPres.Slides.Count))
    unavailableItems = 0
    For i = 1 To pairedSlides
        If HasRule(selectedRules, "页脚可见性") Then
            sampleOk = TryGetFooterVisible(samplePres.Slides(i), sampleValue)
            targetOk = TryGetFooterVisible(targetPres.Slides(i), targetValue)
            If sampleOk And targetOk Then
                If sampleValue <> targetValue Then footerChanges = footerChanges + 1
            Else
                unavailableItems = unavailableItems + 1
            End If
        End If
        If HasRule(selectedRules, "页码可见性") Then
            sampleOk = TryGetSlideNumberVisible(samplePres.Slides(i), sampleValue)
            targetOk = TryGetSlideNumberVisible(targetPres.Slides(i), targetValue)
            If sampleOk And targetOk Then
                If sampleValue <> targetValue Then pageNumberChanges = pageNumberChanges + 1
            Else
                unavailableItems = unavailableItems + 1
            End If
        End If
    Next
End Sub

Sub ApplyVisibilityRules(samplePres, targetPres, selectedRules, ByRef appliedFooter, ByRef appliedPageNumber, ByRef failureCount)
    Dim pairedSlides, i, sampleValue, targetValue, sampleOk, targetOk
    pairedSlides = Minimum(SafeCount(samplePres.Slides.Count), SafeCount(targetPres.Slides.Count))
    For i = 1 To pairedSlides
        If HasRule(selectedRules, "页脚可见性") Then
            sampleOk = TryGetFooterVisible(samplePres.Slides(i), sampleValue)
            targetOk = TryGetFooterVisible(targetPres.Slides(i), targetValue)
            If sampleOk And targetOk Then
                If sampleValue <> targetValue Then
                    If SetFooterVisible(targetPres.Slides(i), sampleValue) Then appliedFooter = appliedFooter + 1 Else failureCount = failureCount + 1
                End If
            Else
                failureCount = failureCount + 1
            End If
        End If
        If HasRule(selectedRules, "页码可见性") Then
            sampleOk = TryGetSlideNumberVisible(samplePres.Slides(i), sampleValue)
            targetOk = TryGetSlideNumberVisible(targetPres.Slides(i), targetValue)
            If sampleOk And targetOk Then
                If sampleValue <> targetValue Then
                    If SetSlideNumberVisible(targetPres.Slides(i), sampleValue) Then appliedPageNumber = appliedPageNumber + 1 Else failureCount = failureCount + 1
                End If
            Else
                failureCount = failureCount + 1
            End If
        End If
    Next
End Sub

Function TryGetFooterVisible(slide, ByRef visible)
    On Error Resume Next
    Err.Clear
    visible = (CLng(slide.HeadersFooters.Footer.Visible) = msoTrue)
    TryGetFooterVisible = (Err.Number = 0)
    If Err.Number <> 0 Then visible = False
    Err.Clear
End Function

Function TryGetSlideNumberVisible(slide, ByRef visible)
    On Error Resume Next
    Err.Clear
    visible = (CLng(slide.HeadersFooters.SlideNumber.Visible) = msoTrue)
    TryGetSlideNumberVisible = (Err.Number = 0)
    If Err.Number <> 0 Then visible = False
    Err.Clear
End Function

Function SetFooterVisible(slide, visible)
    On Error Resume Next
    SetFooterVisible = False
    If visible Then
        slide.HeadersFooters.Footer.Visible = msoTrue
    Else
        slide.HeadersFooters.Footer.Visible = msoFalse
    End If
    SetFooterVisible = (Err.Number = 0)
    Err.Clear
End Function

Function SetSlideNumberVisible(slide, visible)
    On Error Resume Next
    SetSlideNumberVisible = False
    If visible Then
        slide.HeadersFooters.SlideNumber.Visible = msoTrue
    Else
        slide.HeadersFooters.SlideNumber.Visible = msoFalse
    End If
    SetSlideNumberVisible = (Err.Number = 0)
    Err.Clear
End Function

Function BusinessTextFingerprint(pres)
    On Error Resume Next
    Dim slide, shape, checksum, textLength
    checksum = 17
    textLength = 0
    For Each slide In pres.Slides
        For Each shape In slide.Shapes
            If Not IsPageDecorationPlaceholder(shape) Then
                If shape.HasTextFrame = msoTrue And shape.TextFrame.HasText = msoTrue Then
                    AppendTextFingerprint checksum, textLength, CStr(shape.TextFrame.TextRange.Text), slide.SlideIndex
                End If
            End If
        Next
    Next
    BusinessTextFingerprint = CStr(textLength) & ":" & CStr(checksum)
    Err.Clear
End Function

Function FooterTextFingerprint(pres)
    On Error Resume Next
    Dim slide, checksum, textLength, textValue
    checksum = 23
    textLength = 0
    For Each slide In pres.Slides
        textValue = ""
        textValue = CStr(slide.HeadersFooters.Footer.Text)
        AppendTextFingerprint checksum, textLength, textValue, slide.SlideIndex
        If Err.Number <> 0 Then Err.Clear
    Next
    FooterTextFingerprint = CStr(textLength) & ":" & CStr(checksum)
End Function

Sub AppendTextFingerprint(ByRef checksum, ByRef textLength, textValue, seed)
    Dim i, codeValue
    textLength = textLength + Len(textValue)
    For i = 1 To Len(textValue)
        codeValue = AscW(Mid(textValue, i, 1))
        checksum = (checksum * 31 + codeValue + seed + i) Mod 1000003
    Next
End Sub

Function IsPageDecorationPlaceholder(shape)
    On Error Resume Next
    Dim placeholderType
    IsPageDecorationPlaceholder = False
    If CLng(shape.Type) = msoPlaceholder Then
        placeholderType = CLng(shape.PlaceholderFormat.Type)
        IsPageDecorationPlaceholder = (placeholderType = ppPlaceholderSlideNumber Or placeholderType = ppPlaceholderHeader Or _
            placeholderType = ppPlaceholderFooter Or placeholderType = ppPlaceholderDate)
    End If
    If Err.Number <> 0 Then IsPageDecorationPlaceholder = False
    Err.Clear
End Function

Function SameSlideSize(samplePres, targetPres)
    On Error Resume Next
    SameSlideSize = (Abs(CDbl(samplePres.PageSetup.SlideWidth) - CDbl(targetPres.PageSetup.SlideWidth)) < 0.1 And _
        Abs(CDbl(samplePres.PageSetup.SlideHeight) - CDbl(targetPres.PageSetup.SlideHeight)) < 0.1)
    If Err.Number <> 0 Then SameSlideSize = False
    Err.Clear
End Function

Function IsPresentationClean(pres)
    On Error Resume Next
    IsPresentationClean = CBool(pres.Saved)
    If Err.Number <> 0 Then IsPresentationClean = False
    Err.Clear
End Function

Function SafePresentationPath(pres)
    On Error Resume Next
    SafePresentationPath = CStr(pres.FullName)
    If Err.Number <> 0 Then SafePresentationPath = ""
    Err.Clear
End Function

Function DiscardAndReloadPresentation(appObj, targetPres, targetPath)
    On Error Resume Next
    targetPres.Saved = msoTrue
    targetPres.Close
    Dim reopenedPres
    Set reopenedPres = appObj.Presentations.Open(targetPath, False, False, True)
    DiscardAndReloadPresentation = (Err.Number = 0 And Not reopenedPres Is Nothing)
    Err.Clear
End Function

Function IsSamePresentation(targetPres, samplePath)
    On Error Resume Next
    IsSamePresentation = (StrComp(CStr(targetPres.FullName), samplePath, vbTextCompare) = 0)
    If Err.Number <> 0 Then IsSamePresentation = False
    Err.Clear
End Function

Function Minimum(leftValue, rightValue)
    If leftValue < rightValue Then Minimum = leftValue Else Minimum = rightValue
End Function

Function SafeCount(value)
    On Error Resume Next
    SafeCount = CLng(value)
    If Err.Number <> 0 Then SafeCount = 0
    Err.Clear
End Function

Function NormalizeRules(value)
    NormalizeRules = Replace(Replace(Trim(CStr(value)), "，", ","), " ", "")
End Function

Function HasRule(rules, ruleName)
    HasRule = (InStr(1, "," & rules & ",", "," & ruleName & ",", vbTextCompare) > 0)
End Function

Function HasAnySupportedRule(rules)
    HasAnySupportedRule = (HasRule(rules, "页脚可见性") Or HasRule(rules, "页码可见性"))
End Function

Function SafeSelectFile(title, filter)
    On Error Resume Next
    SafeSelectFile = Host.SelectFile(title, filter)
    If Err.Number <> 0 Then SafeSelectFile = ""
    Err.Clear
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
    Host.RecordWrite actionId, paramsJson, "office", "ppt.reload_from_backup", "{""strategy"":""backup_restore""}", capabilityId
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

Sub SafeClosePresentation(pres)
    On Error Resume Next
    pres.Close
    Err.Clear
End Sub

Sub SafeWriteLog(message)
    On Error Resume Next
    Host.WriteLog message
    Err.Clear
End Sub

Function BuildPlanParams(samplePath, selectedRules, pairedSlides, footerChanges, pageNumberChanges, unmatchedSlides, unavailableItems)
    BuildPlanParams = "{""samplePath"":""" & EscapeJson(samplePath) & """,""rules"":""" & EscapeJson(selectedRules) & _
        """,""pairedSlides"":" & CStr(pairedSlides) & ",""footerChanges"":" & CStr(footerChanges) & _
        ",""pageNumberChanges"":" & CStr(pageNumberChanges) & ",""unmatchedSlides"":" & CStr(unmatchedSlides) & _
        ",""unavailableItems"":" & CStr(unavailableItems) & "}"
End Function

Function PreviewJson(samplePath, selectedRules, sameSize, pairedSlides, footerChanges, pageNumberChanges, unmatchedSlides, unavailableItems, planPreview, planValidation, message)
    PreviewJson = "{""ok"":true,""previewOnly"":true,""samplePath"":""" & EscapeJson(samplePath) & _
        """,""rules"":""" & EscapeJson(selectedRules) & """,""sameSlideSize"":" & JsonBool(sameSize) & _
        ",""impact"":{""pairedSlides"":" & CStr(pairedSlides) & ",""footerVisibility"":" & CStr(footerChanges) & _
        ",""slideNumberVisibility"":" & CStr(pageNumberChanges) & ",""unmatchedSlides"":" & CStr(unmatchedSlides) & _
        ",""unavailableItems"":" & CStr(unavailableItems) & "},""planPreview"":""" & EscapeJson(planPreview) & _
        """,""planValidation"":""" & EscapeJson(planValidation) & _
        """,""requiresConfirmation"":true,""message"":""" & EscapeJson(message) & """}"
End Function

Function FailureJson(code, message)
    FailureJson = "{""ok"":false,""code"":""" & EscapeJson(code) & """,""message"":""" & EscapeJson(message) & """}"
End Function

Function RollbackJson(code, message, reloadOk, contentPreserved, footerTextPreserved, failureCount, backupPath)
    RollbackJson = "{""ok"":false,""code"":""" & EscapeJson(code) & """,""rollbackReloaded"":" & JsonBool(reloadOk) & _
        ",""contentPreserved"":" & JsonBool(contentPreserved) & ",""footerTextPreserved"":" & JsonBool(footerTextPreserved) & _
        ",""failureCount"":" & CStr(failureCount) & ",""backupPath"":""" & EscapeJson(backupPath) & _
        """,""message"":""" & EscapeJson(message) & """}"
End Function

Function BoolLabel(value)
    If value Then BoolLabel = "是" Else BoolLabel = "否"
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
