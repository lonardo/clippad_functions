' 函数名: HostPptApplySampleLayout
' 描述: 按用户选择的 PPT 样本预览并应用已确认对象的格式与几何属性；不复制文字、图片内容、图表数据或动画
' 适用应用: PowerPoint
' 搜索范围: 全文
' 搜索对象: 无

Option Explicit

Const msoTrue = -1
Const msoPicture = 13
Const msoLinkedPicture = 11
Const msoChart = 3
Const msoGroup = 6
Const msoEmbeddedOLEObject = 7
Const msoLinkedOLEObject = 10
Const msoMedia = 16
Const msoTable = 19
Const msoSmartArt = 24

Function Main()
    On Error Resume Next
    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = FailureJson("E_NO_PPT_APP", "未取得 PowerPoint 应用，请在 PowerPoint 中运行该预设")
        Exit Function
    End If
    Main = HostPptApplySampleLayout(appObj)
End Function

Function HostPptApplySampleLayout(appObj)
    On Error Resume Next
    Dim targetPres
    Set targetPres = appObj.ActivePresentation
    If Err.Number <> 0 Or TypeName(targetPres) = "Empty" Or TypeName(targetPres) = "Nothing" Then
        Err.Clear
        HostPptApplySampleLayout = FailureJson("E_NO_PRESENTATION", "当前没有活动演示文稿")
        Exit Function
    End If

    Dim samplePath, runMode, selectedRules
    samplePath = Trim(SafeSelectFile("选择规范 PowerPoint 样本（默认只预览；不会复制文字、图片内容、图表数据或动画）", "PowerPoint 演示文稿|*.ppt;*.pptx;*.pptm;*.ppsx"))
    If Len(samplePath) = 0 Then
        HostPptApplySampleLayout = FailureJson("E_SAMPLE_REQUIRED", "未选择样本演示文稿；未修改当前演示文稿")
        Exit Function
    End If
    If IsSamePresentation(targetPres, samplePath) Then
        HostPptApplySampleLayout = FailureJson("E_SAMPLE_IS_TARGET", "样本不能与当前目标演示文稿相同；请选择独立的规范样本")
        Exit Function
    End If
    runMode = LCase(Trim(SafePrompt("运行方式：输入“预览”仅查看影响，输入“应用”进入二次确认", "预览")))
    selectedRules = NormalizeRules(SafePrompt("选择要迁移的已确认对象类别（逗号分隔）：标题、文本、图片、形状。默认仅标题", "标题"))
    If Len(selectedRules) = 0 Then selectedRules = "标题"
    If Not HasAnySupportedRule(selectedRules) Then
        HostPptApplySampleLayout = FailureJson("E_RULES_REQUIRED", "未选择有效对象类别；可选：标题、文本、图片、形状")
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
        HostPptApplySampleLayout = FailureJson("E_SAMPLE_OPEN", "无法以只读方式打开样本演示文稿：" & openError)
        Exit Function
    End If

    Dim titleCount, textCount, pictureCount, shapeCount, unavailableCount
    CountApplicableObjects samplePres, targetPres, selectedRules, titleCount, textCount, pictureCount, shapeCount, unavailableCount

    Dim planId, planPreview, planValidation, previewText
    previewText = "PPT 按样本套版预览；样本=" & samplePres.Name & "；目标=" & targetPres.Name & _
        "；规则=" & selectedRules & "；标题=" & CStr(titleCount) & "；文本=" & CStr(textCount) & _
        "；图片=" & CStr(pictureCount) & "；形状=" & CStr(shapeCount) & "；不可用/不匹配=" & CStr(unavailableCount) & vbCrLf & _
        "仅迁移已确认对象的格式和位置/尺寸；不复制文字、图片内容、图表数据、动画、母版或页面。"
    planId = SafeBeginWritePlan()
    SafeRecordWrite "ppt_sample_layout_transfer", BuildPlanParams(samplePath, selectedRules, titleCount, textCount, pictureCount, shapeCount, unavailableCount), "office.ppt.layout.sampleTransfer"
    planPreview = SafePreviewWritePlan()
    planValidation = SafeValidateWritePlan(planId)
    SafeRollbackWritePlan planId

    If runMode <> "apply" And runMode <> "应用" Then
        Host.WriteClipboard previewText
        SafeWriteLog previewText
        SafeClosePresentation samplePres
        HostPptApplySampleLayout = PreviewJson(samplePath, selectedRules, titleCount, textCount, pictureCount, shapeCount, unavailableCount, planPreview, planValidation, previewText)
        Exit Function
    End If

    Dim confirmation
    confirmation = Trim(SafePrompt(previewText & vbCrLf & vbCrLf & "确认只迁移格式和几何属性且不复制业务内容，请输入：应用", ""))
    If StrComp(confirmation, "应用", vbTextCompare) <> 0 Then
        SafeClosePresentation samplePres
        HostPptApplySampleLayout = FailureJson("E_CONFIRM_REQUIRED", "未输入“应用”，已取消且未修改演示文稿")
        Exit Function
    End If

    Dim targetPath
    targetPath = SafePresentationPath(targetPres)
    If Len(targetPath) = 0 Or Not IsPresentationClean(targetPres) Then
        SafeClosePresentation samplePres
        HostPptApplySampleLayout = FailureJson("E_SAVE_REQUIRED", "为保证可恢复，请先保存当前演示文稿及已有修改，再重新运行套版")
        Exit Function
    End If

    Dim backupPath
    backupPath = SafeCreateBackup("ppt_sample_layout_transfer")
    If Len(backupPath) = 0 Then
        SafeClosePresentation samplePres
        HostPptApplySampleLayout = FailureJson("E_BACKUP_REQUIRED", "未能创建执行前备份，已停止且未修改演示文稿")
        Exit Function
    End If

    Dim contentBefore
    contentBefore = PresentationTextFingerprint(targetPres)
    Dim changedTitle, changedText, changedPicture, changedShape, failureCount
    changedTitle = 0
    changedText = 0
    changedPicture = 0
    changedShape = 0
    failureCount = 0
    ApplyConfirmedObjects samplePres, targetPres, selectedRules, changedTitle, changedText, changedPicture, changedShape, failureCount

    Dim contentPreserved
    contentPreserved = (PresentationTextFingerprint(targetPres) = contentBefore)
    If failureCount > 0 Or Not contentPreserved Then
        SafeClosePresentation samplePres
        Dim reloadOk
        reloadOk = DiscardAndReloadPresentation(appObj, targetPres, targetPath)
        HostPptApplySampleLayout = RollbackJson("E_APPLY_ROLLED_BACK", "套版出现失败或检测到文字变化，已放弃未保存结果并重新加载原文件", reloadOk, contentPreserved, failureCount, backupPath)
        Exit Function
    End If

    SafeClosePresentation samplePres
    Dim summary
    summary = "PPT 按样本套版完成；标题=" & CStr(changedTitle) & "；文本=" & CStr(changedText) & _
        "；图片=" & CStr(changedPicture) & "；形状=" & CStr(changedShape) & _
        "；业务文字保持=true；本次结果尚未保存，可关闭不保存或使用备份恢复"
    If Len(backupPath) > 0 Then summary = summary & "；备份=" & backupPath
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostPptApplySampleLayout = "{""ok"":true,""applied"":true,""contentPreserved"":true,""changed"":{""title"":" & CStr(changedTitle) & _
        ",""text"":" & CStr(changedText) & ",""picture"":" & CStr(changedPicture) & ",""shape"":" & CStr(changedShape) & _
        "},""backupPath"":""" & EscapeJson(backupPath) & """,""message"":""" & EscapeJson(summary) & """}"
End Function

Sub CountApplicableObjects(samplePres, targetPres, selectedRules, ByRef titleCount, ByRef textCount, ByRef pictureCount, ByRef shapeCount, ByRef unavailableCount)
    Dim pairCount, i
    titleCount = 0: textCount = 0: pictureCount = 0: shapeCount = 0: unavailableCount = 0
    pairCount = Minimum(SafeCount(samplePres.Slides.Count), SafeCount(targetPres.Slides.Count))
    For i = 1 To pairCount
        CountSlideObjects samplePres.Slides(i), targetPres.Slides(i), selectedRules, titleCount, textCount, pictureCount, shapeCount, unavailableCount
    Next
    unavailableCount = unavailableCount + Abs(SafeCount(samplePres.Slides.Count) - SafeCount(targetPres.Slides.Count))
End Sub

Sub CountSlideObjects(sampleSlide, targetSlide, selectedRules, ByRef titleCount, ByRef textCount, ByRef pictureCount, ByRef shapeCount, ByRef unavailableCount)
    Dim sampleTitle, targetTitle, i, targetShape, categoryName, ordinal, sampleShape
    sampleTitle = GetTitleShapeName(sampleSlide): targetTitle = GetTitleShapeName(targetSlide)
    If HasRule(selectedRules, "标题") Then
        If Len(sampleTitle) > 0 And Len(targetTitle) > 0 Then titleCount = titleCount + 1 Else unavailableCount = unavailableCount + 1
    End If
    For i = 1 To SafeCount(targetSlide.Shapes.Count)
        Set targetShape = targetSlide.Shapes(i)
        categoryName = ShapeCategory(targetShape, targetTitle)
        If categoryName <> "title" And categoryName <> "unsupported" And HasRule(selectedRules, CategoryRule(categoryName)) Then
            ordinal = ShapeOrdinal(targetSlide, i, categoryName, targetTitle)
            Set sampleShape = FindShapeByCategoryAndOrdinal(sampleSlide, categoryName, ordinal, sampleTitle)
            If TypeName(sampleShape) = "Nothing" Then
                unavailableCount = unavailableCount + 1
            ElseIf categoryName = "text" Then
                textCount = textCount + 1
            ElseIf categoryName = "picture" Then
                pictureCount = pictureCount + 1
            Else
                shapeCount = shapeCount + 1
            End If
        ElseIf categoryName = "unsupported" Then
            unavailableCount = unavailableCount + 1
        End If
    Next
End Sub

Sub ApplyConfirmedObjects(samplePres, targetPres, selectedRules, ByRef changedTitle, ByRef changedText, ByRef changedPicture, ByRef changedShape, ByRef failureCount)
    Dim pairCount, i
    pairCount = Minimum(SafeCount(samplePres.Slides.Count), SafeCount(targetPres.Slides.Count))
    For i = 1 To pairCount
        ApplySlideObjects samplePres.Slides(i), targetPres.Slides(i), selectedRules, changedTitle, changedText, changedPicture, changedShape, failureCount
    Next
End Sub

Sub ApplySlideObjects(sampleSlide, targetSlide, selectedRules, ByRef changedTitle, ByRef changedText, ByRef changedPicture, ByRef changedShape, ByRef failureCount)
    Dim sampleTitleName, targetTitleName, sampleTitle, targetTitle
    sampleTitleName = GetTitleShapeName(sampleSlide): targetTitleName = GetTitleShapeName(targetSlide)
    If HasRule(selectedRules, "标题") And Len(sampleTitleName) > 0 And Len(targetTitleName) > 0 Then
        Set sampleTitle = sampleSlide.Shapes(sampleTitleName)
        Set targetTitle = targetSlide.Shapes(targetTitleName)
        If CopyTextShapeFormat(sampleTitle, targetTitle) Then changedTitle = changedTitle + 1 Else failureCount = failureCount + 1
    End If

    Dim i, targetShape, sampleShape, categoryName, ordinal
    For i = 1 To SafeCount(targetSlide.Shapes.Count)
        Set targetShape = targetSlide.Shapes(i)
        categoryName = ShapeCategory(targetShape, targetTitleName)
        If categoryName <> "title" And categoryName <> "unsupported" And HasRule(selectedRules, CategoryRule(categoryName)) Then
            ordinal = ShapeOrdinal(targetSlide, i, categoryName, targetTitleName)
            Set sampleShape = FindShapeByCategoryAndOrdinal(sampleSlide, categoryName, ordinal, sampleTitleName)
            If Not (TypeName(sampleShape) = "Nothing") Then
                If categoryName = "text" Then
                    If CopyTextShapeFormat(sampleShape, targetShape) Then changedText = changedText + 1 Else failureCount = failureCount + 1
                ElseIf categoryName = "picture" Then
                    If CopyGeometry(sampleShape, targetShape) Then changedPicture = changedPicture + 1 Else failureCount = failureCount + 1
                Else
                    If CopyBasicShapeFormat(sampleShape, targetShape) Then changedShape = changedShape + 1 Else failureCount = failureCount + 1
                End If
            End If
        End If
    Next
End Sub

Function CopyTextShapeFormat(sourceShape, targetShape)
    On Error Resume Next
    Dim geometryOk
    CopyTextShapeFormat = False
    geometryOk = CopyGeometry(sourceShape, targetShape)
    targetShape.Rotation = sourceShape.Rotation
    targetShape.Fill.ForeColor.RGB = sourceShape.Fill.ForeColor.RGB
    targetShape.Line.ForeColor.RGB = sourceShape.Line.ForeColor.RGB
    targetShape.Line.Weight = sourceShape.Line.Weight
    targetShape.TextFrame.MarginLeft = sourceShape.TextFrame.MarginLeft
    targetShape.TextFrame.MarginRight = sourceShape.TextFrame.MarginRight
    targetShape.TextFrame.MarginTop = sourceShape.TextFrame.MarginTop
    targetShape.TextFrame.MarginBottom = sourceShape.TextFrame.MarginBottom
    targetShape.TextFrame.TextRange.Font.Name = sourceShape.TextFrame.TextRange.Font.Name
    targetShape.TextFrame.TextRange.Font.Size = sourceShape.TextFrame.TextRange.Font.Size
    targetShape.TextFrame.TextRange.Font.Bold = sourceShape.TextFrame.TextRange.Font.Bold
    targetShape.TextFrame.TextRange.Font.Italic = sourceShape.TextFrame.TextRange.Font.Italic
    targetShape.TextFrame.TextRange.Font.Color.RGB = sourceShape.TextFrame.TextRange.Font.Color.RGB
    targetShape.TextFrame.TextRange.ParagraphFormat.Alignment = sourceShape.TextFrame.TextRange.ParagraphFormat.Alignment
    CopyTextShapeFormat = (geometryOk And Err.Number = 0)
    Err.Clear
End Function

Function CopyBasicShapeFormat(sourceShape, targetShape)
    On Error Resume Next
    Dim geometryOk
    CopyBasicShapeFormat = False
    geometryOk = CopyGeometry(sourceShape, targetShape)
    targetShape.Rotation = sourceShape.Rotation
    targetShape.Fill.ForeColor.RGB = sourceShape.Fill.ForeColor.RGB
    targetShape.Line.ForeColor.RGB = sourceShape.Line.ForeColor.RGB
    targetShape.Line.Weight = sourceShape.Line.Weight
    CopyBasicShapeFormat = (geometryOk And Err.Number = 0)
    Err.Clear
End Function

Function CopyGeometry(sourceShape, targetShape)
    On Error Resume Next
    CopyGeometry = False
    targetShape.Left = sourceShape.Left
    targetShape.Top = sourceShape.Top
    targetShape.Width = sourceShape.Width
    targetShape.Height = sourceShape.Height
    CopyGeometry = (Err.Number = 0)
    Err.Clear
End Function

Function GetTitleShapeName(slide)
    On Error Resume Next
    Dim titleShape
    GetTitleShapeName = ""
    Set titleShape = slide.Shapes.Title
    If Err.Number = 0 And Not titleShape Is Nothing Then GetTitleShapeName = CStr(titleShape.Name)
    Err.Clear
End Function

Function ShapeCategory(shape, titleName)
    On Error Resume Next
    ShapeCategory = "shape"
    If Len(titleName) > 0 And StrComp(CStr(shape.Name), titleName, vbTextCompare) = 0 Then
        ShapeCategory = "title"
    ElseIf IsUnsupportedShape(shape) Then
        ShapeCategory = "unsupported"
    ElseIf shape.Type = msoPicture Or shape.Type = msoLinkedPicture Then
        ShapeCategory = "picture"
    ElseIf shape.HasTextFrame = msoTrue Then
        ShapeCategory = "text"
    End If
    If Err.Number <> 0 Then ShapeCategory = "unsupported": Err.Clear
End Function

Function IsUnsupportedShape(shape)
    On Error Resume Next
    IsUnsupportedShape = (shape.Type = msoChart Or shape.Type = msoTable Or shape.Type = msoGroup Or shape.Type = msoSmartArt Or shape.Type = msoEmbeddedOLEObject Or shape.Type = msoLinkedOLEObject Or shape.Type = msoMedia)
    If Err.Number <> 0 Then IsUnsupportedShape = True: Err.Clear
End Function

Function ShapeOrdinal(slide, shapeIndex, categoryName, titleName)
    Dim i, count, currentShape
    count = 0
    For i = 1 To shapeIndex
        Set currentShape = slide.Shapes(i)
        If ShapeCategory(currentShape, titleName) = categoryName Then count = count + 1
    Next
    ShapeOrdinal = count
End Function

Function FindShapeByCategoryAndOrdinal(slide, categoryName, ordinal, titleName)
    On Error Resume Next
    Dim i, count, shape
    Set FindShapeByCategoryAndOrdinal = Nothing
    count = 0
    For i = 1 To SafeCount(slide.Shapes.Count)
        Set shape = slide.Shapes(i)
        If ShapeCategory(shape, titleName) = categoryName Then
            count = count + 1
            If count = ordinal Then Set FindShapeByCategoryAndOrdinal = shape: Exit Function
        End If
    Next
    Err.Clear
End Function

Function CategoryRule(categoryName)
    If categoryName = "text" Then
        CategoryRule = "文本"
    ElseIf categoryName = "picture" Then
        CategoryRule = "图片"
    Else
        CategoryRule = "形状"
    End If
End Function

Function NormalizeRules(value)
    Dim text
    text = Replace(Replace(Trim(CStr(value)), "，", ","), " ", "")
    NormalizeRules = text
End Function

Function HasRule(rules, ruleName)
    HasRule = (InStr(1, "," & rules & ",", "," & ruleName & ",", vbTextCompare) > 0)
End Function

Function HasAnySupportedRule(rules)
    HasAnySupportedRule = (HasRule(rules, "标题") Or HasRule(rules, "文本") Or HasRule(rules, "图片") Or HasRule(rules, "形状"))
End Function

Function PresentationTextFingerprint(pres)
    On Error Resume Next
    Dim slide, shape, text, i, codeValue, checksum, textLength
    checksum = 17
    textLength = 0
    For Each slide In pres.Slides
        For Each shape In slide.Shapes
            If shape.HasTextFrame = msoTrue And shape.TextFrame.HasText = msoTrue Then
                text = CStr(shape.TextFrame.TextRange.Text)
                textLength = textLength + Len(text)
                For i = 1 To Len(text)
                    codeValue = AscW(Mid(text, i, 1))
                    checksum = (checksum * 31 + codeValue + slide.SlideIndex + i) Mod 1000003
                Next
            End If
        Next
    Next
    PresentationTextFingerprint = CStr(textLength) & ":" & CStr(checksum)
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

Function BuildPlanParams(samplePath, selectedRules, titleCount, textCount, pictureCount, shapeCount, unavailableCount)
    BuildPlanParams = "{""samplePath"":""" & EscapeJson(samplePath) & """,""rules"":""" & EscapeJson(selectedRules) & _
        """,""title"":" & CStr(titleCount) & ",""text"":" & CStr(textCount) & ",""picture"":" & CStr(pictureCount) & _
        ",""shape"":" & CStr(shapeCount) & ",""unavailable"":" & CStr(unavailableCount) & "}"
End Function

Function PreviewJson(samplePath, selectedRules, titleCount, textCount, pictureCount, shapeCount, unavailableCount, planPreview, planValidation, message)
    PreviewJson = "{""ok"":true,""previewOnly"":true,""samplePath"":""" & EscapeJson(samplePath) & _
        """,""rules"":""" & EscapeJson(selectedRules) & """,""impact"":{""title"":" & CStr(titleCount) & _
        ",""text"":" & CStr(textCount) & ",""picture"":" & CStr(pictureCount) & ",""shape"":" & CStr(shapeCount) & _
        ",""unavailable"":" & CStr(unavailableCount) & "},""planPreview"":""" & EscapeJson(planPreview) & _
        """,""planValidation"":""" & EscapeJson(planValidation) & """,""requiresConfirmation"":true,""message"":""" & EscapeJson(message) & """}"
End Function

Function FailureJson(code, message)
    FailureJson = "{""ok"":false,""code"":""" & EscapeJson(code) & """,""message"":""" & EscapeJson(message) & """}"
End Function

Function RollbackJson(code, message, reloadOk, contentPreserved, failureCount, backupPath)
    RollbackJson = "{""ok"":false,""code"":""" & EscapeJson(code) & """,""rollbackReloaded"":" & JsonBool(reloadOk) & _
        ",""contentPreserved"":" & JsonBool(contentPreserved) & ",""failureCount"":" & CStr(failureCount) & _
        ",""backupPath"":""" & EscapeJson(backupPath) & """,""message"":""" & EscapeJson(message) & """}"
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
