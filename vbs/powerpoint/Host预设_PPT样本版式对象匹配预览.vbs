' 函数名: HostPptPreviewSampleLayoutMapping
' 描述: 只读打开用户选择的 PowerPoint 样本，按页面与对象类别生成可映射、需确认和不可映射清单；不修改样本或当前演示文稿
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
Const MaxDetailLines = 120

Function Main()
    On Error Resume Next

    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = "{""ok"":false,""code"":""E_NO_PPT_APP"",""message"":""未取得 PowerPoint 应用，请在 PowerPoint 中运行该预设""}"
        Exit Function
    End If

    Main = HostPptPreviewSampleLayoutMapping(appObj)
End Function

Function HostPptPreviewSampleLayoutMapping(appObj)
    On Error Resume Next

    Dim targetPres
    Set targetPres = appObj.ActivePresentation
    If Err.Number <> 0 Or TypeName(targetPres) = "Empty" Or TypeName(targetPres) = "Nothing" Then
        Err.Clear
        HostPptPreviewSampleLayoutMapping = "{""ok"":false,""code"":""E_NO_PRESENTATION"",""message"":""当前没有活动演示文稿""}"
        Exit Function
    End If

    Dim samplePath
    samplePath = Trim(SafeSelectFile("选择规范 PowerPoint 样本（本次只生成对象匹配预览，不会修改任何演示文稿）", "PowerPoint 演示文稿|*.ppt;*.pptx;*.pptm;*.ppsx"))
    If Len(samplePath) = 0 Then
        HostPptPreviewSampleLayoutMapping = "{""ok"":false,""code"":""E_SAMPLE_REQUIRED"",""message"":""未选择样本演示文稿；未修改当前演示文稿""}"
        Exit Function
    End If
    If IsSamePresentation(targetPres, samplePath) Then
        HostPptPreviewSampleLayoutMapping = "{""ok"":false,""code"":""E_SAMPLE_IS_TARGET"",""message"":""样本不能与当前目标演示文稿相同；请选择独立的规范样本""}"
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
        HostPptPreviewSampleLayoutMapping = "{""ok"":false,""code"":""E_SAMPLE_OPEN"",""message"":""无法以只读方式打开样本演示文稿：" & EscapeJson(openError) & """}"
        Exit Function
    End If

    Dim report, detailLines, detailTruncated, mappableCount, confirmationCount, unmappableCount, sourceOnlyCount
    report = "PPT 样本版式对象匹配预览（只读）" & vbCrLf & String(32, "=") & vbCrLf & _
        "样本=" & samplePres.Name & "；目标=" & targetPres.Name & vbCrLf & _
        "映射规则：仅比较同页序号和对象类别；不复制文字、图片、图表数据或动画。" & vbCrLf
    detailLines = 0
    detailTruncated = False
    mappableCount = 0
    confirmationCount = 0
    unmappableCount = 0
    sourceOnlyCount = 0

    Dim pairCount, i
    pairCount = Minimum(SafeCount(samplePres.Slides.Count), SafeCount(targetPres.Slides.Count))
    For i = 1 To pairCount
        CompareSlidePair samplePres.Slides(i), targetPres.Slides(i), report, detailLines, detailTruncated, mappableCount, confirmationCount, unmappableCount, sourceOnlyCount
    Next
    For i = pairCount + 1 To SafeCount(targetPres.Slides.Count)
        unmappableCount = unmappableCount + 1
        AppendDetail report, detailLines, detailTruncated, "目标第 " & CStr(i) & " 页：不可映射（样本缺少对应页面）"
    Next
    For i = pairCount + 1 To SafeCount(samplePres.Slides.Count)
        sourceOnlyCount = sourceOnlyCount + 1
        AppendDetail report, detailLines, detailTruncated, "样本第 " & CStr(i) & " 页：仅样本存在（不会创建或复制目标页面）"
    Next
    If detailTruncated Then report = report & "明细已限制为前 " & CStr(MaxDetailLines) & " 项；请缩小样本或目标范围后重试。" & vbCrLf

    Dim slideSummary, summary
    slideSummary = SafeHostText("GetPptSlideSummary")
    summary = "PPT 样本版式对象匹配预览完成（只读）；样本页=" & CStr(SafeCount(samplePres.Slides.Count)) & _
        "；目标页=" & CStr(SafeCount(targetPres.Slides.Count)) & "；可映射=" & CStr(mappableCount) & _
        "；需确认=" & CStr(confirmationCount) & "；不可映射=" & CStr(unmappableCount) & _
        "；仅样本=" & CStr(sourceOnlyCount) & vbCrLf & _
        "Host.GetPptSlideSummary: " & slideSummary
    Host.WriteClipboard summary & vbCrLf & report
    SafeWriteLog summary

    SafeClosePresentation samplePres
    HostPptPreviewSampleLayoutMapping = "{""ok"":true,""readOnly"":true,""samplePath"":""" & EscapeJson(samplePath) & _
        """,""mappableCount"":" & CStr(mappableCount) & ",""confirmationCount"":" & CStr(confirmationCount) & _
        ",""unmappableCount"":" & CStr(unmappableCount) & ",""sourceOnlyCount"":" & CStr(sourceOnlyCount) & _
        ",""detailTruncated"":" & JsonBool(detailTruncated) & ",""requiresPreview"":true,""requiresConfirmation"":true,""message"":""" & EscapeJson(summary) & """}"
End Function

Sub CompareSlidePair(sampleSlide, targetSlide, ByRef report, ByRef detailLines, ByRef detailTruncated, ByRef mappableCount, ByRef confirmationCount, ByRef unmappableCount, ByRef sourceOnlyCount)
    Dim sampleTitleName, targetTitleName
    sampleTitleName = GetTitleShapeName(sampleSlide)
    targetTitleName = GetTitleShapeName(targetSlide)
    If Len(sampleTitleName) > 0 And Len(targetTitleName) > 0 Then
        mappableCount = mappableCount + 1
        AppendDetail report, detailLines, detailTruncated, "第 " & CStr(targetSlide.SlideIndex) & " 页标题：可映射（标题格式、位置和尺寸；不复制文字）"
    ElseIf Len(sampleTitleName) > 0 Then
        unmappableCount = unmappableCount + 1
        AppendDetail report, detailLines, detailTruncated, "第 " & CStr(targetSlide.SlideIndex) & " 页标题：不可映射（目标页没有标题占位符）"
    ElseIf Len(targetTitleName) > 0 Then
        AppendDetail report, detailLines, detailTruncated, "第 " & CStr(targetSlide.SlideIndex) & " 页标题：需确认（样本页没有标题占位符，保持目标标题不变）"
        confirmationCount = confirmationCount + 1
    End If

    Dim targetIndex, targetShape, categoryName, ordinal, sampleShape
    For targetIndex = 1 To SafeCount(targetSlide.Shapes.Count)
        Set targetShape = targetSlide.Shapes(targetIndex)
        categoryName = ShapeCategory(targetShape, targetTitleName)
        If categoryName = "title" Then
        ElseIf categoryName = "unsupported" Then
            unmappableCount = unmappableCount + 1
            AppendDetail report, detailLines, detailTruncated, ShapeLabel(targetSlide, targetShape) & "：不可映射（" & UnsupportedReason(targetShape) & "）"
        Else
            ordinal = ShapeOrdinal(targetSlide, targetIndex, categoryName, targetTitleName)
            Set sampleShape = FindShapeByCategoryAndOrdinal(sampleSlide, categoryName, ordinal, sampleTitleName)
            If TypeName(sampleShape) = "Nothing" Then
                unmappableCount = unmappableCount + 1
                AppendDetail report, detailLines, detailTruncated, ShapeLabel(targetSlide, targetShape) & "：不可映射（样本同类对象不足）"
            Else
                confirmationCount = confirmationCount + 1
                AppendDetail report, detailLines, detailTruncated, ShapeLabel(targetSlide, targetShape) & "：需确认（与样本第 " & CStr(sampleSlide.SlideIndex) & " 页第 " & CStr(ordinal) & " 个" & CategoryLabel(categoryName) & "候选匹配；只可迁移格式与几何属性）"
            End If
        End If
        If Err.Number <> 0 Then Err.Clear
    Next

    AddSampleOnlyObjects sampleSlide, targetSlide, sampleTitleName, targetTitleName, "text", report, detailLines, detailTruncated, sourceOnlyCount
    AddSampleOnlyObjects sampleSlide, targetSlide, sampleTitleName, targetTitleName, "picture", report, detailLines, detailTruncated, sourceOnlyCount
    AddSampleOnlyObjects sampleSlide, targetSlide, sampleTitleName, targetTitleName, "shape", report, detailLines, detailTruncated, sourceOnlyCount
End Sub

Sub AddSampleOnlyObjects(sampleSlide, targetSlide, sampleTitleName, targetTitleName, categoryName, ByRef report, ByRef detailLines, ByRef detailTruncated, ByRef sourceOnlyCount)
    Dim sampleCount, targetCount, i
    sampleCount = CountShapesByCategory(sampleSlide, categoryName, sampleTitleName)
    targetCount = CountShapesByCategory(targetSlide, categoryName, targetTitleName)
    For i = targetCount + 1 To sampleCount
        sourceOnlyCount = sourceOnlyCount + 1
        AppendDetail report, detailLines, detailTruncated, "样本第 " & CStr(sampleSlide.SlideIndex) & " 页第 " & CStr(i) & " 个" & CategoryLabel(categoryName) & "：仅样本存在（不会自动创建目标对象）"
    Next
End Sub

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
    If Err.Number <> 0 Then
        ShapeCategory = "unsupported"
        Err.Clear
    End If
End Function

Function IsUnsupportedShape(shape)
    On Error Resume Next
    IsUnsupportedShape = (shape.Type = msoChart Or shape.Type = msoTable Or shape.Type = msoGroup Or _
        shape.Type = msoSmartArt Or shape.Type = msoEmbeddedOLEObject Or shape.Type = msoLinkedOLEObject Or shape.Type = msoMedia)
    If Err.Number <> 0 Then
        IsUnsupportedShape = True
        Err.Clear
    End If
End Function

Function UnsupportedReason(shape)
    On Error Resume Next
    UnsupportedReason = "宿主专有或复合对象"
    If shape.Type = msoChart Then
        UnsupportedReason = "图表数据与布局"
    ElseIf shape.Type = msoTable Then
        UnsupportedReason = "表格内容与单元格格式"
    ElseIf shape.Type = msoGroup Then
        UnsupportedReason = "组合对象层级"
    ElseIf shape.Type = msoSmartArt Then
        UnsupportedReason = "SmartArt 结构"
    ElseIf shape.Type = msoMedia Then
        UnsupportedReason = "媒体对象"
    ElseIf shape.Type = msoEmbeddedOLEObject Or shape.Type = msoLinkedOLEObject Then
        UnsupportedReason = "嵌入或链接对象"
    End If
    If Err.Number <> 0 Then Err.Clear
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
            If count = ordinal Then
                Set FindShapeByCategoryAndOrdinal = shape
                Exit Function
            End If
        End If
    Next
    Err.Clear
End Function

Function CountShapesByCategory(slide, categoryName, titleName)
    Dim i, count, shape
    count = 0
    For i = 1 To SafeCount(slide.Shapes.Count)
        Set shape = slide.Shapes(i)
        If ShapeCategory(shape, titleName) = categoryName Then count = count + 1
    Next
    CountShapesByCategory = count
End Function

Function ShapeLabel(slide, shape)
    On Error Resume Next
    ShapeLabel = "第 " & CStr(slide.SlideIndex) & " 页对象 " & CStr(shape.Name)
    If Err.Number <> 0 Then
        ShapeLabel = "当前页对象"
        Err.Clear
    End If
End Function

Function CategoryLabel(categoryName)
    If categoryName = "text" Then
        CategoryLabel = "文本对象"
    ElseIf categoryName = "picture" Then
        CategoryLabel = "图片对象"
    Else
        CategoryLabel = "普通形状"
    End If
End Function

Sub AppendDetail(ByRef report, ByRef detailLines, ByRef detailTruncated, detail)
    If detailLines < MaxDetailLines Then
        report = report & detail & vbCrLf
        detailLines = detailLines + 1
    Else
        detailTruncated = True
    End If
End Sub

Function IsSamePresentation(targetPres, samplePath)
    On Error Resume Next
    Dim targetPath
    IsSamePresentation = False
    targetPath = CStr(targetPres.FullName)
    If Err.Number = 0 And Len(targetPath) > 0 Then IsSamePresentation = (StrComp(targetPath, samplePath, vbTextCompare) = 0)
    Err.Clear
End Function

Function Minimum(leftValue, rightValue)
    If leftValue < rightValue Then
        Minimum = leftValue
    Else
        Minimum = rightValue
    End If
End Function

Function SafeCount(value)
    On Error Resume Next
    SafeCount = CLng(value)
    If Err.Number <> 0 Then
        SafeCount = 0
        Err.Clear
    End If
End Function

Function SafeSelectFile(title, filter)
    On Error Resume Next
    SafeSelectFile = Host.SelectFile(title, filter)
    If Err.Number <> 0 Then
        SafeSelectFile = ""
        Err.Clear
    End If
End Function

Function SafeHostText(methodName)
    On Error Resume Next
    If methodName = "GetPptSlideSummary" Then
        SafeHostText = Host.GetPptSlideSummary()
    Else
        SafeHostText = ""
    End If
    If Err.Number <> 0 Then
        SafeHostText = ""
        Err.Clear
    End If
End Function

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

Function JsonBool(value)
    If value Then
        JsonBool = "true"
    Else
        JsonBool = "false"
    End If
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
