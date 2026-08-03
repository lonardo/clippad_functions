' 函数名: HostPptPreflightMasterFooterPageNumberMapping
' 描述: 只读比较用户选择的 PowerPoint 样本与当前演示文稿的页面尺寸、母版/版式数量、页脚和页码规则；不应用母版或修改任何内容
' 适用应用: PowerPoint
' 搜索范围: 全文
' 搜索对象: 无

Option Explicit

Const msoTrue = -1

Function Main()
    On Error Resume Next
    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = FailureJson("E_NO_PPT_APP", "未取得 PowerPoint 应用，请在 PowerPoint 中运行该预设")
        Exit Function
    End If
    Main = HostPptPreflightMasterFooterPageNumberMapping(appObj)
End Function

Function HostPptPreflightMasterFooterPageNumberMapping(appObj)
    On Error Resume Next
    Dim targetPres
    Set targetPres = appObj.ActivePresentation
    If Err.Number <> 0 Or TypeName(targetPres) = "Empty" Or TypeName(targetPres) = "Nothing" Then
        Err.Clear
        HostPptPreflightMasterFooterPageNumberMapping = FailureJson("E_NO_PRESENTATION", "当前没有活动演示文稿")
        Exit Function
    End If

    Dim samplePath
    samplePath = Trim(SafeSelectFile("选择规范 PowerPoint 样本（本次只预检页面、母版、页脚和页码规则）", "PowerPoint 演示文稿|*.ppt;*.pptx;*.pptm;*.ppsx"))
    If Len(samplePath) = 0 Then
        HostPptPreflightMasterFooterPageNumberMapping = FailureJson("E_SAMPLE_REQUIRED", "未选择样本演示文稿；未修改当前演示文稿")
        Exit Function
    End If
    If IsSamePresentation(targetPres, samplePath) Then
        HostPptPreflightMasterFooterPageNumberMapping = FailureJson("E_SAMPLE_IS_TARGET", "样本不能与当前目标演示文稿相同")
        Exit Function
    End If

    Dim targetFingerprint
    targetFingerprint = PresentationTextFingerprint(targetPres)
    Dim samplePres
    Set samplePres = Nothing
    Err.Clear
    Set samplePres = appObj.Presentations.Open(samplePath, True, False, False)
    If Err.Number <> 0 Or TypeName(samplePres) = "Empty" Or TypeName(samplePres) = "Nothing" Then
        Dim openError
        openError = Err.Description
        Err.Clear
        HostPptPreflightMasterFooterPageNumberMapping = FailureJson("E_SAMPLE_OPEN", "无法以只读方式打开样本演示文稿：" & openError)
        Exit Function
    End If

    Dim sampleFingerprint
    sampleFingerprint = PresentationTextFingerprint(samplePres)

    Dim sampleSlides, targetSlides, sampleDesigns, targetDesigns, sampleMasters, targetMasters
    Dim sampleLayouts, targetLayouts, sampleFooterVisible, targetFooterVisible
    Dim sampleFooterText, targetFooterText, sampleSlideNumberVisible, targetSlideNumberVisible
    Dim sampleDateVisible, targetDateVisible
    sampleSlides = SafeCount(samplePres.Slides.Count)
    targetSlides = SafeCount(targetPres.Slides.Count)
    sampleDesigns = SafeDesignCount(samplePres)
    targetDesigns = SafeDesignCount(targetPres)
    sampleMasters = SafeMasterCount(samplePres)
    targetMasters = SafeMasterCount(targetPres)
    sampleLayouts = SafeLayoutCount(samplePres)
    targetLayouts = SafeLayoutCount(targetPres)
    CountFooterRules samplePres, sampleFooterVisible, sampleFooterText, sampleSlideNumberVisible, sampleDateVisible
    CountFooterRules targetPres, targetFooterVisible, targetFooterText, targetSlideNumberVisible, targetDateVisible

    Dim sameSlideSize, footerStatus, pageNumberStatus, masterStatus
    sameSlideSize = SameSlideSize(samplePres, targetPres)
    footerStatus = MappingStatus(sampleFooterVisible, targetFooterVisible, sameSlideSize)
    pageNumberStatus = MappingStatus(sampleSlideNumberVisible, targetSlideNumberVisible, sameSlideSize)
    masterStatus = "blocked"

    Dim report, contextSummary
    report = "PPT 母版/页脚/页码映射预检（只读）" & vbCrLf & String(34, "=") & vbCrLf & _
        "页面尺寸一致=" & BoolLabel(sameSlideSize) & "；样本页=" & CStr(sampleSlides) & "；目标页=" & CStr(targetSlides) & vbCrLf & _
        "样本设计=" & CStr(sampleDesigns) & "；目标设计=" & CStr(targetDesigns) & _
        "；样本母版=" & CStr(sampleMasters) & "；目标母版=" & CStr(targetMasters) & _
        "；样本自定义版式=" & CStr(sampleLayouts) & "；目标自定义版式=" & CStr(targetLayouts) & vbCrLf & _
        "页脚：样本可见页=" & CStr(sampleFooterVisible) & "；目标可见页=" & CStr(targetFooterVisible) & _
        "；样本含文字页=" & CStr(sampleFooterText) & "；目标含文字页=" & CStr(targetFooterText) & _
        "；状态=" & footerStatus & vbCrLf & _
        "页码：样本可见页=" & CStr(sampleSlideNumberVisible) & "；目标可见页=" & CStr(targetSlideNumberVisible) & _
        "；状态=" & pageNumberStatus & vbCrLf & _
        "日期：样本可见页=" & CStr(sampleDateVisible) & "；目标可见页=" & CStr(targetDateVisible) & vbCrLf & _
        "母版/主题：状态=blocked；首版不复制设计、母版、主题、版式 ID、背景或占位符层级。" & vbCrLf & _
        "后续写入只允许按页应用已确认的页脚可见性、页码可见性和页脚格式；页脚文字仍需用户明确确认。"
    contextSummary = SafeHostText("GetPptContextInfo")
    Host.WriteClipboard report & vbCrLf & "Host.GetPptContextInfo: " & contextSummary
    SafeWriteLog "PPT master/footer/page-number mapping preflight completed; masterMapping=blocked"

    Dim samplePreserved, targetPreserved
    samplePreserved = (PresentationTextFingerprint(samplePres) = sampleFingerprint)
    targetPreserved = (PresentationTextFingerprint(targetPres) = targetFingerprint)
    SafeClosePresentation samplePres

    HostPptPreflightMasterFooterPageNumberMapping = "{""ok"":true,""readOnly"":true,""samplePath"":""" & EscapeJson(samplePath) & _
        """,""sameSlideSize"":" & JsonBool(sameSlideSize) & ",""sampleSlides"":" & CStr(sampleSlides) & _
        ",""targetSlides"":" & CStr(targetSlides) & ",""sampleDesigns"":" & CStr(sampleDesigns) & _
        ",""targetDesigns"":" & CStr(targetDesigns) & ",""sampleMasters"":" & CStr(sampleMasters) & _
        ",""targetMasters"":" & CStr(targetMasters) & ",""sampleLayouts"":" & CStr(sampleLayouts) & _
        ",""targetLayouts"":" & CStr(targetLayouts) & ",""sampleFooterVisible"":" & CStr(sampleFooterVisible) & _
        ",""targetFooterVisible"":" & CStr(targetFooterVisible) & ",""sampleSlideNumberVisible"":" & CStr(sampleSlideNumberVisible) & _
        ",""targetSlideNumberVisible"":" & CStr(targetSlideNumberVisible) & ",""footerMapping"":""" & footerStatus & _
        """,""slideNumberMapping"":""" & pageNumberStatus & """,""masterMapping"":""" & masterStatus & _
        """,""sampleContentPreserved"":" & JsonBool(samplePreserved) & ",""targetContentPreserved"":" & JsonBool(targetPreserved) & _
        ",""requiresPreview"":true,""requiresConfirmation"":true,""message"":""" & EscapeJson(report) & """}"
End Function

Sub CountFooterRules(pres, ByRef footerVisible, ByRef footerTextCount, ByRef slideNumberVisible, ByRef dateVisible)
    On Error Resume Next
    Dim slide, footerText
    footerVisible = 0
    footerTextCount = 0
    slideNumberVisible = 0
    dateVisible = 0
    For Each slide In pres.Slides
        If IsTrue(slide.HeadersFooters.Footer.Visible) Then footerVisible = footerVisible + 1
        footerText = ""
        footerText = CStr(slide.HeadersFooters.Footer.Text)
        If Len(Trim(footerText)) > 0 Then footerTextCount = footerTextCount + 1
        If IsTrue(slide.HeadersFooters.SlideNumber.Visible) Then slideNumberVisible = slideNumberVisible + 1
        If IsTrue(slide.HeadersFooters.DateAndTime.Visible) Then dateVisible = dateVisible + 1
        If Err.Number <> 0 Then Err.Clear
    Next
End Sub

Function SameSlideSize(samplePres, targetPres)
    On Error Resume Next
    SameSlideSize = (Abs(CDbl(samplePres.PageSetup.SlideWidth) - CDbl(targetPres.PageSetup.SlideWidth)) < 0.1 And _
        Abs(CDbl(samplePres.PageSetup.SlideHeight) - CDbl(targetPres.PageSetup.SlideHeight)) < 0.1)
    If Err.Number <> 0 Then SameSlideSize = False
    Err.Clear
End Function

Function SafeDesignCount(pres)
    On Error Resume Next
    SafeDesignCount = CLng(pres.Designs.Count)
    If Err.Number <> 0 Then SafeDesignCount = 0
    Err.Clear
End Function

Function SafeMasterCount(pres)
    On Error Resume Next
    SafeMasterCount = CLng(pres.Designs.Count)
    If Err.Number <> 0 Then SafeMasterCount = 0
    Err.Clear
End Function

Function SafeLayoutCount(pres)
    On Error Resume Next
    SafeLayoutCount = CLng(pres.SlideMaster.CustomLayouts.Count)
    If Err.Number <> 0 Then SafeLayoutCount = 0
    Err.Clear
End Function

Function MappingStatus(sampleVisible, targetVisible, sameSlideSize)
    If sampleVisible <= 0 Then
        MappingStatus = "not_applicable"
    ElseIf Not sameSlideSize Then
        MappingStatus = "needs_confirmation_size_mismatch"
    ElseIf sampleVisible = targetVisible Then
        MappingStatus = "candidate_same_coverage"
    Else
        MappingStatus = "candidate_coverage_differs"
    End If
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

Function IsTrue(value)
    On Error Resume Next
    IsTrue = (CLng(value) = msoTrue)
    If Err.Number <> 0 Then IsTrue = False
    Err.Clear
End Function

Function BoolLabel(value)
    If value Then BoolLabel = "是" Else BoolLabel = "否"
End Function

Function IsSamePresentation(targetPres, samplePath)
    On Error Resume Next
    IsSamePresentation = (StrComp(CStr(targetPres.FullName), samplePath, vbTextCompare) = 0)
    If Err.Number <> 0 Then IsSamePresentation = False
    Err.Clear
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

Function SafeHostText(methodName)
    On Error Resume Next
    If methodName = "GetPptContextInfo" Then
        SafeHostText = Host.GetPptContextInfo()
    Else
        SafeHostText = ""
    End If
    If Err.Number <> 0 Then SafeHostText = ""
    Err.Clear
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

Function FailureJson(code, message)
    FailureJson = "{""ok"":false,""code"":""" & EscapeJson(code) & """,""message"":""" & EscapeJson(message) & """}"
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
