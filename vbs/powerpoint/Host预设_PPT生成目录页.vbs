' 函数名: HostPptBuildAgendaSlide
' 描述: 从当前 PowerPoint 演示文稿各页标题生成目录页，并插入到第一页；适合汇报、培训和方案文档快速生成结构页
' 适用应用: PowerPoint
' 搜索范围: 全文
' 搜索对象: 无

Option Explicit

Const ppLayoutTitleOnly = 11
Const msoTrue = -1

Function Main()
    On Error Resume Next

    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = "{""ok"":false,""code"":""E_NO_PPT_APP"",""message"":""未取得 PowerPoint 应用，请在 PowerPoint 中运行该预设""}"
        Exit Function
    End If

    Main = HostPptBuildAgendaSlide(appObj)
End Function

Function HostPptBuildAgendaSlide(appObj)
    On Error Resume Next

    Dim pres
    Set pres = appObj.ActivePresentation
    If Err.Number <> 0 Or TypeName(pres) = "Empty" Or TypeName(pres) = "Nothing" Then
        Err.Clear
        HostPptBuildAgendaSlide = "{""ok"":false,""code"":""E_NO_PRESENTATION"",""message"":""当前没有活动演示文稿""}"
        Exit Function
    End If
    If pres.Slides.Count < 2 Then
        HostPptBuildAgendaSlide = "{""ok"":false,""code"":""E_TOO_FEW_SLIDES"",""message"":""至少需要两页幻灯片才适合生成目录页""}"
        Exit Function
    End If

    Dim slideSummary, planId, previewJson
    slideSummary = SafeHostText("GetPptSlideSummary")
    planId = SafeBeginWritePlan()
    SafeRecordWrite "ppt_build_agenda_slide", "{""scope"":""activePresentation""}", "office.ppt.structure"
    previewJson = SafePreviewWritePlan()

    Dim agendaSlide
    Set agendaSlide = pres.Slides.Add(1, ppLayoutTitleOnly)
    SetTitle agendaSlide, "目录"

    Dim bodyShape, agendaText, itemCount
    agendaText = BuildAgendaText(pres, itemCount)
    Set bodyShape = agendaSlide.Shapes.AddTextbox(1, 72, 120, pres.PageSetup.SlideWidth - 144, pres.PageSetup.SlideHeight - 180)
    bodyShape.TextFrame.TextRange.Text = agendaText
    bodyShape.TextFrame.TextRange.Font.NameFarEast = "微软雅黑"
    bodyShape.TextFrame.TextRange.Font.Name = "Arial"
    bodyShape.TextFrame.TextRange.Font.Size = 20
    bodyShape.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 8
    SafeCloseWritePlan planId

    Dim summary
    summary = "PPT 目录页生成完成；目录项=" & CStr(itemCount) & "；总页数=" & CStr(pres.Slides.Count) & vbCrLf & _
        "Host.GetPptSlideSummary: " & slideSummary & vbCrLf & _
        "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostPptBuildAgendaSlide = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""agendaItems"":" & CStr(itemCount) & _
        ",""slideCount"":" & CStr(pres.Slides.Count) & ",""planId"":""" & EscapeJson(planId) & """}"

End Function

Sub SetTitle(slide, titleText)
    On Error Resume Next
    Dim titleShape
    Set titleShape = slide.Shapes.Title
    If Err.Number <> 0 Or TypeName(titleShape) = "Empty" Or TypeName(titleShape) = "Nothing" Then
        Err.Clear
        Set titleShape = slide.Shapes.AddTextbox(1, 72, 36, 600, 60)
    End If
    titleShape.TextFrame.TextRange.Text = titleText
    titleShape.TextFrame.TextRange.Font.NameFarEast = "微软雅黑"
    titleShape.TextFrame.TextRange.Font.Name = "Arial"
    titleShape.TextFrame.TextRange.Font.Size = 34
    titleShape.TextFrame.TextRange.Font.Bold = msoTrue
End Sub

Function BuildAgendaText(pres, ByRef itemCount)
    On Error Resume Next
    Dim i, titleText, result
    result = ""
    itemCount = 0
    For i = 2 To pres.Slides.Count
        titleText = Trim(GetSlideTitle(pres.Slides(i)))
        If Len(titleText) = 0 Then titleText = "第 " & CStr(i - 1) & " 部分"
        result = result & CStr(i - 1) & ". " & titleText & vbCrLf
        itemCount = itemCount + 1
    Next
    BuildAgendaText = result
End Function

Function GetSlideTitle(slide)
    On Error Resume Next
    Dim titleShape
    GetSlideTitle = ""
    Set titleShape = slide.Shapes.Title
    If Err.Number = 0 Then
        If titleShape.HasTextFrame = msoTrue Then
            GetSlideTitle = titleShape.TextFrame.TextRange.Text
        End If
    End If
    Err.Clear
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
    text = Replace(text, vbTab, "\t")
    EscapeJson = text
End Function
