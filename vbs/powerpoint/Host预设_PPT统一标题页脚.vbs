' 函数名: HostPptNormalizeTitleFooter
' 描述: 统一 PowerPoint 标题字体并写入页脚文本；验证 Host.GetApplication、Host.Prompt、PPT 上下文、写入计划和日志能力
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
        Main = "{""ok"":false,""code"":""E_NO_PPT_APP"",""message"":""未取得 PowerPoint 应用，请在 PowerPoint 中运行该预设""}"
        Exit Function
    End If

    Main = HostPptNormalizeTitleFooter(appObj)
End Function

Function HostPptNormalizeTitleFooter(appObj)
    On Error Resume Next

    Dim pres
    Set pres = appObj.ActivePresentation
    If Err.Number <> 0 Or TypeName(pres) = "Empty" Or TypeName(pres) = "Nothing" Then
        Err.Clear
        HostPptNormalizeTitleFooter = "{""ok"":false,""code"":""E_NO_PRESENTATION"",""message"":""当前没有活动演示文稿""}"
        Exit Function
    End If

    Dim footerText
    footerText = Trim(SafePrompt("请输入页脚文本，留空则只统一标题样式。", ""))

    Dim planId, previewJson
    planId = SafeBeginWritePlan()
    SafeRecordWrite "ppt_title_footer_normalize", "{""scope"":""activePresentation""}", "office.ppt.format"
    previewJson = SafePreviewWritePlan()

    Dim slide, titleCount, footerCount
    titleCount = 0
    footerCount = 0

    For Each slide In pres.Slides
        If NormalizeSlideTitle(slide) Then titleCount = titleCount + 1
        If Len(footerText) > 0 Then
            If ApplySlideFooter(slide, footerText) Then footerCount = footerCount + 1
        End If
    Next

    SafeCloseWritePlan planId

    Dim pptContext, summary
    pptContext = SafeHostText("GetPptContextInfo")
    summary = "PPT 标题/页脚统一完成；标题=" & CStr(titleCount) & "，页脚=" & CStr(footerCount) & vbCrLf & "Host.GetPptContextInfo: " & pptContext
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostPptNormalizeTitleFooter = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""titleCount"":" & CStr(titleCount) & _
        ",""footerCount"":" & CStr(footerCount) & ",""planId"":""" & EscapeJson(planId) & _
        """,""preview"":""" & EscapeJson(previewJson) & """}"

End Function

Function NormalizeSlideTitle(slide)
    On Error Resume Next
    NormalizeSlideTitle = False
    Dim titleShape
    Set titleShape = Nothing
    Set titleShape = slide.Shapes.Title
    If Err.Number <> 0 Or TypeName(titleShape) = "Empty" Or TypeName(titleShape) = "Nothing" Then
        Err.Clear
        Exit Function
    End If
    If titleShape.HasTextFrame = msoTrue Then
        titleShape.TextFrame.TextRange.Font.NameFarEast = "微软雅黑"
        titleShape.TextFrame.TextRange.Font.Name = "Arial"
        titleShape.TextFrame.TextRange.Font.Size = 28
        titleShape.TextFrame.TextRange.Font.Bold = msoTrue
        NormalizeSlideTitle = True
    End If
End Function

Function ApplySlideFooter(slide, footerText)
    On Error Resume Next
    ApplySlideFooter = False
    slide.HeadersFooters.Footer.Visible = msoTrue
    slide.HeadersFooters.Footer.Text = footerText
    If Err.Number = 0 Then
        ApplySlideFooter = True
    Else
        Err.Clear
    End If
End Function

Function SafePrompt(message, defaultValue)
    On Error Resume Next
    SafePrompt = Host.Prompt(message, defaultValue)
    If Err.Number <> 0 Then
        SafePrompt = defaultValue
        Err.Clear
    End If
End Function

Function SafeHostText(methodName)
    On Error Resume Next
    Select Case methodName
        Case "GetPptContextInfo"
            SafeHostText = Host.GetPptContextInfo()
        Case Else
            SafeHostText = ""
    End Select
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
