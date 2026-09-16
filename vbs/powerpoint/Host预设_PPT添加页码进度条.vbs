' 函数名: HostPptAddPageNumberProgress
' 描述: 给当前 PowerPoint 演示文稿每页添加页码和底部进度条；适合路演、培训课件和长文档汇报
' 适用应用: PowerPoint
' 搜索范围: 全文
' 搜索对象: 无

Option Explicit

Const msoTrue = -1
Const msoFalse = 0

Function Main()
    On Error Resume Next

    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = "{""ok"":false,""code"":""E_NO_PPT_APP"",""message"":""未取得 PowerPoint 应用，请在 PowerPoint 中运行该预设""}"
        Exit Function
    End If

    Main = HostPptAddPageNumberProgress(appObj)
End Function

Function HostPptAddPageNumberProgress(appObj)
    On Error Resume Next

    Dim pres
    Set pres = appObj.ActivePresentation
    If Err.Number <> 0 Or TypeName(pres) = "Empty" Or TypeName(pres) = "Nothing" Then
        Err.Clear
        HostPptAddPageNumberProgress = "{""ok"":false,""code"":""E_NO_PRESENTATION"",""message"":""当前没有活动演示文稿""}"
        Exit Function
    End If

    Dim planId, previewJson, pptContext
    pptContext = SafeHostText("GetPptContextInfo")
    planId = SafeBeginWritePlan()
    SafeRecordWrite "ppt_add_page_number_progress", "{""scope"":""activePresentation""}", "office.ppt.decorate"
    previewJson = SafePreviewWritePlan()

    Dim slide, i, total, barCount
    total = pres.Slides.Count
    barCount = 0
    For i = 1 To total
        Set slide = pres.Slides(i)
        RemoveExistingProgress slide
        AddProgress slide, i, total, pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight
        barCount = barCount + 1
    Next
    SafeCloseWritePlan planId

    Dim summary
    summary = "PPT 页码进度条添加完成；页数=" & CStr(total) & "；写入页=" & CStr(barCount) & vbCrLf & _
        "Host.GetPptContextInfo: " & pptContext & vbCrLf & _
        "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostPptAddPageNumberProgress = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""slideCount"":" & CStr(total) & _
        ",""barCount"":" & CStr(barCount) & ",""planId"":""" & EscapeJson(planId) & """}"

End Function

Sub RemoveExistingProgress(slide)
    On Error Resume Next
    Dim i, shapeName
    For i = slide.Shapes.Count To 1 Step -1
        shapeName = slide.Shapes(i).Name
        If Left(shapeName, 16) = "HostProgressBar_" Or Left(shapeName, 19) = "HostProgressLabel_" Then
            slide.Shapes(i).Delete
        End If
        If Err.Number <> 0 Then Err.Clear
    Next
End Sub

Sub AddProgress(slide, index, total, slideWidth, slideHeight)
    On Error Resume Next
    Dim barHeight, fullWidth, doneWidth, leftEdge, topEdge, bg, fg, label
    barHeight = 6
    fullWidth = slideWidth
    doneWidth = slideWidth * index / total
    leftEdge = 0
    topEdge = slideHeight - barHeight

    Set bg = slide.Shapes.AddShape(1, leftEdge, topEdge, fullWidth, barHeight)
    bg.Name = "HostProgressBar_bg"
    bg.Fill.ForeColor.RGB = RGB(220, 225, 232)
    bg.Line.Visible = msoFalse

    Set fg = slide.Shapes.AddShape(1, leftEdge, topEdge, doneWidth, barHeight)
    fg.Name = "HostProgressBar_fg"
    fg.Fill.ForeColor.RGB = RGB(47, 117, 181)
    fg.Line.Visible = msoFalse

    Set label = slide.Shapes.AddTextbox(1, slideWidth - 90, slideHeight - 32, 72, 18)
    label.Name = "HostProgressLabel_page"
    label.TextFrame.TextRange.Text = CStr(index) & " / " & CStr(total)
    label.TextFrame.TextRange.Font.Size = 9
    label.TextFrame.TextRange.Font.Name = "Arial"
    label.TextFrame.TextRange.Font.Color.RGB = RGB(90, 90, 90)
    Err.Clear
End Sub

Function SafeHostText(methodName)
    On Error Resume Next
    If methodName = "GetPptContextInfo" Then
        SafeHostText = Host.GetPptContextInfo()
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
