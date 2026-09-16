' 函数名: HostPptFitPicturesToCanvas
' 描述: 将当前 PowerPoint 演示文稿中的图片按比例适配到统一版心区域并居中；适合批量截图、产品图和案例图排版
' 适用应用: PowerPoint
' 搜索范围: 全文
' 搜索对象: 无

Option Explicit

Const msoPicture = 13
Const msoLinkedPicture = 11

Function Main()
    On Error Resume Next

    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = "{""ok"":false,""code"":""E_NO_PPT_APP"",""message"":""未取得 PowerPoint 应用，请在 PowerPoint 中运行该预设""}"
        Exit Function
    End If

    Main = HostPptFitPicturesToCanvas(appObj)
End Function

Function HostPptFitPicturesToCanvas(appObj)
    On Error Resume Next

    Dim pres
    Set pres = appObj.ActivePresentation
    If Err.Number <> 0 Or TypeName(pres) = "Empty" Or TypeName(pres) = "Nothing" Then
        Err.Clear
        HostPptFitPicturesToCanvas = "{""ok"":false,""code"":""E_NO_PRESENTATION"",""message"":""当前没有活动演示文稿""}"
        Exit Function
    End If

    Dim margin
    margin = CDblSafe(SafePrompt("请输入图片距离页面边缘的留白，默认 48 磅", "48"), 48)
    If margin < 0 Then margin = 48

    Dim planId, previewJson, pptSummary
    pptSummary = SafeHostText("GetPptSummary")
    planId = SafeBeginWritePlan()
    SafeRecordWrite "ppt_fit_pictures_to_canvas", "{""scope"":""activePresentation"",""margin"":" & NumberText(margin) & "}", "office.ppt.picture"
    previewJson = SafePreviewWritePlan()

    Dim slide, shape, count, fitWidth, fitHeight, leftEdge, topEdge
    count = 0
    leftEdge = margin
    topEdge = margin
    fitWidth = pres.PageSetup.SlideWidth - margin * 2
    fitHeight = pres.PageSetup.SlideHeight - margin * 2
    If fitWidth <= 0 Or fitHeight <= 0 Then
        SafeCloseWritePlan planId
        HostPptFitPicturesToCanvas = "{""ok"":false,""code"":""E_BAD_MARGIN"",""message"":""留白过大，无法适配图片""}"
        Exit Function
    End If

    For Each slide In pres.Slides
        For Each shape In slide.Shapes
            If shape.Type = msoPicture Or shape.Type = msoLinkedPicture Then
                FitShape shape, leftEdge, topEdge, fitWidth, fitHeight
                count = count + 1
            End If
        Next
    Next
    SafeCloseWritePlan planId

    Dim summary
    summary = "PPT 图片版心适配完成；图片数=" & CStr(count) & "；留白=" & FormatNumber(margin, 1) & _
        "；版心=" & FormatNumber(fitWidth, 1) & "x" & FormatNumber(fitHeight, 1) & vbCrLf & _
        "Host.GetPptSummary: " & pptSummary & vbCrLf & _
        "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostPptFitPicturesToCanvas = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""pictureCount"":" & CStr(count) & _
        ",""margin"":" & NumberText(margin) & ",""fitWidth"":" & NumberText(fitWidth) & ",""fitHeight"":" & NumberText(fitHeight) & "}"

End Function

Sub FitShape(shape, leftEdge, topEdge, fitWidth, fitHeight)
    On Error Resume Next
    Dim scaleW, scaleH, scaleValue
    If shape.Width <= 0 Or shape.Height <= 0 Then Exit Sub
    shape.LockAspectRatio = -1
    scaleW = fitWidth / shape.Width
    scaleH = fitHeight / shape.Height
    If scaleW < scaleH Then
        scaleValue = scaleW
    Else
        scaleValue = scaleH
    End If
    shape.Width = shape.Width * scaleValue
    shape.Left = leftEdge + (fitWidth - shape.Width) / 2
    shape.Top = topEdge + (fitHeight - shape.Height) / 2
    Err.Clear
End Sub

Function CDblSafe(value, fallback)
    On Error Resume Next
    CDblSafe = CDbl(value)
    If Err.Number <> 0 Then
        CDblSafe = fallback
        Err.Clear
    End If
End Function

Function NumberText(value)
    NumberText = Replace(CStr(value), ",", "")
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
    If methodName = "GetPptSummary" Then
        SafeHostText = Host.GetPptSummary()
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
