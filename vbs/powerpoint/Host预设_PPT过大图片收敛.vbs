' 函数名: HostPptShrinkLargePictures
' 描述: 将演示文稿中超大图片统一缩小到版心内并锁定纵横比；对标 iSlide 图片整理中的过大图片收敛
' 适用应用: PowerPoint
' 搜索范围: 全文
' 搜索对象: 无

Option Explicit

Const msoPicture = 13
Const msoLinkedPicture = 11
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
    Main = HostPptShrinkLargePictures(appObj)
End Function

Function HostPptShrinkLargePictures(appObj)
    On Error Resume Next
    Dim pres
    Set pres = appObj.ActivePresentation
    If Err.Number <> 0 Or TypeName(pres) = "Empty" Or TypeName(pres) = "Nothing" Then
        Err.Clear
        HostPptShrinkLargePictures = "{""ok"":false,""code"":""E_NO_PRESENTATION"",""message"":""当前没有活动演示文稿""}"
        Exit Function
    End If

    Dim maxRatio, pptSummary, planId, previewJson, slide, shape, changed, scanned
    maxRatio = CDblSafe(SafePrompt("请输入图片最大占版心比例，默认 0.9", "0.9"), 0.9)
    If maxRatio <= 0 Or maxRatio > 1 Then maxRatio = 0.9

    pptSummary = SafeHostText("GetPptSummary")
    planId = SafeBeginWritePlan()
    SafeRecordWrite "ppt_shrink_large_pictures", "{""maxRatio"":" & Replace(CStr(maxRatio), ",", ".") & "}", "office.ppt.picture"
    previewJson = SafePreviewWritePlan()

    Dim maxW, maxH
    maxW = pres.PageSetup.SlideWidth * maxRatio
    maxH = pres.PageSetup.SlideHeight * maxRatio
    changed = 0
    scanned = 0
    For Each slide In pres.Slides
        For Each shape In slide.Shapes
            If shape.Type = msoPicture Or shape.Type = msoLinkedPicture Then
                scanned = scanned + 1
                If shape.Width > maxW Or shape.Height > maxH Then
                    shape.LockAspectRatio = msoTrue
                    If shape.Width / maxW > shape.Height / maxH Then
                        shape.Width = maxW
                    Else
                        shape.Height = maxH
                    End If
                    shape.Left = (pres.PageSetup.SlideWidth - shape.Width) / 2
                    shape.Top = (pres.PageSetup.SlideHeight - shape.Height) / 2
                    changed = changed + 1
                End If
            End If
            Err.Clear
        Next
    Next
    SafeCloseWritePlan planId

    Dim summary
    summary = "PPT 过大图片收敛完成；扫描图片=" & CStr(scanned) & "；调整=" & CStr(changed) & _
        "；最大比例=" & Replace(CStr(maxRatio), ",", ".") & vbCrLf & _
        "Host.GetPptSummary: " & pptSummary & vbCrLf & "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostPptShrinkLargePictures = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""scanned"":" & CStr(scanned) & _
        ",""changed"":" & CStr(changed) & ",""planId"":""" & EscapeJson(planId) & """}"
End Function

Function CDblSafe(value, fallback)
    On Error Resume Next
    CDblSafe = CDbl(value)
    If Err.Number <> 0 Then
        CDblSafe = fallback
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
Function SafePrompt(message, defaultValue)
    On Error Resume Next
    SafePrompt = Host.Prompt(message, defaultValue)
    If Err.Number <> 0 Then
        SafePrompt = defaultValue
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