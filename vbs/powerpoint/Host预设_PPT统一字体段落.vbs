' 函数名: HostPptNormalizeFontsParagraphs
' 描述: 批量统一当前 PowerPoint 演示文稿文本框的中英文字体、字号和段落间距；适合模板混用后的版式统一
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

    Main = HostPptNormalizeFontsParagraphs(appObj)
End Function

Function HostPptNormalizeFontsParagraphs(appObj)
    On Error Resume Next

    Dim pres
    Set pres = appObj.ActivePresentation
    If Err.Number <> 0 Or TypeName(pres) = "Empty" Or TypeName(pres) = "Nothing" Then
        Err.Clear
        HostPptNormalizeFontsParagraphs = "{""ok"":false,""code"":""E_NO_PRESENTATION"",""message"":""当前没有活动演示文稿""}"
        Exit Function
    End If

    Dim cnFont, enFont, titleSize, bodySize
    cnFont = SafePrompt("请输入中文字体", "微软雅黑")
    enFont = SafePrompt("请输入英文字体", "Arial")
    titleSize = CDblSafe(SafePrompt("请输入标题字号", "28"), 28)
    bodySize = CDblSafe(SafePrompt("请输入正文字号", "18"), 18)

    Dim pptSummary, planId, previewJson
    pptSummary = SafeHostText("GetPptSummary")
    planId = SafeBeginWritePlan()
    SafeRecordWrite "ppt_normalize_fonts_paragraphs", "{""scope"":""activePresentation""}", "office.ppt.typography"
    previewJson = SafePreviewWritePlan()

    Dim slide, shape, changed, titleCount, bodyCount
    changed = 0
    titleCount = 0
    bodyCount = 0
    For Each slide In pres.Slides
        For Each shape In slide.Shapes
            If shape.HasTextFrame = msoTrue Then
                If shape.TextFrame.HasText = msoTrue Then
                    If IsTitleShape(slide, shape) Then
                        ApplyTextStyle shape, cnFont, enFont, titleSize, True
                        titleCount = titleCount + 1
                    Else
                        ApplyTextStyle shape, cnFont, enFont, bodySize, False
                        bodyCount = bodyCount + 1
                    End If
                    changed = changed + 1
                End If
            End If
            If Err.Number <> 0 Then Err.Clear
        Next
    Next
    SafeCloseWritePlan planId

    Dim summary
    summary = "PPT 字体段落统一完成；文本框=" & CStr(changed) & "；标题=" & CStr(titleCount) & _
        "；正文=" & CStr(bodyCount) & "；中文字体=" & cnFont & "；英文字体=" & enFont & vbCrLf & _
        "Host.GetPptSummary: " & pptSummary & vbCrLf & _
        "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostPptNormalizeFontsParagraphs = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""textShapeCount"":" & CStr(changed) & _
        ",""titleCount"":" & CStr(titleCount) & ",""bodyCount"":" & CStr(bodyCount) & "}"

End Function

Function IsTitleShape(slide, shape)
    On Error Resume Next
    Dim titleShape
    IsTitleShape = False
    Set titleShape = slide.Shapes.Title
    If Err.Number = 0 Then
        If titleShape.Name = shape.Name Then IsTitleShape = True
    End If
    Err.Clear
End Function

Sub ApplyTextStyle(shape, cnFont, enFont, fontSize, isTitle)
    On Error Resume Next
    shape.TextFrame.TextRange.Font.NameFarEast = cnFont
    shape.TextFrame.TextRange.Font.Name = enFont
    shape.TextFrame.TextRange.Font.Size = fontSize
    shape.TextFrame.TextRange.Font.Bold = BoolToMso(isTitle)
    shape.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 6
    shape.TextFrame.TextRange.ParagraphFormat.SpaceBefore = 0
    Err.Clear
End Sub

Function BoolToMso(value)
    If CBool(value) Then
        BoolToMso = -1
    Else
        BoolToMso = 0
    End If
End Function

Function CDblSafe(value, fallback)
    On Error Resume Next
    CDblSafe = CDbl(value)
    If Err.Number <> 0 Then
        CDblSafe = fallback
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
