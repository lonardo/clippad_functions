' 函数名: HostPptSnapShapesToGrid
' 描述: 将当前 PowerPoint 演示文稿中的可见对象位置和尺寸吸附到指定网格；适合批量整理轻微错位的页面版式
' 适用应用: PowerPoint
' 搜索范围: 全文
' 搜索对象: 无

Option Explicit

Function Main()
    On Error Resume Next

    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = "{""ok"":false,""code"":""E_NO_PPT_APP"",""message"":""未取得 PowerPoint 应用，请在 PowerPoint 中运行该预设""}"
        Exit Function
    End If

    Main = HostPptSnapShapesToGrid(appObj)
End Function

Function HostPptSnapShapesToGrid(appObj)
    On Error Resume Next

    Dim pres
    Set pres = appObj.ActivePresentation
    If Err.Number <> 0 Or TypeName(pres) = "Empty" Or TypeName(pres) = "Nothing" Then
        Err.Clear
        HostPptSnapShapesToGrid = "{""ok"":false,""code"":""E_NO_PRESENTATION"",""message"":""当前没有活动演示文稿""}"
        Exit Function
    End If

    Dim gridSize
    gridSize = CDblSafe(SafePrompt("请输入网格大小，默认 12 磅", "12"), 12)
    If gridSize <= 0 Then gridSize = 12

    Dim pptContext, planId, previewJson
    pptContext = SafeHostText("GetPptContextInfo")
    planId = SafeBeginWritePlan()
    SafeRecordWrite "ppt_snap_shapes_to_grid", "{""scope"":""activePresentation"",""gridSize"":" & NumberText(gridSize) & "}", "office.ppt.layout"
    previewJson = SafePreviewWritePlan()

    Dim slide, shape, adjusted, skipped
    adjusted = 0
    skipped = 0
    For Each slide In pres.Slides
        For Each shape In slide.Shapes
            If ShouldSkipShape(shape) Then
                skipped = skipped + 1
            Else
                SnapShape shape, gridSize
                adjusted = adjusted + 1
            End If
            If Err.Number <> 0 Then Err.Clear
        Next
    Next
    SafeCloseWritePlan planId

    Dim summary
    summary = "PPT 网格对齐整理完成；网格=" & FormatNumber(gridSize, 1) & _
        "；调整对象=" & CStr(adjusted) & "；跳过=" & CStr(skipped) & vbCrLf & _
        "Host.GetPptContextInfo: " & pptContext & vbCrLf & _
        "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostPptSnapShapesToGrid = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""adjusted"":" & CStr(adjusted) & _
        ",""skipped"":" & CStr(skipped) & ",""gridSize"":" & NumberText(gridSize) & "}"

End Function

Function ShouldSkipShape(shape)
    On Error Resume Next
    Dim nameText
    nameText = CStr(shape.Name)
    ShouldSkipShape = (Left(nameText, 16) = "HostProgressBar_" Or Left(nameText, 19) = "HostProgressLabel_")
    If Err.Number <> 0 Then
        ShouldSkipShape = True
        Err.Clear
    End If
End Function

Sub SnapShape(shape, gridSize)
    On Error Resume Next
    shape.Left = SnapValue(shape.Left, gridSize)
    shape.Top = SnapValue(shape.Top, gridSize)
    If shape.Width > gridSize Then shape.Width = SnapValue(shape.Width, gridSize)
    If shape.Height > gridSize Then shape.Height = SnapValue(shape.Height, gridSize)
    Err.Clear
End Sub

Function SnapValue(value, gridSize)
    SnapValue = Int((CDbl(value) / gridSize) + 0.5) * gridSize
End Function

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
