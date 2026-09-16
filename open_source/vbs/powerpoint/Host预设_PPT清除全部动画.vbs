' 函数名: HostPptClearAllAnimations
' 描述: 清除当前 PowerPoint 全部幻灯片的进入、强调、退出和路径动画；适合正式汇报前简化演示节奏
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

    Main = HostPptClearAllAnimations(appObj)
End Function

Function HostPptClearAllAnimations(appObj)
    On Error Resume Next

    Dim pres
    Set pres = appObj.ActivePresentation
    If Err.Number <> 0 Or TypeName(pres) = "Empty" Or TypeName(pres) = "Nothing" Then
        Err.Clear
        HostPptClearAllAnimations = "{""ok"":false,""code"":""E_NO_PRESENTATION"",""message"":""当前没有活动演示文稿""}"
        Exit Function
    End If

    Dim confirmation
    confirmation = SafePrompt("将清除全部动画效果。请输入 清除 继续", "")
    If confirmation <> "清除" Then
        HostPptClearAllAnimations = "{""ok"":false,""code"":""E_CONFIRM_REQUIRED"",""message"":""已取消清除动画""}"
        Exit Function
    End If

    Dim pptContext, planId, previewJson, slide, removed, scanned
    pptContext = SafeHostText("GetPptContextInfo")
    planId = SafeBeginWritePlan()
    SafeRecordWrite "ppt_clear_all_animations", "{""scope"":""activePresentation""}", "office.ppt.animation.cleanup"
    previewJson = SafePreviewWritePlan()

    removed = 0
    scanned = 0
    For Each slide In pres.Slides
        scanned = scanned + 1
        removed = removed + ClearSlideAnimations(slide)
    Next
    SafeCloseWritePlan planId

    Dim summary
    summary = "PPT 动画清除完成；扫描页=" & CStr(scanned) & "；删除动画=" & CStr(removed) & vbCrLf & _
        "Host.GetPptContextInfo: " & pptContext & vbCrLf & _
        "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostPptClearAllAnimations = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""scanned"":" & CStr(scanned) & _
        ",""removed"":" & CStr(removed) & ",""planId"":""" & EscapeJson(planId) & """}"
End Function

Function ClearSlideAnimations(slide)
    On Error Resume Next
    Dim timeline, seq, count, i
    count = 0
    Set timeline = slide.TimeLine
    If Err.Number <> 0 Then
        Err.Clear
        ClearSlideAnimations = 0
        Exit Function
    End If

    Set seq = timeline.MainSequence
    If Err.Number = 0 Then
        For i = seq.Count To 1 Step -1
            seq.Item(i).Delete
            If Err.Number = 0 Then count = count + 1
            Err.Clear
        Next
    Else
        Err.Clear
    End If

    Dim interactive, j, interactiveCount
    interactiveCount = 0
    On Error Resume Next
    interactiveCount = timeline.InteractiveSequences.Count
    If Err.Number <> 0 Then
        Err.Clear
        interactiveCount = 0
    End If
    For j = interactiveCount To 1 Step -1
        Set interactive = Nothing
        Set interactive = timeline.InteractiveSequences(j)
        If Err.Number = 0 And Not interactive Is Nothing Then
            For i = interactive.Count To 1 Step -1
                interactive.Item(i).Delete
                If Err.Number = 0 Then count = count + 1
                Err.Clear
            Next
        Else
            Err.Clear
        End If
    Next

    ClearSlideAnimations = count
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