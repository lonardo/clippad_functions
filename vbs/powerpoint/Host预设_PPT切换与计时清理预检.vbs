' 函数名: HostPptTransitionTimingCleanupPreflight
' 描述: 预检幻灯片切换效果与自动换片计时，确认后清除；默认可保留隐藏页设置，与汇报交付清理包互补
' 适用应用: PowerPoint
' 搜索范围: 全文
' 搜索对象: 无

Option Explicit

Const msoTrue = -1
Const msoFalse = 0
Const ppEffectNone = 0

Function Main()
    On Error Resume Next
    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = FailureJson("E_NO_PPT_APP", "未取得 PowerPoint 应用，请在 PowerPoint 中运行该预设")
        Exit Function
    End If
    Main = HostPptTransitionTimingCleanupPreflight(appObj)
End Function

Function HostPptTransitionTimingCleanupPreflight(appObj)
    On Error Resume Next

    Dim pres
    Set pres = appObj.ActivePresentation
    If Err.Number <> 0 Or TypeName(pres) = "Empty" Or TypeName(pres) = "Nothing" Then
        Err.Clear
        HostPptTransitionTimingCleanupPreflight = FailureJson("E_NO_PRESENTATION", "当前没有活动演示文稿")
        Exit Function
    End If

    Dim doTransText, doTimeText, keepHiddenText
    Dim doTrans, doTime, keepHidden
    doTransText = Trim(SafePrompt("清除切换效果？是 / 否", "是"))
    doTimeText = Trim(SafePrompt("清除自动换片计时？是 / 否", "是"))
    keepHiddenText = Trim(SafePrompt("保留隐藏页设置？是 / 否", "是"))
    doTrans = ParseYes(doTransText)
    doTime = ParseYes(doTimeText)
    keepHidden = ParseYes(keepHiddenText)

    If Not (doTrans Or doTime) Then
        HostPptTransitionTimingCleanupPreflight = FailureJson("E_NO_ACTION", "未选择任何清理项，已取消")
        Exit Function
    End If

    Dim slideCount, effectCount, timedCount, hiddenCount, sampleText
    slideCount = pres.Slides.Count
    effectCount = 0
    timedCount = 0
    hiddenCount = 0
    sampleText = ""
    CountTransitionIssues pres, effectCount, timedCount, hiddenCount, sampleText

    Dim slideSummary, layoutSummary, previewText
    slideSummary = SafeHostText("GetPptSlideSummary")
    layoutSummary = SafeHostText("GetPptLayoutIssueSummary")
    previewText = "PPT 切换与计时清理预检（尚未写入）" & vbCrLf & _
        "幻灯片=" & CStr(slideCount) & "；有切换效果=" & CStr(effectCount) & _
        "；有自动计时=" & CStr(timedCount) & "；隐藏页=" & CStr(hiddenCount) & vbCrLf & _
        "将执行：" & BuildPptActionLabel(doTrans, doTime) & _
        "；隐藏页设置=" & KeepHiddenLabel(keepHidden) & vbCrLf & _
        "样例=" & sampleText & vbCrLf & _
        "Host.GetPptSlideSummary: " & slideSummary & vbCrLf & _
        "Host.GetPptLayoutIssueSummary: " & layoutSummary

    Dim planId, planPreview, confirmText
    planId = SafeBeginWritePlan()
    SafeRecordWrite "ppt_transition_timing_cleanup", "{""trans"":" & LCase(CStr(doTrans)) & ",""timing"":" & LCase(CStr(doTime)) & ",""keepHidden"":" & LCase(CStr(keepHidden)) & "}", "office.ppt.transition.cleanup"
    planPreview = SafePreviewWritePlan()

    confirmText = Trim(SafePrompt(previewText & vbCrLf & vbCrLf & "以上尚未写入。确认清理切换与计时，请输入：清理", ""))
    If StrComp(confirmText, "清理", vbTextCompare) <> 0 Then
        SafeRollbackWritePlan planId
        HostPptTransitionTimingCleanupPreflight = FailureJson("E_CONFIRM_REQUIRED", "未输入“清理”，已取消且未修改演示文稿")
        Exit Function
    End If

    Dim clearedEffect, clearedTime
    clearedEffect = 0
    clearedTime = 0
    ApplyTransitionCleanup pres, doTrans, doTime, keepHidden, clearedEffect, clearedTime
    SafeRollbackWritePlan planId

    Dim summary
    summary = "PPT 切换与计时清理完成；清除切换=" & CStr(clearedEffect) & _
        "；清除计时=" & CStr(clearedTime) & "；隐藏页保留=" & KeepHiddenLabel(keepHidden) & vbCrLf & _
        "Preview: " & planPreview
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostPptTransitionTimingCleanupPreflight = "{""ok"":true,""changed"":true,""clearedEffects"":" & CStr(clearedEffect) & _
        ",""clearedTimings"":" & CStr(clearedTime) & ",""message"":""" & EscapeJson(summary) & """}"
End Function

Function ParseYes(text)
    Dim t
    t = UCase(Trim(CStr(text)))
    ParseYes = (t = "是" Or t = "Y" Or t = "YES" Or t = "TRUE" Or t = "1")
End Function

Function BuildPptActionLabel(doTrans, doTime)
    Dim parts, n
    parts = ""
    n = 0
    If doTrans Then
        parts = parts & "切换效果"
        n = n + 1
    End If
    If doTime Then
        If n > 0 Then parts = parts & "+"
        parts = parts & "自动计时"
        n = n + 1
    End If
    If n = 0 Then parts = "无"
    BuildPptActionLabel = parts
End Function

Function KeepHiddenLabel(keepHidden)
    If keepHidden Then
        KeepHiddenLabel = "保留"
    Else
        KeepHiddenLabel = "不保证保留"
    End If
End Function

Sub CountTransitionIssues(pres, ByRef effectCount, ByRef timedCount, ByRef hiddenCount, ByRef sampleText)
    On Error Resume Next
    Dim slide, tr, sampleCount, entry
    effectCount = 0
    timedCount = 0
    hiddenCount = 0
    sampleCount = 0
    sampleText = "（无）"
    For Each slide In pres.Slides
        Set tr = slide.SlideShowTransition
        If tr.Hidden Then hiddenCount = hiddenCount + 1
        If tr.EntryEffect <> ppEffectNone Then
            effectCount = effectCount + 1
            If sampleCount < 6 Then
                entry = "S" & CStr(slide.SlideIndex) & ":fx"
                If sampleCount = 0 Then sampleText = entry Else sampleText = sampleText & "; " & entry
                sampleCount = sampleCount + 1
            End If
        End If
        If tr.AdvanceOnTime Then
            timedCount = timedCount + 1
            If sampleCount < 6 Then
                entry = "S" & CStr(slide.SlideIndex) & ":t" & CStr(tr.AdvanceTime) & "s"
                If sampleText = "（无）" Then
                    sampleText = entry
                Else
                    sampleText = sampleText & "; " & entry
                End If
                sampleCount = sampleCount + 1
            End If
        End If
        Err.Clear
    Next
End Sub

Sub ApplyTransitionCleanup(pres, doTrans, doTime, keepHidden, ByRef clearedEffect, ByRef clearedTime)
    On Error Resume Next
    Dim slide, tr, wasHidden
    clearedEffect = 0
    clearedTime = 0
    For Each slide In pres.Slides
        Set tr = slide.SlideShowTransition
        wasHidden = tr.Hidden
        If doTrans Then
            If tr.EntryEffect <> ppEffectNone Then
                tr.EntryEffect = ppEffectNone
                If Err.Number = 0 Then clearedEffect = clearedEffect + 1
                Err.Clear
            End If
        End If
        If doTime Then
            If tr.AdvanceOnTime Or tr.AdvanceTime <> 0 Then
                tr.AdvanceOnTime = msoFalse
                tr.AdvanceTime = 0
                tr.AdvanceOnClick = msoTrue
                If Err.Number = 0 Then clearedTime = clearedTime + 1
                Err.Clear
            End If
        End If
        If keepHidden Then
            tr.Hidden = wasHidden
        End If
        Err.Clear
    Next
End Sub

Function SafeHostText(methodName)
    On Error Resume Next
    If methodName = "GetPptSlideSummary" Then
        SafeHostText = Host.GetPptSlideSummary()
    ElseIf methodName = "GetPptLayoutIssueSummary" Then
        SafeHostText = Host.GetPptLayoutIssueSummary()
    ElseIf methodName = "GetPptContextInfo" Then
        SafeHostText = Host.GetPptContextInfo()
    Else
        SafeHostText = ""
    End If
    If Err.Number <> 0 Then SafeHostText = ""
    Err.Clear
End Function

Function FailureJson(code, message)
    FailureJson = "{""ok"":false,""code"":""" & EscapeJson(code) & """,""message"":""" & EscapeJson(message) & """}"
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

Function SafePrompt(message, defaultValue)
    On Error Resume Next
    SafePrompt = Host.Prompt(message, defaultValue)
    If Err.Number <> 0 Then SafePrompt = defaultValue
    Err.Clear
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

Sub SafeRollbackWritePlan(planId)
    On Error Resume Next
    Host.RollbackWritePlan planId
    Err.Clear
End Sub

Sub SafeWriteLog(message)
    On Error Resume Next
    Host.WriteLog message
    Err.Clear
End Sub
