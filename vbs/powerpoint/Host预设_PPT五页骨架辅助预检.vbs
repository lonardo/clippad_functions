' 函数名: HostPptFivePageSkeletonAssist
' 描述: 预览并可选生成 5 页汇报骨架（封面/目录/现状/方案/结论），仅插入标题页结构，不覆盖已有业务内容页；确认词“生成”
' 适用应用: PowerPoint
' 搜索范围: 全文
' 搜索对象: 无

Option Explicit

Const ppLayoutTitle = 1
Const ppLayoutText = 2

Function Main()
    On Error Resume Next
    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = FailureJson("E_NO_PPT_APP", "未取得 PowerPoint 应用，请在 PowerPoint 中运行该预设")
        Exit Function
    End If
    Main = HostPptFivePageSkeletonAssist(appObj)
End Function

Function HostPptFivePageSkeletonAssist(appObj)
    On Error Resume Next

    Dim pres
    Set pres = appObj.ActivePresentation
    If Err.Number <> 0 Or TypeName(pres) = "Empty" Or TypeName(pres) = "Nothing" Then
        Err.Clear
        HostPptFivePageSkeletonAssist = FailureJson("E_NO_PRESENTATION", "当前没有活动演示文稿")
        Exit Function
    End If

    Dim topic, titles(5), idx
    topic = Trim(SafePrompt("汇报主题（用于封面副标题）", "工作汇报"))
    If Len(topic) = 0 Then topic = "工作汇报"
    titles(1) = Trim(SafePrompt("第1页标题", "封面：" & topic))
    titles(2) = Trim(SafePrompt("第2页标题", "目录"))
    titles(3) = Trim(SafePrompt("第3页标题", "现状与问题"))
    titles(4) = Trim(SafePrompt("第4页标题", "方案与进展"))
    titles(5) = Trim(SafePrompt("第5页标题", "结论与下一步"))

    Dim modeText, appendMode
    modeText = Trim(SafePrompt("插入方式：追加到末尾 / 仅预览清单", "追加到末尾"))
    appendMode = (InStr(1, modeText, "预览", vbTextCompare) = 0)

    Dim slideSummary, layoutSummary, previewText, listText
    listText = ""
    For idx = 1 To 5
        listText = listText & CStr(idx) & ". " & titles(idx) & vbCrLf
    Next
    slideSummary = SafeHostText("GetPptSlideSummary")
    layoutSummary = SafeHostText("GetPptLayoutIssueSummary")
    previewText = "PPT 5页骨架辅助预检（尚未写入）" & vbCrLf & _
        "当前页数=" & CStr(pres.Slides.Count) & "；主题=" & topic & vbCrLf & _
        "将新增骨架页：" & vbCrLf & listText & _
        "不覆盖已有页面内容；仅追加结构页" & vbCrLf & _
        "Host.GetPptSlideSummary: " & slideSummary & vbCrLf & _
        "Host.GetPptLayoutIssueSummary: " & layoutSummary

    If Not appendMode Then
        Host.WriteClipboard previewText
        SafeWriteLog previewText
        HostPptFivePageSkeletonAssist = "{""ok"":true,""changed"":false,""message"":""" & EscapeJson(previewText) & """}"
        Exit Function
    End If

    Dim planId, planPreview
    planId = SafeBeginWritePlan()
    SafeRecordWrite "ppt_five_page_skeleton", "{""topic"":""" & EscapeJson(topic) & """,""pages"":5}", "office.ppt.skeleton"
    planPreview = SafePreviewWritePlan()

    Dim confirmText
    confirmText = Trim(SafePrompt(previewText & vbCrLf & vbCrLf & "确认追加 5 页汇报骨架，请输入：生成", ""))
    If StrComp(confirmText, "生成", vbTextCompare) <> 0 Then
        SafeRollbackWritePlan planId
        HostPptFivePageSkeletonAssist = FailureJson("E_CONFIRM_REQUIRED", "未输入“生成”，已取消且未修改演示文稿")
        Exit Function
    End If

    Dim created, slide, custom, pos
    created = 0
    custom = ""
    On Error Resume Next
    If pres.SlideMaster.CustomLayouts.Count >= 1 Then
        ' prefer title layout if available index 1
    End If
    For idx = 1 To 5
        pos = pres.Slides.Count + 1
        Set slide = Nothing
        On Error Resume Next
        Set slide = pres.Slides.Add(pos, ppLayoutTitle)
        If Err.Number <> 0 Or TypeName(slide) = "Nothing" Then
            Err.Clear
            Set slide = pres.Slides.Add(pos, ppLayoutText)
        End If
        If Not slide Is Nothing Then
            ApplySkeletonTitle slide, titles(idx), topic, idx
            created = created + 1
        End If
        Err.Clear
    Next
    SafeRollbackWritePlan planId

    Dim summary
    summary = "PPT 5页骨架已追加；新增=" & CStr(created) & "；主题=" & topic & vbCrLf & listText & "Preview: " & planPreview
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostPptFivePageSkeletonAssist = "{""ok"":true,""changed"":true,""created"":" & CStr(created) & ",""topic"":""" & EscapeJson(topic) & """,""message"":""" & EscapeJson(summary) & """}"
End Function

Sub ApplySkeletonTitle(slide, titleText, topic, idx)
    On Error Resume Next
    Dim shp
    If slide.Shapes.HasTitle Then
        slide.Shapes.Title.TextFrame.TextRange.Text = titleText
    End If
    ' try subtitle / first body
    For Each shp In slide.Shapes
        If shp.HasTextFrame Then
            If Not slide.Shapes.HasTitle Or shp.Name <> slide.Shapes.Title.Name Then
                If idx = 1 Then
                    shp.TextFrame.TextRange.Text = topic & vbCrLf & "汇报骨架（可编辑）"
                ElseIf idx = 2 Then
                    shp.TextFrame.TextRange.Text = "1. 现状与问题" & vbCrLf & "2. 方案与进展" & vbCrLf & "3. 结论与下一步"
                Else
                    shp.TextFrame.TextRange.Text = "要点 1" & vbCrLf & "要点 2" & vbCrLf & "要点 3"
                End If
                Exit For
            End If
        End If
    Next
    Err.Clear
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
