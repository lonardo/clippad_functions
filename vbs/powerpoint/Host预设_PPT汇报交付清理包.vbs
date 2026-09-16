' 函数名: HostPptDeliveryCleanupPack
' 描述: 汇报交付前预检并可选清理：空文本框、备注、动画、隐藏页；先给影响摘要，确认词“清理”后按勾选项执行
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
        Main = FailureJson("E_NO_PPT_APP", "未取得 PowerPoint 应用，请在 PowerPoint 中运行该预设")
        Exit Function
    End If
    Main = HostPptDeliveryCleanupPack(appObj)
End Function

Function HostPptDeliveryCleanupPack(appObj)
    On Error Resume Next

    Dim pres
    Set pres = appObj.ActivePresentation
    If Err.Number <> 0 Or TypeName(pres) = "Empty" Or TypeName(pres) = "Nothing" Then
        Err.Clear
        HostPptDeliveryCleanupPack = FailureJson("E_NO_PRESENTATION", "当前没有活动演示文稿")
        Exit Function
    End If

    Dim doEmpty, doNotes, doAnim, doHidden
    doEmpty = ParseYes(SafePrompt("清理空文本框？是 / 否", "是"))
    doNotes = ParseYes(SafePrompt("清空备注？是 / 否", "是"))
    doAnim = ParseYes(SafePrompt("清除全部动画？是 / 否", "是"))
    doHidden = ParseYes(SafePrompt("删除隐藏页？是 / 否", "否"))

    If Not (doEmpty Or doNotes Or doAnim Or doHidden) Then
        HostPptDeliveryCleanupPack = FailureJson("E_NO_ACTION", "未选择任何清理项，已取消")
        Exit Function
    End If

    Dim emptyCount, notesCount, animCount, hiddenCount, slideCount
    emptyCount = 0
    notesCount = 0
    animCount = 0
    hiddenCount = 0
    slideCount = pres.Slides.Count
    CountDeliveryIssues pres, emptyCount, notesCount, animCount, hiddenCount

    Dim slideSummary, layoutSummary, previewText
    slideSummary = SafeHostText("GetPptSlideSummary")
    layoutSummary = SafeHostText("GetPptLayoutIssueSummary")
    previewText = "PPT 汇报交付清理包预检（尚未写入）" & vbCrLf & _
        "幻灯片=" & CStr(slideCount) & vbCrLf & _
        "空文本框=" & CStr(emptyCount) & "；有备注页=" & CStr(notesCount) & _
        "；动画数=" & CStr(animCount) & "；隐藏页=" & CStr(hiddenCount) & vbCrLf & _
        "将执行：" & BuildActionLabel(doEmpty, doNotes, doAnim, doHidden) & vbCrLf & _
        "删除隐藏页不可恢复，请谨慎" & vbCrLf & _
        "Host.GetPptSlideSummary: " & slideSummary & vbCrLf & _
        "Host.GetPptLayoutIssueSummary: " & layoutSummary

    Dim planId, planPreview
    planId = SafeBeginWritePlan()
    SafeRecordWrite "ppt_delivery_cleanup_pack", "{""empty"":" & LCase(CStr(doEmpty)) & ",""notes"":" & LCase(CStr(doNotes)) & ",""anim"":" & LCase(CStr(doAnim)) & ",""hidden"":" & LCase(CStr(doHidden)) & "}", "office.ppt.delivery.cleanup"
    planPreview = SafePreviewWritePlan()

    Dim confirmText
    confirmText = Trim(SafePrompt(previewText & vbCrLf & vbCrLf & "以上尚未写入。确认执行汇报交付清理，请输入：清理", ""))
    If StrComp(confirmText, "清理", vbTextCompare) <> 0 Then
        SafeRollbackWritePlan planId
        HostPptDeliveryCleanupPack = FailureJson("E_CONFIRM_REQUIRED", "未输入“清理”，已取消且未修改演示文稿")
        Exit Function
    End If

    Dim removedEmpty, clearedNotes, removedAnim, deletedHidden
    removedEmpty = 0
    clearedNotes = 0
    removedAnim = 0
    deletedHidden = 0
    ApplyDeliveryCleanup pres, doEmpty, doNotes, doAnim, doHidden, removedEmpty, clearedNotes, removedAnim, deletedHidden
    SafeRollbackWritePlan planId

    Dim summary
    summary = "PPT 汇报交付清理完成；删空文本框=" & CStr(removedEmpty) & _
        "；清空备注页=" & CStr(clearedNotes) & _
        "；删动画=" & CStr(removedAnim) & _
        "；删隐藏页=" & CStr(deletedHidden) & vbCrLf & _
        "Preview: " & planPreview
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostPptDeliveryCleanupPack = "{""ok"":true,""changed"":true,""removedEmpty"":" & CStr(removedEmpty) & ",""clearedNotes"":" & CStr(clearedNotes) & ",""removedAnim"":" & CStr(removedAnim) & ",""deletedHidden"":" & CStr(deletedHidden) & ",""message"":""" & EscapeJson(summary) & """}"
End Function

Function ParseYes(text)
    Dim t
    t = UCase(Trim(CStr(text)))
    ParseYes = (t = "是" Or t = "Y" Or t = "YES" Or t = "TRUE" Or t = "1")
End Function

Function BuildActionLabel(doEmpty, doNotes, doAnim, doHidden)
    Dim parts, n
    parts = ""
    n = 0
    If doEmpty Then
        parts = parts & "空文本框"
        n = n + 1
    End If
    If doNotes Then
        If n > 0 Then parts = parts & "+"
        parts = parts & "备注"
        n = n + 1
    End If
    If doAnim Then
        If n > 0 Then parts = parts & "+"
        parts = parts & "动画"
        n = n + 1
    End If
    If doHidden Then
        If n > 0 Then parts = parts & "+"
        parts = parts & "隐藏页"
        n = n + 1
    End If
    If n = 0 Then parts = "无"
    BuildActionLabel = parts
End Function

Sub CountDeliveryIssues(pres, ByRef emptyCount, ByRef notesCount, ByRef animCount, ByRef hiddenCount)
    On Error Resume Next
    Dim slide, shape, notesText, i, seq, interactive, j
    emptyCount = 0
    notesCount = 0
    animCount = 0
    hiddenCount = 0
    For Each slide In pres.Slides
        If slide.SlideShowTransition.Hidden Then hiddenCount = hiddenCount + 1
        If Err.Number <> 0 Then Err.Clear

        notesText = ""
        notesText = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
        If Err.Number <> 0 Then
            Err.Clear
            notesText = ""
        End If
        If Len(NormalizeText(notesText)) > 0 Then notesCount = notesCount + 1

        For i = 1 To slide.Shapes.Count
            Set shape = slide.Shapes(i)
            If shape.HasTextFrame = msoTrue Then
                If Len(NormalizeText(SafeShapeText(shape))) = 0 Then emptyCount = emptyCount + 1
            End If
        Next

        Set seq = slide.TimeLine.MainSequence
        If Err.Number = 0 Then
            animCount = animCount + seq.Count
        Else
            Err.Clear
        End If
        On Error Resume Next
        For j = 1 To slide.TimeLine.InteractiveSequences.Count
            Set interactive = slide.TimeLine.InteractiveSequences(j)
            If Err.Number = 0 Then animCount = animCount + interactive.Count
            Err.Clear
        Next
    Next
End Sub

Sub ApplyDeliveryCleanup(pres, doEmpty, doNotes, doAnim, doHidden, ByRef removedEmpty, ByRef clearedNotes, ByRef removedAnim, ByRef deletedHidden)
    On Error Resume Next
    Dim slide, i, notesText, shape
    removedEmpty = 0
    clearedNotes = 0
    removedAnim = 0
    deletedHidden = 0

    ' delete hidden slides from end
    If doHidden Then
        For i = pres.Slides.Count To 1 Step -1
            Set slide = pres.Slides(i)
            If slide.SlideShowTransition.Hidden Then
                slide.Delete
                If Err.Number = 0 Then deletedHidden = deletedHidden + 1
                Err.Clear
            End If
        Next
    End If

    For Each slide In pres.Slides
        If doNotes Then
            notesText = ""
            notesText = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
            If Err.Number <> 0 Then
                Err.Clear
                notesText = ""
            End If
            If Len(NormalizeText(notesText)) > 0 Then
                slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = ""
                If Err.Number = 0 Then clearedNotes = clearedNotes + 1
                Err.Clear
            End If
        End If

        If doAnim Then
            removedAnim = removedAnim + ClearSlideAnimations(slide)
        End If

        If doEmpty Then
            For i = slide.Shapes.Count To 1 Step -1
                Set shape = slide.Shapes(i)
                If shape.HasTextFrame = msoTrue Then
                    If Len(NormalizeText(SafeShapeText(shape))) = 0 Then
                        shape.Delete
                        If Err.Number = 0 Then removedEmpty = removedEmpty + 1
                        Err.Clear
                    End If
                End If
            Next
        End If
    Next
End Sub

Function ClearSlideAnimations(slide)
    On Error Resume Next
    Dim timeline, seq, count, i, interactive, j, interactiveCount
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
    interactiveCount = 0
    interactiveCount = timeline.InteractiveSequences.Count
    If Err.Number <> 0 Then
        Err.Clear
        interactiveCount = 0
    End If
    For j = interactiveCount To 1 Step -1
        Set interactive = timeline.InteractiveSequences(j)
        If Err.Number = 0 Then
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

Function SafeShapeText(shape)
    On Error Resume Next
    SafeShapeText = ""
    If shape.TextFrame.HasText = msoTrue Then SafeShapeText = shape.TextFrame.TextRange.Text
    If Err.Number <> 0 Then
        SafeShapeText = ""
        Err.Clear
    End If
End Function

Function NormalizeText(value)
    Dim text
    text = CStr(value)
    text = Replace(text, vbCr, "")
    text = Replace(text, vbLf, "")
    text = Replace(text, vbTab, "")
    text = Replace(text, ChrW(160), " ")
    text = Replace(text, ChrW(12288), " ")
    NormalizeText = Trim(text)
End Function

Function SafeHostText(methodName)
    On Error Resume Next
    If methodName = "GetPptSlideSummary" Then
        SafeHostText = Host.GetPptSlideSummary()
    ElseIf methodName = "GetPptLayoutIssueSummary" Then
        SafeHostText = Host.GetPptLayoutIssueSummary()
    ElseIf methodName = "GetPptContextInfo" Then
        SafeHostText = Host.GetPptContextInfo()
    ElseIf methodName = "GetPptSummary" Then
        SafeHostText = Host.GetPptSummary()
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
