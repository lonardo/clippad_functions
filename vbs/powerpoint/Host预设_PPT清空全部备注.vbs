' 函数名: HostPptClearAllNotes
' 描述: 清空当前 PowerPoint 全部幻灯片备注，适合对外分享前去除讲稿和内部提示
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

    Main = HostPptClearAllNotes(appObj)
End Function

Function HostPptClearAllNotes(appObj)
    On Error Resume Next

    Dim pres
    Set pres = appObj.ActivePresentation
    If Err.Number <> 0 Or TypeName(pres) = "Empty" Or TypeName(pres) = "Nothing" Then
        Err.Clear
        HostPptClearAllNotes = "{""ok"":false,""code"":""E_NO_PRESENTATION"",""message"":""当前没有活动演示文稿""}"
        Exit Function
    End If

    Dim confirmation
    confirmation = SafePrompt("将清空全部幻灯片备注。请输入 清空 继续", "")
    If confirmation <> "清空" Then
        HostPptClearAllNotes = "{""ok"":false,""code"":""E_CONFIRM_REQUIRED"",""message"":""已取消清空备注""}"
        Exit Function
    End If

    Dim slideSummary, planId, previewJson, slide, cleared, scanned, notesText
    slideSummary = SafeHostText("GetPptSlideSummary")
    planId = SafeBeginWritePlan()
    SafeRecordWrite "ppt_clear_all_notes", "{""scope"":""activePresentation""}", "office.ppt.notes.cleanup"
    previewJson = SafePreviewWritePlan()

    cleared = 0
    scanned = 0
    For Each slide In pres.Slides
        scanned = scanned + 1
        notesText = ""
        notesText = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
        If Err.Number <> 0 Then
            Err.Clear
            notesText = ""
        End If
        If Len(Trim(Replace(Replace(notesText, vbCr, ""), vbLf, ""))) > 0 Then
            slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = ""
            If Err.Number = 0 Then cleared = cleared + 1
            Err.Clear
        End If
    Next
    SafeCloseWritePlan planId

    Dim summary
    summary = "PPT 备注清空完成；扫描页=" & CStr(scanned) & "；清空页=" & CStr(cleared) & vbCrLf & _
        "Host.GetPptSlideSummary: " & slideSummary & vbCrLf & _
        "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostPptClearAllNotes = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""scanned"":" & CStr(scanned) & _
        ",""cleared"":" & CStr(cleared) & ",""planId"":""" & EscapeJson(planId) & """}"
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
    If methodName = "GetPptSlideSummary" Then
        SafeHostText = Host.GetPptSlideSummary()
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