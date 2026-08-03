' 函数名: HostPptRemoveHyperlinks
' 描述: 删除当前 PowerPoint 中全部超链接并保留显示文字；适合对外发布前去掉内部跳转和外部链接
' 适用应用: PowerPoint
' 搜索范围: 全文
' 搜索对象: 无

Option Explicit

Const msoTrue = -1
Const msoGroup = 6
Const ppMouseClick = 1
Const ppMouseOver = 2
Const ppActionNone = 0

Function Main()
    On Error Resume Next

    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = "{""ok"":false,""code"":""E_NO_PPT_APP"",""message"":""未取得 PowerPoint 应用，请在 PowerPoint 中运行该预设""}"
        Exit Function
    End If

    Main = HostPptRemoveHyperlinks(appObj)
End Function

Function HostPptRemoveHyperlinks(appObj)
    On Error Resume Next

    Dim pres
    Set pres = appObj.ActivePresentation
    If Err.Number <> 0 Or TypeName(pres) = "Empty" Or TypeName(pres) = "Nothing" Then
        Err.Clear
        HostPptRemoveHyperlinks = "{""ok"":false,""code"":""E_NO_PRESENTATION"",""message"":""当前没有活动演示文稿""}"
        Exit Function
    End If

    Dim confirmation
    confirmation = SafePrompt("将删除全部超链接并保留文字。请输入 删除 继续", "")
    If confirmation <> "删除" Then
        HostPptRemoveHyperlinks = "{""ok"":false,""code"":""E_CONFIRM_REQUIRED"",""message"":""已取消删除超链接""}"
        Exit Function
    End If

    Dim pptSummary, planId, previewJson, slide, shape, removed, scanned
    pptSummary = SafeHostText("GetPptSummary")
    planId = SafeBeginWritePlan()
    SafeRecordWrite "ppt_remove_hyperlinks", "{""scope"":""activePresentation""}", "office.ppt.link.cleanup"
    previewJson = SafePreviewWritePlan()

    removed = 0
    scanned = 0
    For Each slide In pres.Slides
        For Each shape In slide.Shapes
            removed = removed + RemoveShapeHyperlinks(shape, scanned)
        Next
    Next
    SafeCloseWritePlan planId

    Dim summary
    summary = "PPT 超链接清理完成；扫描对象=" & CStr(scanned) & "；清理动作=" & CStr(removed) & vbCrLf & _
        "Host.GetPptSummary: " & pptSummary & vbCrLf & _
        "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostPptRemoveHyperlinks = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""scanned"":" & CStr(scanned) & _
        ",""removed"":" & CStr(removed) & ",""planId"":""" & EscapeJson(planId) & """}"
End Function

Function RemoveShapeHyperlinks(shape, ByRef scanned)
    On Error Resume Next
    Dim removed, i, action, j
    removed = 0
    scanned = scanned + 1

    If shape.Type = msoGroup Then
        For i = 1 To shape.GroupItems.Count
            removed = removed + RemoveShapeHyperlinks(shape.GroupItems(i), scanned)
        Next
        RemoveShapeHyperlinks = removed
        Exit Function
    End If

    Set action = shape.ActionSettings(ppMouseClick)
    If Err.Number = 0 Then
        If action.Action <> ppActionNone Then
            action.Action = ppActionNone
            If Err.Number = 0 Then removed = removed + 1
            Err.Clear
        End If
    Else
        Err.Clear
    End If

    Set action = shape.ActionSettings(ppMouseOver)
    If Err.Number = 0 Then
        If action.Action <> ppActionNone Then
            action.Action = ppActionNone
            If Err.Number = 0 Then removed = removed + 1
            Err.Clear
        End If
    Else
        Err.Clear
    End If

    If shape.HasTextFrame = msoTrue Then
        If shape.TextFrame.HasText = msoTrue Then
            For j = shape.TextFrame.TextRange.Hyperlinks.Count To 1 Step -1
                shape.TextFrame.TextRange.Hyperlinks(j).Delete
                If Err.Number = 0 Then removed = removed + 1
                Err.Clear
            Next
        End If
    End If

    RemoveShapeHyperlinks = removed
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