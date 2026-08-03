' 函数名: HostPptCountTextStats
' 描述: 统计当前演示文稿页数、文本框数、字符数和备注页数并复制摘要；对标 PPT 体检插件的内容统计
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
    Main = HostPptCountTextStats(appObj)
End Function

Function HostPptCountTextStats(appObj)
    On Error Resume Next
    Dim pres
    Set pres = appObj.ActivePresentation
    If Err.Number <> 0 Or TypeName(pres) = "Empty" Or TypeName(pres) = "Nothing" Then
        Err.Clear
        HostPptCountTextStats = "{""ok"":false,""code"":""E_NO_PRESENTATION"",""message"":""当前没有活动演示文稿""}"
        Exit Function
    End If

    Dim slideSummary, slide, shape, textBoxes, chars, notesPages, notesText
    slideSummary = SafeHostText("GetPptSlideSummary")
    textBoxes = 0
    chars = 0
    notesPages = 0
    For Each slide In pres.Slides
        For Each shape In slide.Shapes
            If shape.HasTextFrame = msoTrue Then
                If shape.TextFrame.HasText = msoTrue Then
                    textBoxes = textBoxes + 1
                    chars = chars + Len(shape.TextFrame.TextRange.Text)
                End If
            End If
        Next
        notesText = ""
        notesText = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
        If Err.Number = 0 Then
            If Len(Trim(Replace(Replace(notesText, vbCr, ""), vbLf, ""))) > 0 Then notesPages = notesPages + 1
        Else
            Err.Clear
        End If
    Next

    Dim summary
    summary = "PPT 内容统计完成；页数=" & CStr(pres.Slides.Count) & "；文本框=" & CStr(textBoxes) & _
        "；字符数=" & CStr(chars) & "；有备注页=" & CStr(notesPages) & vbCrLf & _
        "Host.GetPptSlideSummary: " & slideSummary
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostPptCountTextStats = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""slides"":" & CStr(pres.Slides.Count) & _
        ",""textBoxes"":" & CStr(textBoxes) & ",""chars"":" & CStr(chars) & ",""notesPages"":" & CStr(notesPages) & "}"
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