' 函数名: HostPptBatchReplaceText
' 描述: 在当前 PowerPoint 演示文稿所有文本框中批量替换文本；验证 Host.Prompt、GetPptSlideSummary、写入计划和剪贴板摘要
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

    Main = HostPptBatchReplaceText(appObj)
End Function

Function HostPptBatchReplaceText(appObj)
    On Error Resume Next

    Dim pres
    Set pres = appObj.ActivePresentation
    If Err.Number <> 0 Or TypeName(pres) = "Empty" Or TypeName(pres) = "Nothing" Then
        Err.Clear
        HostPptBatchReplaceText = "{""ok"":false,""code"":""E_NO_PRESENTATION"",""message"":""当前没有活动演示文稿""}"
        Exit Function
    End If

    Dim findText, replaceText
    findText = SafePrompt("请输入要替换的文本", "")
    If Len(findText) = 0 Then
        HostPptBatchReplaceText = "{""ok"":false,""code"":""E_EMPTY_FIND"",""message"":""查找文本为空，已取消""}"
        Exit Function
    End If
    replaceText = SafePrompt("请输入替换后的文本", "")

    Dim slideSummary, planId, previewJson
    slideSummary = SafeHostText("GetPptSlideSummary")
    planId = SafeBeginWritePlan()
    SafeRecordWrite "ppt_batch_replace_text", "{""scope"":""activePresentation"",""find"":""" & EscapeJson(findText) & """,""replace"":""" & EscapeJson(replaceText) & """}", "office.ppt.textReplace"
    previewJson = SafePreviewWritePlan()

    Dim replacedCount
    replacedCount = ReplacePresentationText(pres, findText, replaceText)
    SafeCloseWritePlan planId

    Dim summary
    summary = "PPT 文本批量替换完成；查找=" & findText & "；替换为=" & replaceText & "；替换次数=" & CStr(replacedCount) & vbCrLf & _
        "Host.GetPptSlideSummary: " & slideSummary & vbCrLf & _
        "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostPptBatchReplaceText = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""replacedCount"":" & CStr(replacedCount) & _
        ",""planId"":""" & EscapeJson(planId) & """}"

End Function

Function ReplacePresentationText(pres, findText, replaceText)
    On Error Resume Next
    Dim slide, shape, beforeText, afterText, count
    count = 0
    For Each slide In pres.Slides
        For Each shape In slide.Shapes
            If shape.HasTextFrame = msoTrue Then
                If shape.TextFrame.HasText = msoTrue Then
                    beforeText = CStr(shape.TextFrame.TextRange.Text)
                    afterText = Replace(beforeText, findText, replaceText, 1, -1, vbTextCompare)
                    If afterText <> beforeText Then
                        count = count + CountOccurrences(beforeText, findText)
                        shape.TextFrame.TextRange.Text = afterText
                    End If
                End If
            End If
        Next
    Next
    ReplacePresentationText = count
End Function

Function CountOccurrences(text, needle)
    Dim count, startAt, pos
    count = 0
    startAt = 1
    If Len(needle) = 0 Then
        CountOccurrences = 0
        Exit Function
    End If
    Do
        pos = InStr(startAt, text, needle, vbTextCompare)
        If pos <= 0 Then Exit Do
        count = count + 1
        startAt = pos + Len(needle)
    Loop
    CountOccurrences = count
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
