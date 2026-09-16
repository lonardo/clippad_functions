' 函数名: HostPptRemoveEmptyTextBoxes
' 描述: 删除当前 PowerPoint 演示文稿中的空文本框和仅含空白字符的文本框，并复制清理摘要；适合模板复用后的版面清理
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

    Main = HostPptRemoveEmptyTextBoxes(appObj)
End Function

Function HostPptRemoveEmptyTextBoxes(appObj)
    On Error Resume Next

    Dim pres
    Set pres = appObj.ActivePresentation
    If Err.Number <> 0 Or TypeName(pres) = "Empty" Or TypeName(pres) = "Nothing" Then
        Err.Clear
        HostPptRemoveEmptyTextBoxes = "{""ok"":false,""code"":""E_NO_PRESENTATION"",""message"":""当前没有活动演示文稿""}"
        Exit Function
    End If

    Dim slideSummary, planId, previewJson
    slideSummary = SafeHostText("GetPptSlideSummary")
    planId = SafeBeginWritePlan()
    SafeRecordWrite "ppt_remove_empty_textboxes", "{""scope"":""activePresentation""}", "office.ppt.clean"
    previewJson = SafePreviewWritePlan()

    Dim slide, removed, scanned
    removed = 0
    scanned = 0
    For Each slide In pres.Slides
        removed = removed + RemoveEmptyTextBoxesOnSlide(slide, scanned)
    Next
    SafeCloseWritePlan planId

    Dim summary
    summary = "PPT 空文本框清理完成；扫描文本框=" & CStr(scanned) & "；删除=" & CStr(removed) & _
        "；幻灯片数=" & CStr(pres.Slides.Count) & vbCrLf & _
        "Host.GetPptSlideSummary: " & slideSummary & vbCrLf & _
        "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostPptRemoveEmptyTextBoxes = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""scanned"":" & CStr(scanned) & _
        ",""removed"":" & CStr(removed) & ",""slideCount"":" & CStr(pres.Slides.Count) & "}"

End Function

Function RemoveEmptyTextBoxesOnSlide(slide, ByRef scanned)
    On Error Resume Next
    Dim i, shape, removed, textValue
    removed = 0
    For i = slide.Shapes.Count To 1 Step -1
        Set shape = slide.Shapes(i)
        If shape.HasTextFrame = msoTrue Then
            scanned = scanned + 1
            textValue = ""
            If shape.TextFrame.HasText = msoTrue Then textValue = shape.TextFrame.TextRange.Text
            If Len(NormalizeText(textValue)) = 0 Then
                shape.Delete
                removed = removed + 1
            End If
        End If
        If Err.Number <> 0 Then Err.Clear
    Next
    RemoveEmptyTextBoxesOnSlide = removed
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
