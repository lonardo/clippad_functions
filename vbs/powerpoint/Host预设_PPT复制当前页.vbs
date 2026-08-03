' 函数名: HostPptDuplicateSelectedSlide
' 描述: 复制当前幻灯片并紧随其后插入；对标 PPT 效率插件中的当前页快速复制
' 适用应用: PowerPoint
' 搜索范围: 无
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
    Main = HostPptDuplicateSelectedSlide(appObj)
End Function

Function HostPptDuplicateSelectedSlide(appObj)
    On Error Resume Next
    Dim pres, slide
    Set pres = appObj.ActivePresentation
    Set slide = appObj.ActiveWindow.View.Slide
    If Err.Number <> 0 Or TypeName(pres) = "Empty" Or TypeName(slide) = "Empty" Then
        Err.Clear
        HostPptDuplicateSelectedSlide = "{""ok"":false,""code"":""E_NO_SLIDE"",""message"":""请先选中一张幻灯片""}"
        Exit Function
    End If

    Dim pptContext, planId, previewJson, beforeCount, afterCount, newIndex
    pptContext = SafeHostText("GetPptContextInfo")
    planId = SafeBeginWritePlan()
    SafeRecordWrite "ppt_duplicate_selected_slide", "{""index"":" & CStr(slide.SlideIndex) & "}", "office.ppt.structure"
    previewJson = SafePreviewWritePlan()

    beforeCount = pres.Slides.Count
    slide.Duplicate
    afterCount = pres.Slides.Count
    newIndex = slide.SlideIndex + 1
    SafeCloseWritePlan planId

    Dim summary
    summary = "PPT 当前页复制完成；原页=" & CStr(slide.SlideIndex) & "；新页约=" & CStr(newIndex) & _
        "；总页数 " & CStr(beforeCount) & " -> " & CStr(afterCount) & vbCrLf & _
        "Host.GetPptContextInfo: " & pptContext & vbCrLf & "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostPptDuplicateSelectedSlide = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""beforeCount"":" & CStr(beforeCount) & _
        ",""afterCount"":" & CStr(afterCount) & ",""planId"":""" & EscapeJson(planId) & """}"
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
