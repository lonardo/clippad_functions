' 函数名: HostPptRemoveHiddenSlides
' 描述: 删除当前 PowerPoint 中的隐藏幻灯片；适合对外发布前清理内部备份页
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

    Main = HostPptRemoveHiddenSlides(appObj)
End Function

Function HostPptRemoveHiddenSlides(appObj)
    On Error Resume Next

    Dim pres, hiddenCount, index, confirmation, removed, planId
    Set pres = appObj.ActivePresentation
    If Err.Number <> 0 Or TypeName(pres) = "Empty" Or TypeName(pres) = "Nothing" Then
        Err.Clear
        HostPptRemoveHiddenSlides = "{""ok"":false,""code"":""E_NO_PRESENTATION"",""message"":""当前没有活动演示文稿""}"
        Exit Function
    End If

    hiddenCount = CountHiddenSlides(pres)
    If hiddenCount = 0 Then
        HostPptRemoveHiddenSlides = "{""ok"":true,""message"":""当前演示文稿没有隐藏页"",""removedCount"":0}"
        Exit Function
    End If
    If hiddenCount >= pres.Slides.Count Then
        HostPptRemoveHiddenSlides = "{""ok"":false,""code"":""E_ALL_SLIDES_HIDDEN"",""message"":""全部幻灯片均为隐藏页，为避免清空演示文稿未执行删除""}"
        Exit Function
    End If
    confirmation = SafePrompt("将删除 " & CStr(hiddenCount) & " 张隐藏幻灯片。请输入 删除 继续", "")
    If confirmation <> "删除" Then
        HostPptRemoveHiddenSlides = "{""ok"":false,""code"":""E_CONFIRM_REQUIRED"",""message"":""已取消删除隐藏幻灯片""}"
        Exit Function
    End If

    planId = SafeBeginWritePlan()
    SafeRecordWrite "ppt_remove_hidden_slides", "{""hiddenCount"":" & CStr(hiddenCount) & "}", "office.ppt.slide.cleanup"
    removed = 0
    For index = pres.Slides.Count To 1 Step -1
        If pres.Slides(index).SlideShowTransition.Hidden = msoTrue Then
            pres.Slides(index).Delete
            If Err.Number = 0 Then removed = removed + 1
            Err.Clear
        End If
    Next
    SafeCloseWritePlan planId

    Dim summary
    summary = "PPT 隐藏页清理完成；删除=" & CStr(removed) & " 页；剩余=" & CStr(pres.Slides.Count) & " 页"
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostPptRemoveHiddenSlides = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""removedCount"":" & CStr(removed) & "}"
End Function

Function CountHiddenSlides(pres)
    On Error Resume Next
    Dim slide, count
    count = 0
    For Each slide In pres.Slides
        If slide.SlideShowTransition.Hidden = msoTrue Then count = count + 1
        Err.Clear
    Next
    CountHiddenSlides = count
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
    EscapeJson = text
End Function
