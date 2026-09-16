' 函数名: HostPptInsertSectionDivider
' 描述: 在当前页前插入一页分区标题页；对标 iSlide/OfficePLUS 的章节页一键生成
' 适用应用: PowerPoint
' 搜索范围: 全文
' 搜索对象: 无

Option Explicit

Const ppLayoutTitle = 1
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
    Main = HostPptInsertSectionDivider(appObj)
End Function

Function HostPptInsertSectionDivider(appObj)
    On Error Resume Next
    Dim pres
    Set pres = appObj.ActivePresentation
    If Err.Number <> 0 Or TypeName(pres) = "Empty" Or TypeName(pres) = "Nothing" Then
        Err.Clear
        HostPptInsertSectionDivider = "{""ok"":false,""code"":""E_NO_PRESENTATION"",""message"":""当前没有活动演示文稿""}"
        Exit Function
    End If

    Dim titleText, insertIndex, slideSummary, planId, previewJson, newSlide
    titleText = Trim(SafePrompt("请输入分区标题，例如 第二章 方案设计", "章节标题"))
    If Len(titleText) = 0 Then titleText = "章节标题"
    insertIndex = 1
    On Error Resume Next
    insertIndex = appObj.ActiveWindow.View.Slide.SlideIndex
    If Err.Number <> 0 Or insertIndex < 1 Then insertIndex = 1
    Err.Clear

    slideSummary = SafeHostText("GetPptSlideSummary")
    planId = SafeBeginWritePlan()
    SafeRecordWrite "ppt_insert_section_divider", "{""index"":" & CStr(insertIndex) & "}", "office.ppt.structure"
    previewJson = SafePreviewWritePlan()

    Set newSlide = pres.Slides.Add(insertIndex, ppLayoutTitle)
    If newSlide.Shapes.HasTitle Then
        newSlide.Shapes.Title.TextFrame.TextRange.Text = titleText
        newSlide.Shapes.Title.TextFrame.TextRange.Font.NameFarEast = "微软雅黑"
        newSlide.Shapes.Title.TextFrame.TextRange.Font.Name = "Arial"
        newSlide.Shapes.Title.TextFrame.TextRange.Font.Size = 36
        newSlide.Shapes.Title.TextFrame.TextRange.Font.Bold = msoTrue
    End If
    SafeCloseWritePlan planId

    Dim summary
    summary = "PPT 分区标题页已插入；标题=" & titleText & "；位置=第 " & CStr(insertIndex) & " 页" & vbCrLf & _
        "Host.GetPptSlideSummary: " & slideSummary & vbCrLf & "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostPptInsertSectionDivider = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""title"":""" & EscapeJson(titleText) & _
        """,""index"":" & CStr(insertIndex) & ",""planId"":""" & EscapeJson(planId) & """}"
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