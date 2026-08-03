' 函数名: HostPptAlignSelectedShapesLeft
' 描述: 将当前选中对象按最左侧对象左对齐；对标 iSlide 对齐工具中的左对齐
' 适用应用: PowerPoint
' 搜索范围: 选区
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
    Main = HostPptAlignSelectedShapesLeft(appObj)
End Function

Function HostPptAlignSelectedShapesLeft(appObj)
    On Error Resume Next
    Dim sel
    Set sel = appObj.ActiveWindow.Selection
    If Err.Number <> 0 Or TypeName(sel) = "Empty" Or TypeName(sel) = "Nothing" Then
        Err.Clear
        HostPptAlignSelectedShapesLeft = "{""ok"":false,""code"":""E_NO_SELECTION"",""message"":""请在 PowerPoint 中选择至少两个对象""}"
        Exit Function
    End If
    If sel.Type <> 2 Or sel.ShapeRange.Count < 2 Then
        HostPptAlignSelectedShapesLeft = "{""ok"":false,""code"":""E_NEED_SHAPES"",""message"":""请至少选择两个形状或图片对象""}"
        Exit Function
    End If

    Dim selectionSummary, planId, previewJson, i, minLeft
    selectionSummary = SafeHostText("GetPptSelectionSummary")
    planId = SafeBeginWritePlan()
    SafeRecordWrite "ppt_align_selected_shapes_left", "{""scope"":""selection""}", "office.ppt.align"
    previewJson = SafePreviewWritePlan()

    minLeft = sel.ShapeRange(1).Left
    For i = 2 To sel.ShapeRange.Count
        If sel.ShapeRange(i).Left < minLeft Then minLeft = sel.ShapeRange(i).Left
    Next
    For i = 1 To sel.ShapeRange.Count
        sel.ShapeRange(i).Left = minLeft
    Next
    SafeCloseWritePlan planId

    Dim summary
    summary = "PPT 所选对象左对齐完成；对象数=" & CStr(sel.ShapeRange.Count) & "；Left=" & Replace(CStr(minLeft), ",", ".") & vbCrLf & _
        "Host.GetPptSelectionSummary: " & selectionSummary & vbCrLf & "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostPptAlignSelectedShapesLeft = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""shapeCount"":" & CStr(sel.ShapeRange.Count) & _
        ",""left"":" & Replace(CStr(minLeft), ",", ".") & ",""planId"":""" & EscapeJson(planId) & """}"
End Function

Function SafeHostText(methodName)
    On Error Resume Next
    If methodName = "GetPptSelectionSummary" Then
        SafeHostText = Host.GetPptSelectionSummary()
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