' 函数名: HostPptDistributeSelectedShapesVertical
' 描述: 将当前 PowerPoint 选中的多个对象统一宽度后纵向等距排列；适合流程节点、对比卡片和侧栏模块对齐
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

    Main = HostPptDistributeSelectedShapesVertical(appObj)
End Function

Function HostPptDistributeSelectedShapesVertical(appObj)
    On Error Resume Next

    Dim sel
    Set sel = appObj.ActiveWindow.Selection
    If Err.Number <> 0 Or TypeName(sel) = "Empty" Or TypeName(sel) = "Nothing" Then
        Err.Clear
        HostPptDistributeSelectedShapesVertical = "{""ok"":false,""code"":""E_NO_SELECTION"",""message"":""请在 PowerPoint 中选择至少两个对象""}"
        Exit Function
    End If
    If sel.Type <> 2 Or sel.ShapeRange.Count < 2 Then
        HostPptDistributeSelectedShapesVertical = "{""ok"":false,""code"":""E_NEED_SHAPES"",""message"":""请至少选择两个形状或图片对象""}"
        Exit Function
    End If

    Dim selectionSummary, planId, previewJson
    selectionSummary = SafeHostText("GetPptSelectionSummary")
    planId = SafeBeginWritePlan()
    SafeRecordWrite "ppt_distribute_selected_shapes_vertical", "{""scope"":""selection""}", "office.ppt.align"
    previewJson = SafePreviewWritePlan()

    Dim count, firstShape, baseLeft, baseTop, baseWidth, baseHeight, gap, i
    count = sel.ShapeRange.Count
    Set firstShape = sel.ShapeRange(1)
    baseLeft = firstShape.Left
    baseTop = firstShape.Top
    baseWidth = firstShape.Width
    baseHeight = firstShape.Height
    gap = CDblSafe(SafePrompt("请输入纵向对象间距，默认 12 磅", "12"), 12)

    For i = 1 To count
        sel.ShapeRange(i).Width = baseWidth
        sel.ShapeRange(i).Height = baseHeight
        sel.ShapeRange(i).Left = baseLeft
        sel.ShapeRange(i).Top = baseTop + (i - 1) * (baseHeight + gap)
    Next
    SafeCloseWritePlan planId

    Dim summary
    summary = "PPT 所选对象纵向等距完成；对象数=" & CStr(count) & _
        "；宽=" & FormatNumber(baseWidth, 1) & "；高=" & FormatNumber(baseHeight, 1) & "；间距=" & FormatNumber(gap, 1) & vbCrLf & _
        "Host.GetPptSelectionSummary: " & selectionSummary & vbCrLf & _
        "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostPptDistributeSelectedShapesVertical = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""shapeCount"":" & CStr(count) & _
        ",""width"":" & NumberText(baseWidth) & ",""height"":" & NumberText(baseHeight) & ",""gap"":" & NumberText(gap) & "}"
End Function

Function CDblSafe(value, fallback)
    On Error Resume Next
    CDblSafe = CDbl(value)
    If Err.Number <> 0 Then
        CDblSafe = fallback
        Err.Clear
    End If
End Function

Function NumberText(value)
    NumberText = Replace(CStr(value), ",", "")
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