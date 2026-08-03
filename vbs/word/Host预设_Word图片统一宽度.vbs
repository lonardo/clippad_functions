' 函数名: HostWordNormalizeInlinePictureWidth
' 描述: 将当前 Word 文档中的图片统一缩放到指定厘米宽度并保持纵横比；适合报告插图、截图和扫描图快速规范
' 适用应用: Word
' 搜索范围: 全文
' 搜索对象: 无

Option Explicit

Function Main()
    On Error Resume Next

    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = "{""ok"":false,""code"":""E_NO_WORD_APP"",""message"":""未取得 Word 应用，请在 Word 中运行该预设""}"
        Exit Function
    End If

    Main = HostWordNormalizeInlinePictureWidth(appObj)
End Function

Function HostWordNormalizeInlinePictureWidth(appObj)
    On Error Resume Next

    Dim doc
    Set doc = appObj.ActiveDocument
    If Err.Number <> 0 Or TypeName(doc) = "Empty" Or TypeName(doc) = "Nothing" Then
        Err.Clear
        HostWordNormalizeInlinePictureWidth = "{""ok"":false,""code"":""E_NO_DOCUMENT"",""message"":""当前没有活动 Word 文档""}"
        Exit Function
    End If

    Dim widthCmText, widthCm, targetPoints
    widthCmText = SafePrompt("请输入统一图片宽度（厘米），例如 14", "14")
    If Not IsNumeric(widthCmText) Then
        HostWordNormalizeInlinePictureWidth = "{""ok"":false,""code"":""E_INVALID_WIDTH"",""message"":""宽度必须是数字""}"
        Exit Function
    End If
    widthCm = CDbl(widthCmText)
    If widthCm <= 0 Or widthCm > 40 Then
        HostWordNormalizeInlinePictureWidth = "{""ok"":false,""code"":""E_INVALID_WIDTH"",""message"":""宽度应在 0 到 40 厘米之间""}"
        Exit Function
    End If
    targetPoints = appObj.CentimetersToPoints(widthCm)

    Dim wordSummary, planId, previewJson, changed, total, shape, inlineShape
    wordSummary = SafeHostText("GetWordContextInfo")
    planId = SafeBeginWritePlan()
    SafeRecordWrite "word_normalize_picture_width", "{""widthCm"":" & Replace(CStr(widthCm), ",", ".") & "}", "office.word.image.normalize"
    previewJson = SafePreviewWritePlan()

    total = 0
    changed = 0
    For Each inlineShape In doc.InlineShapes
        total = total + 1
        inlineShape.LockAspectRatio = -1
        inlineShape.Width = targetPoints
        If Err.Number = 0 Then changed = changed + 1
        Err.Clear
    Next
    For Each shape In doc.Shapes
        total = total + 1
        shape.LockAspectRatio = -1
        shape.Width = targetPoints
        If Err.Number = 0 Then changed = changed + 1
        Err.Clear
    Next
    SafeCloseWritePlan planId

    Dim summary
    summary = "Word 图片统一宽度完成；目标宽度=" & CStr(widthCm) & " 厘米；处理对象=" & CStr(changed) & "；扫描对象=" & CStr(total) & vbCrLf & _
        "Host.GetWordContextInfo: " & wordSummary & vbCrLf & _
        "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostWordNormalizeInlinePictureWidth = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""changed"":" & CStr(changed) & _
        ",""scanned"":" & CStr(total) & ",""widthCm"":" & Replace(CStr(widthCm), ",", ".") & ",""planId"":""" & EscapeJson(planId) & """}"
End Function

Function SafeHostText(methodName)
    On Error Resume Next
    If methodName = "GetWordContextInfo" Then
        SafeHostText = Host.GetWordContextInfo()
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

Function SafeValidateWritePlan(planId)
    On Error Resume Next
    SafeValidateWritePlan = Host.ValidateWritePlan(planId, False)
    If Err.Number <> 0 Then SafeValidateWritePlan = ""
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
