' 函数名: HostWordNormalizeSelectionWhitespace
' 描述: 读取当前选区文本，规范化全角空格、制表符、重复空白和空行后写回选区；验证 Host.ExtractRange、ReplaceText、WriteRange
' 适用应用: Word
' 搜索范围: 选区
' 搜索对象: 文本

Option Explicit

Function Main()
    On Error Resume Next

    Dim appObj
    Set appObj = Nothing
    Err.Clear
    Set appObj = Host.GetApplication()
    Err.Clear

    Main = HostWordNormalizeSelectionWhitespace(appObj)
End Function

Function HostWordNormalizeSelectionWhitespace(appObj)
    On Error Resume Next

    Dim sourceText
    sourceText = Host.ExtractRange("selection")
    If Err.Number <> 0 Then
        Err.Clear
        HostWordNormalizeSelectionWhitespace = "{""ok"":false,""code"":""E_EXTRACT_RANGE"",""message"":""无法读取当前选区，请先选择 Word 文本""}"
        Exit Function
    End If

    If Len(sourceText) = 0 Then
        HostWordNormalizeSelectionWhitespace = "{""ok"":false,""code"":""E_EMPTY_SELECTION"",""message"":""当前选区为空""}"
        Exit Function
    End If

    Dim originalLength, normalized, changed
    originalLength = Len(sourceText)
    normalized = sourceText
    normalized = Host.ReplaceText(normalized, ChrW(12288), " ", False, True)
    normalized = Host.ReplaceText(normalized, vbTab, " ", False, True)
    normalized = CollapseRepeatedSpaces(normalized)
    normalized = CollapseBlankLines(normalized)
    changed = (normalized <> sourceText)

    If Not changed Then
        HostWordNormalizeSelectionWhitespace = "{""ok"":true,""changed"":false,""message"":""选区空白已经较规范，无需写回"",""length"":" & CStr(originalLength) & "}"
        Exit Function
    End If

    Dim planId, previewJson, validationJson
    planId = SafeBeginWritePlan()
    SafeRecordWrite "range_write", "{""target"":""selection"",""textLength"":" & CStr(Len(normalized)) & "}", "office.range.write"
    validationJson = SafeValidateWritePlan(planId)
    previewJson = SafePreviewWritePlan()

    Dim ok
    ok = Host.WriteRange(normalized, "selection")
    SafeCloseWritePlan planId

    Dim summary
    summary = "Word 选区空白规范化完成；原长度=" & CStr(originalLength) & "，新长度=" & CStr(Len(normalized))
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostWordNormalizeSelectionWhitespace = "{""ok"":" & JsonBool(ok) & ",""changed"":true,""message"":""" & EscapeJson(summary) & _
        """,""planId"":""" & EscapeJson(planId) & """,""preview"":""" & EscapeJson(previewJson) & _
        """,""validation"":""" & EscapeJson(validationJson) & """}"

End Function

Function CollapseRepeatedSpaces(text)
    Dim result
    result = CStr(text)
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    CollapseRepeatedSpaces = result
End Function

Function CollapseBlankLines(text)
    Dim result
    result = CStr(text)
    Do While InStr(result, vbCrLf & vbCrLf & vbCrLf) > 0
        result = Replace(result, vbCrLf & vbCrLf & vbCrLf, vbCrLf & vbCrLf)
    Loop
    Do While InStr(result, vbLf & vbLf & vbLf) > 0
        result = Replace(result, vbLf & vbLf & vbLf, vbLf & vbLf)
    Loop
    CollapseBlankLines = result
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

Function JsonBool(value)
    If CBool(value) Then
        JsonBool = "true"
    Else
        JsonBool = "false"
    End If
End Function

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
