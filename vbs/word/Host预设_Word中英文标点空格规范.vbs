' 函数名: HostWordNormalizeCnEnPunctuationSpacing
' 描述: 规范 Word 当前选区中的中英文标点、全角空格和常见多余空格；适合正文细排前的文本清理
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

    Main = HostWordNormalizeCnEnPunctuationSpacing(appObj)
End Function

Function HostWordNormalizeCnEnPunctuationSpacing(appObj)
    On Error Resume Next

    Dim sourceText
    sourceText = Host.ExtractRange("selection")
    If Err.Number <> 0 Then
        Err.Clear
        HostWordNormalizeCnEnPunctuationSpacing = "{""ok"":false,""code"":""E_EXTRACT_RANGE"",""message"":""无法读取当前选区，请先选择 Word 文本""}"
        Exit Function
    End If
    If Len(sourceText) = 0 Then
        HostWordNormalizeCnEnPunctuationSpacing = "{""ok"":false,""code"":""E_EMPTY_SELECTION"",""message"":""当前选区为空""}"
        Exit Function
    End If

    Dim normalized, changedCount
    normalized = sourceText
    changedCount = 0
    normalized = ReplaceAndCount(normalized, ChrW(12288), " ", changedCount)
    normalized = ReplaceAndCount(normalized, " ,", ",", changedCount)
    normalized = ReplaceAndCount(normalized, " .", ".", changedCount)
    normalized = ReplaceAndCount(normalized, " ;", ";", changedCount)
    normalized = ReplaceAndCount(normalized, " :", ":", changedCount)
    normalized = ReplaceAndCount(normalized, " ，", "，", changedCount)
    normalized = ReplaceAndCount(normalized, " 。", "。", changedCount)
    normalized = ReplaceAndCount(normalized, " ；", "；", changedCount)
    normalized = ReplaceAndCount(normalized, " ：", "：", changedCount)
    normalized = CollapseRepeatedSpaces(normalized, changedCount)

    If normalized = sourceText Then
        HostWordNormalizeCnEnPunctuationSpacing = "{""ok"":true,""changed"":false,""message"":""选区标点空格已经较规范"",""length"":" & CStr(Len(sourceText)) & "}"
        Exit Function
    End If

    Dim planId, previewJson, validationJson
    planId = SafeBeginWritePlan()
    SafeRecordWrite "word_normalize_punctuation_spacing", "{""target"":""selection"",""textLength"":" & CStr(Len(normalized)) & "}", "office.word.textNormalize"
    validationJson = SafeValidateWritePlan(planId)
    previewJson = SafePreviewWritePlan()

    Dim ok
    ok = Host.WriteRange(normalized, "selection")
    SafeCloseWritePlan planId

    Dim summary
    summary = "Word 中英文标点空格规范完成；替换项=" & CStr(changedCount) & _
        "；原长度=" & CStr(Len(sourceText)) & "；新长度=" & CStr(Len(normalized))
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostWordNormalizeCnEnPunctuationSpacing = "{""ok"":" & JsonBool(ok) & ",""changed"":true,""message"":""" & EscapeJson(summary) & _
        """,""changedCount"":" & CStr(changedCount) & ",""preview"":""" & EscapeJson(previewJson) & _
        """,""validation"":""" & EscapeJson(validationJson) & """}"

End Function

Function ReplaceAndCount(text, findText, replaceText, ByRef changedCount)
    Dim beforeText, afterText, localCount
    beforeText = CStr(text)
    localCount = CountOccurrences(beforeText, findText)
    afterText = Replace(beforeText, findText, replaceText)
    changedCount = changedCount + localCount
    ReplaceAndCount = afterText
End Function

Function CollapseRepeatedSpaces(text, ByRef changedCount)
    Dim result
    result = CStr(text)
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
        changedCount = changedCount + 1
    Loop
    CollapseRepeatedSpaces = result
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
        pos = InStr(startAt, text, needle, vbBinaryCompare)
        If pos <= 0 Then Exit Do
        count = count + 1
        startAt = pos + Len(needle)
    Loop
    CountOccurrences = count
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
