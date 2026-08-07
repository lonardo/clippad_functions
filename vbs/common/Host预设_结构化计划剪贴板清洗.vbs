' 函数名: HostStructuredClipboardCleanup
' 描述: 读取剪贴板文本，规范化空白后通过 Host.ExecutePlanJson 执行 clipboard_write 结构化计划写回剪贴板
' 适用应用: Word|Excel|PowerPoint
' 搜索范围: 无
' 搜索对象: 无

Option Explicit

Function Main()
    On Error Resume Next

    Dim appObj
    Set appObj = Nothing
    Err.Clear
    Set appObj = Host.GetApplication()
    Err.Clear

    Main = HostStructuredClipboardCleanup(appObj)
End Function

Function HostStructuredClipboardCleanup(appObj)
    On Error Resume Next

    Dim sourceText
    sourceText = Host.ReadClipboard()
    If Err.Number <> 0 Then
        Err.Clear
        HostStructuredClipboardCleanup = "{""ok"":false,""code"":""E_READ_CLIPBOARD"",""message"":""无法读取剪贴板""}"
        Exit Function
    End If

    If Len(sourceText) = 0 Then
        HostStructuredClipboardCleanup = "{""ok"":false,""code"":""E_EMPTY_CLIPBOARD"",""message"":""剪贴板为空""}"
        Exit Function
    End If

    Dim cleaned
    cleaned = sourceText
    cleaned = Replace(cleaned, ChrW(12288), " ")
    cleaned = Replace(cleaned, vbTab, " ")
    cleaned = CollapseRepeatedSpaces(cleaned)
    cleaned = CollapseBlankLines(cleaned)

    Dim planJson, resultJson
    planJson = "{""runtime"":""oa.executionPlan"",""planId"":""preset_clipboard_cleanup"",""sourceRoute"":""vbs_host_preset"",""steps"":[{""stepId"":""write_clipboard"",""executor"":""host"",""action"":""clipboard_write"",""capabilityId"":""clipboard_write"",""params"":{""text"":""" & EscapeJson(cleaned) & """},""allowInteractive"":false}]}"
    resultJson = Host.ExecutePlanJson(planJson)
    SafeWriteLog "Structured clipboard cleanup result: " & resultJson

    HostStructuredClipboardCleanup = "{""ok"":true,""message"":""剪贴板文本已通过结构化 Host 计划清洗写回"",""originalLength"":" & CStr(Len(sourceText)) & _
        ",""cleanedLength"":" & CStr(Len(cleaned)) & ",""dispatcherResult"":""" & EscapeJson(resultJson) & """}"

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
