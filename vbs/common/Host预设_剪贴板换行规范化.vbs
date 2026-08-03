' 函数名: HostClipboardNormalizeLineEndings
' 描述: 读取剪贴板文本，统一换行后写回并复制统计摘要；验证 Host.NormalizeLineEndings、CountTextLines、ReadClipboard 和 WriteClipboard
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

    Main = HostClipboardNormalizeLineEndings(appObj)
End Function

Function HostClipboardNormalizeLineEndings(appObj)
    On Error Resume Next

    Dim sourceText, normalizedText, beforeLines, afterLines, report
    sourceText = Host.ReadClipboard()
    If Err.Number <> 0 Then
        Err.Clear
        HostClipboardNormalizeLineEndings = "{""ok"":false,""code"":""E_CLIPBOARD_READ"",""message"":""读取剪贴板失败""}"
        Exit Function
    End If

    beforeLines = Host.CountTextLines(sourceText)
    normalizedText = Host.NormalizeLineEndings(sourceText, "crlf")
    afterLines = Host.CountTextLines(normalizedText)
    Host.WriteClipboard normalizedText
    report = "剪贴板换行规范化完成；原行数=" & CStr(beforeLines) & "；规范化后行数=" & CStr(afterLines)
    SafeWriteLog report

    HostClipboardNormalizeLineEndings = "{""ok"":true,""message"":""" & EscapeJson(report) & """,""beforeLines"":" & CStr(beforeLines) & _
        ",""afterLines"":" & CStr(afterLines) & "}"

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
