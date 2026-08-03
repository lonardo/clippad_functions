' 函数名: Host.ContextInfo 快照示例
' Description: Return the current Office context as Host-provided JSON.
' 适用应用: Word/Excel/PowerPoint
' License: Apache-2.0

Function Main()
    On Error Resume Next
    Err.Clear

    Dim contextJson
    contextJson = Host.GetContextInfo()
    If Err.Number <> 0 Or Len(contextJson) = 0 Then
        Main = "{""ok"":false,""code"":""E_CONTEXT_INFO"",""message"":""无法读取当前 Office 上下文 / Unable to read Office context""}"
        Exit Function
    End If

    Main = contextJson
End Function
