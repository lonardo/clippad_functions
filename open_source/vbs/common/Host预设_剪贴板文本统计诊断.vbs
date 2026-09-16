' 函数名: HostClipboardTextDiagnostics
' 描述: 读取剪贴板文本并生成行数、字符数和工作量计划诊断；验证 Host.ReadClipboard、CountTextLines、LimitText、GetTextWorkloadPlan 和 WriteClipboard
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

    Main = HostClipboardTextDiagnostics(appObj)
End Function

Function HostClipboardTextDiagnostics(appObj)
    On Error Resume Next

    Dim sourceText
    sourceText = Host.ReadClipboard()
    If Err.Number <> 0 Then
        Err.Clear
        HostClipboardTextDiagnostics = "{""ok"":false,""code"":""E_CLIPBOARD_READ"",""message"":""读取剪贴板失败""}"
        Exit Function
    End If

    Dim lineCount, charCount, previewText, workloadJson, report
    charCount = Len(sourceText)
    lineCount = Host.CountTextLines(sourceText)
    previewText = Host.LimitText(sourceText, 500, "...")
    workloadJson = Host.GetTextWorkloadPlan(sourceText, 20000, 4000, 200)

    report = "剪贴板文本诊断" & vbCrLf & _
        "字符数: " & CStr(charCount) & vbCrLf & _
        "行数: " & CStr(lineCount) & vbCrLf & _
        "Host.GetTextWorkloadPlan: " & workloadJson & vbCrLf & vbCrLf & _
        "预览:" & vbCrLf & previewText
    Host.WriteClipboard report
    SafeWriteLog report

    HostClipboardTextDiagnostics = "{""ok"":true,""message"":""剪贴板文本诊断已复制"",""charCount"":" & CStr(charCount) & _
        ",""lineCount"":" & CStr(lineCount) & ",""workload"":""" & EscapeJson(workloadJson) & """}"

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
