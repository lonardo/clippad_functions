' 函数名: HostPathOutputPlanDiagnostics
' 描述: 生成安全文件名、临时路径和输出计划诊断并复制到剪贴板；验证 Host.NormalizePath、GetPathParts、SanitizeFileName、EnsureFileExtension 和 ResolveOutputPlan
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

    Main = HostPathOutputPlanDiagnostics(appObj)
End Function

Function HostPathOutputPlanDiagnostics(appObj)
    On Error Resume Next

    Dim rawName, safeName, ensuredName, tempPath, normalizedPath, partsJson, outputJson, report
    rawName = SafePrompt("请输入输出文件名", "office report:demo?.txt")
    safeName = Host.SanitizeFileName(rawName, "office_report")
    ensuredName = Host.EnsureFileExtension(safeName, ".txt")
    tempPath = Host.ResolveTempPath(ensuredName)
    normalizedPath = Host.NormalizePath(tempPath)
    partsJson = Host.GetPathParts(normalizedPath)
    outputJson = Host.ResolveOutputPlan(ensuredName, "avoid")

    report = "路径与输出计划诊断" & vbCrLf & _
        "原始文件名: " & rawName & vbCrLf & _
        "安全文件名: " & safeName & vbCrLf & _
        "扩展名补全: " & ensuredName & vbCrLf & _
        "临时路径: " & tempPath & vbCrLf & _
        "标准化路径: " & normalizedPath & vbCrLf & _
        "Host.GetPathParts: " & partsJson & vbCrLf & _
        "Host.ResolveOutputPlan: " & outputJson
    Host.WriteClipboard report
    SafeWriteLog report

    HostPathOutputPlanDiagnostics = "{""ok"":true,""message"":""路径与输出计划诊断已复制"",""safeName"":""" & EscapeJson(safeName) & _
        """,""path"":""" & EscapeJson(normalizedPath) & """}"

End Function

Function SafePrompt(message, defaultValue)
    On Error Resume Next
    SafePrompt = Host.Prompt(message, defaultValue)
    If Err.Number <> 0 Then
        SafePrompt = defaultValue
        Err.Clear
    End If
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
