' 函数名: HostEnvironmentFileInfoDiagnostics
' 描述: 读取常用环境变量、日志路径和文件信息计划诊断并复制到剪贴板；验证 Host.GetEnvironmentVariable、ExpandEnvironmentStrings、ResolveLogPath、PathExists 和 GetFileInfo
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

    Main = HostEnvironmentFileInfoDiagnostics(appObj)
End Function

Function HostEnvironmentFileInfoDiagnostics(appObj)
    On Error Resume Next

    Dim tempEnv, appDataEnv, expandedTemp, logPath, knownTemp, existsText, infoJson, report
    tempEnv = Host.GetEnvironmentVariable("TEMP")
    appDataEnv = Host.GetEnvironmentVariable("APPDATA")
    expandedTemp = Host.ExpandEnvironmentStrings("%TEMP%\officeaddin_env_probe.txt")
    knownTemp = Host.GetKnownFolderPath("temp")
    logPath = Host.ResolveLogPath("host_preset_diagnostics.log")
    existsText = CStr(Host.PathExists(logPath))
    infoJson = Host.GetFileInfo(logPath)

    report = "环境与文件信息诊断" & vbCrLf & _
        "TEMP: " & tempEnv & vbCrLf & _
        "APPDATA: " & appDataEnv & vbCrLf & _
        "Expanded temp probe: " & expandedTemp & vbCrLf & _
        "Known temp: " & knownTemp & vbCrLf & _
        "Log path: " & logPath & vbCrLf & _
        "Host.PathExists(log): " & existsText & vbCrLf & _
        "Host.GetFileInfo(log): " & infoJson
    Host.WriteClipboard report
    SafeWriteLog report

    HostEnvironmentFileInfoDiagnostics = "{""ok"":true,""message"":""环境与文件信息诊断已复制"",""logPath"":""" & EscapeJson(logPath) & _
        """,""exists"":""" & EscapeJson(existsText) & """}"

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
