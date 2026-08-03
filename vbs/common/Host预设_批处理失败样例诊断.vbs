' 函数名: HostBatchFailureSampleDiagnostics
' 描述: 生成批处理计划和失败样例 JSON 并复制到剪贴板；验证 Host.GetBatchPlan、GetBatchFailureJson 和 WriteLog
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

    Main = HostBatchFailureSampleDiagnostics(appObj)
End Function

Function HostBatchFailureSampleDiagnostics(appObj)
    On Error Resume Next

    Dim itemCount, batchPlanJson, failureJson, report
    itemCount = CLng(SafePrompt("请输入批处理项目数量", "125"))
    If itemCount < 0 Then itemCount = 0
    batchPlanJson = Host.GetBatchPlan(itemCount, 50, 1000, "diagnostic_batch")
    failureJson = Host.GetBatchFailureJson(2, "item-2", "示例失败：目标不可写", "E_SAMPLE_FAILURE", "这是诊断样例，不会实际执行文件操作")

    report = "批处理失败样例诊断" & vbCrLf & _
        "ItemCount: " & CStr(itemCount) & vbCrLf & _
        "Host.GetBatchPlan: " & batchPlanJson & vbCrLf & _
        "Host.GetBatchFailureJson: " & failureJson
    Host.WriteClipboard report
    SafeWriteLog report

    HostBatchFailureSampleDiagnostics = "{""ok"":true,""message"":""批处理失败样例诊断已复制"",""itemCount"":" & CStr(itemCount) & _
        ",""failure"":""" & EscapeJson(failureJson) & """}"

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
