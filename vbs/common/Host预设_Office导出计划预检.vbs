' 函数名: HostOfficeExportPlanPreflight
' 描述: 基于当前 Office 上下文生成 PDF 导出计划预检并复制诊断；验证 Host.GetContextInfo、GetRunPreflightPlan、GetExportPlan 和 WriteClipboard
' 适用应用: Word|Excel|PowerPoint
' 搜索范围: 全文
' 搜索对象: 无

Option Explicit

Function Main()
    On Error Resume Next

    Dim appObj
    Set appObj = Nothing
    Err.Clear
    Set appObj = Host.GetApplication()
    Err.Clear

    Main = HostOfficeExportPlanPreflight(appObj)
End Function

Function HostOfficeExportPlanPreflight(appObj)
    On Error Resume Next

    Dim contextJson, preflightJson, exportJson, report
    contextJson = Host.GetContextInfo()
    preflightJson = Host.GetRunPreflightPlan("", True, False, False)
    exportJson = Host.GetExportPlan("pdf", "office_export.pdf", "avoid")

    report = "Office 导出计划预检" & vbCrLf & _
        "Host.GetContextInfo: " & contextJson & vbCrLf & _
        "Host.GetRunPreflightPlan: " & preflightJson & vbCrLf & _
        "Host.GetExportPlan: " & exportJson
    Host.WriteClipboard report
    SafeWriteLog report

    HostOfficeExportPlanPreflight = "{""ok"":true,""message"":""Office 导出计划预检已复制"",""exportPlan"":""" & EscapeJson(exportJson) & """}"

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
