' 函数名: HostWritePlanStatusDiagnostics
' 描述: 创建一条空写入计划并复制计划状态、预览、校验和最近计划诊断；验证 Host.GetWritePlanStatus、GetLastWritePlan、ValidateWritePlan 和 RollbackWritePlan
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

    Main = HostWritePlanStatusDiagnostics(appObj)
End Function

Function HostWritePlanStatusDiagnostics(appObj)
    On Error Resume Next

    Dim planId, statusBefore, previewJson, validationJson, lastPlanJson, statusAfter, report
    planId = Host.BeginWritePlan("")
    statusBefore = Host.GetWritePlanStatus(planId)
    Host.RecordWrite "diagnostic_noop", "{""scope"":""diagnostic""}", "host", "", "", "host.diagnostic.noop"
    previewJson = Host.PreviewWritePlan()
    validationJson = Host.ValidateWritePlan(planId, False)
    lastPlanJson = Host.GetLastWritePlan()
    Host.RollbackWritePlan planId
    statusAfter = Host.GetWritePlanStatus(planId)

    report = "写入计划状态诊断" & vbCrLf & _
        "PlanId: " & planId & vbCrLf & _
        "Before: " & statusBefore & vbCrLf & _
        "Preview: " & previewJson & vbCrLf & _
        "Validation: " & validationJson & vbCrLf & _
        "LastPlan: " & lastPlanJson & vbCrLf & _
        "AfterRollback: " & statusAfter
    Host.WriteClipboard report
    SafeWriteLog report

    HostWritePlanStatusDiagnostics = "{""ok"":true,""message"":""写入计划状态诊断已复制"",""planId"":""" & EscapeJson(planId) & """}"

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
