' 函数名: HostContextDiagnostics
' 描述: 汇总当前 Host 方法、能力目录、文档/选区/Office 上下文，并复制到剪贴板；用于验证 VBS 已接入本地 Host 能力
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

    Main = HostContextDiagnostics(appObj)
End Function

Function HostContextDiagnostics(appObj)
    On Error Resume Next

    Dim hostInfo, docInfo, selectionInfo, contextInfo
    Dim excelInfo, wordInfo, pptInfo, methodsInfo, capabilitiesInfo
    hostInfo = SafeHostText("GetHostInfo")
    docInfo = SafeHostText("GetDocumentInfo")
    selectionInfo = SafeHostText("GetSelectionInfo")
    contextInfo = SafeHostText("GetContextInfo")
    excelInfo = SafeHostText("GetExcelContextInfo")
    wordInfo = SafeHostText("GetWordContextInfo")
    pptInfo = SafeHostText("GetPptContextInfo")
    methodsInfo = SafeHostText("ListHostMethods")
    capabilitiesInfo = SafeHostText("ListCapabilities")

    Dim report
    report = "OfficeAddin Host context diagnostics" & vbCrLf & _
        "HostInfo: " & hostInfo & vbCrLf & _
        "DocumentInfo: " & docInfo & vbCrLf & _
        "SelectionInfo: " & selectionInfo & vbCrLf & _
        "ContextInfo: " & contextInfo & vbCrLf & _
        "ExcelContext: " & excelInfo & vbCrLf & _
        "WordContext: " & wordInfo & vbCrLf & _
        "PptContext: " & pptInfo & vbCrLf & _
        "Methods: " & methodsInfo & vbCrLf & _
        "Capabilities: " & capabilitiesInfo

    Dim copied
    copied = False
    Err.Clear
    copied = Host.WriteClipboard(report)
    If Err.Number <> 0 Then
        copied = False
        Err.Clear
    End If

    Err.Clear
    Host.WriteLog report
    Err.Clear

    HostContextDiagnostics = "{""ok"":true,""message"":""Host 上下文诊断已生成"",""copied"":" & JsonBool(copied) & _
        ",""hostInfo"":""" & EscapeJson(hostInfo) & _
        """,""documentInfo"":""" & EscapeJson(docInfo) & _
        """,""selectionInfo"":""" & EscapeJson(selectionInfo) & """}"

End Function

Function SafeHostText(methodName)
    On Error Resume Next
    Err.Clear
    Select Case methodName
        Case "GetHostInfo"
            SafeHostText = Host.GetHostInfo()
        Case "GetDocumentInfo"
            SafeHostText = Host.GetDocumentInfo()
        Case "GetSelectionInfo"
            SafeHostText = Host.GetSelectionInfo()
        Case "GetContextInfo"
            SafeHostText = Host.GetContextInfo()
        Case "GetExcelContextInfo"
            SafeHostText = Host.GetExcelContextInfo()
        Case "GetWordContextInfo"
            SafeHostText = Host.GetWordContextInfo()
        Case "GetPptContextInfo"
            SafeHostText = Host.GetPptContextInfo()
        Case "ListHostMethods"
            SafeHostText = Host.ListHostMethods()
        Case "ListCapabilities"
            SafeHostText = Host.ListCapabilities()
        Case Else
            SafeHostText = ""
    End Select
    If Err.Number <> 0 Then
        SafeHostText = ""
        Err.Clear
    End If
End Function

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
