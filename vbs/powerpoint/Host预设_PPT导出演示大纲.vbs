' 函数名: HostPptExportOutline
' 描述: 提取当前 PowerPoint 每页标题和文本为演示大纲，复制到剪贴板并写入文本文件；验证 PPT 上下文、ResolveOutputPlan 和 WriteTextFile
' 适用应用: PowerPoint
' 搜索范围: 全文
' 搜索对象: 无

Option Explicit

Const msoTrue = -1

Function Main()
    On Error Resume Next

    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = "{""ok"":false,""code"":""E_NO_PPT_APP"",""message"":""未取得 PowerPoint 应用，请在 PowerPoint 中运行该预设""}"
        Exit Function
    End If

    Main = HostPptExportOutline(appObj)
End Function

Function HostPptExportOutline(appObj)
    On Error Resume Next

    Dim pres
    Set pres = appObj.ActivePresentation
    If Err.Number <> 0 Or TypeName(pres) = "Empty" Or TypeName(pres) = "Nothing" Then
        Err.Clear
        HostPptExportOutline = "{""ok"":false,""code"":""E_NO_PRESENTATION"",""message"":""当前没有活动演示文稿""}"
        Exit Function
    End If

    Dim pptSummary, slideSummary, outline
    pptSummary = SafeHostText("GetPptContextInfo")
    slideSummary = SafeHostText("GetPptSlideSummary")
    outline = BuildOutline(pres, pptSummary, slideSummary)

    Dim outputJson, outputPath, written
    outputJson = Host.ResolveOutputPlan("ppt_outline.txt", "avoid")
    outputPath = JsonStringValue(outputJson, "path")
    If Len(outputPath) = 0 Then outputPath = Host.ResolveTempPath("ppt_outline.txt")
    written = Host.WriteTextFile(outputPath, outline, True)
    Host.WriteClipboard outline
    SafeWriteLog "PPT outline exported: " & outputPath

    HostPptExportOutline = "{""ok"":" & JsonBool(written) & ",""message"":""PPT 大纲已导出"",""slideCount"":" & CStr(pres.Slides.Count) & _
        ",""outputPath"":""" & EscapeJson(outputPath) & """}"

End Function

Function BuildOutline(pres, pptSummary, slideSummary)
    On Error Resume Next
    Dim text, slide, shape, slideText
    text = "PowerPoint 演示大纲" & vbCrLf & _
        "Presentation: " & pres.Name & vbCrLf & _
        "Host.GetPptContextInfo: " & pptSummary & vbCrLf & _
        "Host.GetPptSlideSummary: " & slideSummary & vbCrLf & vbCrLf

    For Each slide In pres.Slides
        text = text & "## Slide " & CStr(slide.SlideIndex) & vbCrLf
        slideText = ""
        For Each shape In slide.Shapes
            If shape.HasTextFrame = msoTrue Then
                If shape.TextFrame.HasText = msoTrue Then
                    slideText = slideText & Trim(shape.TextFrame.TextRange.Text) & vbCrLf
                End If
            End If
        Next
        If Len(Trim(slideText)) = 0 Then slideText = "(无文本)" & vbCrLf
        text = text & slideText & vbCrLf
    Next
    BuildOutline = text
End Function

Function SafeHostText(methodName)
    On Error Resume Next
    If methodName = "GetPptContextInfo" Then
        SafeHostText = Host.GetPptContextInfo()
    ElseIf methodName = "GetPptSlideSummary" Then
        SafeHostText = Host.GetPptSlideSummary()
    Else
        SafeHostText = ""
    End If
    If Err.Number <> 0 Then
        SafeHostText = ""
        Err.Clear
    End If
End Function

Function JsonStringValue(jsonText, key)
    Dim pattern, keyPos, colonPos, q1, q2
    pattern = """" & key & """"
    keyPos = InStr(1, jsonText, pattern, vbTextCompare)
    If keyPos <= 0 Then
        JsonStringValue = ""
        Exit Function
    End If
    colonPos = InStr(keyPos + Len(pattern), jsonText, ":")
    q1 = InStr(colonPos + 1, jsonText, """")
    q2 = InStr(q1 + 1, jsonText, """")
    If q1 <= 0 Or q2 <= q1 Then
        JsonStringValue = ""
    Else
        JsonStringValue = JsonUnescape(Mid(jsonText, q1 + 1, q2 - q1 - 1))
    End If
End Function

Function JsonUnescape(text)
    Dim t
    t = CStr(text)
    t = Replace(t, "\\", "\")
    t = Replace(t, "\""", """")
    t = Replace(t, "\r", vbCr)
    t = Replace(t, "\n", vbLf)
    t = Replace(t, "\t", vbTab)
    JsonUnescape = t
End Function

Sub SafeWriteLog(message)
    On Error Resume Next
    Host.WriteLog message
    Err.Clear
End Sub

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
