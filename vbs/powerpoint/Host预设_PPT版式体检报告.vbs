' 函数名: HostPptLayoutHealthReport
' 描述: 检查当前 PowerPoint 演示文稿的缺标题页、越界对象、空文本框和字体混用情况，并导出版式体检报告
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

    Main = HostPptLayoutHealthReport(appObj)
End Function

Function HostPptLayoutHealthReport(appObj)
    On Error Resume Next

    Dim pres
    Set pres = appObj.ActivePresentation
    If Err.Number <> 0 Or TypeName(pres) = "Empty" Or TypeName(pres) = "Nothing" Then
        Err.Clear
        HostPptLayoutHealthReport = "{""ok"":false,""code"":""E_NO_PRESENTATION"",""message"":""当前没有活动演示文稿""}"
        Exit Function
    End If

    Dim slideSummary, layoutIssueSummary, outputPlan, outputPath
    slideSummary = SafeHostText("GetPptSlideSummary")
    layoutIssueSummary = SafeHostText("GetPptLayoutIssueSummary")
    outputPlan = SafeResolveOutputPlan("ppt_layout_health_report.txt")
    outputPath = ExtractJsonString(outputPlan, "path")
    If Len(outputPath) = 0 Then outputPath = Host.ResolveTempPath("ppt_layout_health_report.txt")

    Dim missingTitleCount, offCanvasCount, emptyTextCount, fontIssueCount, report
    report = BuildLayoutReport(pres, missingTitleCount, offCanvasCount, emptyTextCount, fontIssueCount)
    Host.WriteTextFile outputPath, report, True

    Dim summary
    summary = "PPT 版式体检完成；缺标题页=" & CStr(missingTitleCount) & _
        "；越界对象=" & CStr(offCanvasCount) & "；空文本框=" & CStr(emptyTextCount) & _
        "；字体混用页=" & CStr(fontIssueCount) & "；输出=" & outputPath & vbCrLf & _
        "Host.GetPptSlideSummary: " & slideSummary & vbCrLf & _
        "Host.GetPptLayoutIssueSummary: " & layoutIssueSummary
    Host.WriteClipboard summary & vbCrLf & report
    SafeWriteLog summary

    HostPptLayoutHealthReport = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""missingTitleCount"":" & CStr(missingTitleCount) & _
        ",""offCanvasCount"":" & CStr(offCanvasCount) & ",""emptyTextCount"":" & CStr(emptyTextCount) & _
        ",""fontIssueCount"":" & CStr(fontIssueCount) & ",""outputPath"":""" & EscapeJson(outputPath) & """}"

End Function

Function BuildLayoutReport(pres, ByRef missingTitleCount, ByRef offCanvasCount, ByRef emptyTextCount, ByRef fontIssueCount)
    On Error Resume Next
    Dim report, slide, issueText, slideFontCount
    missingTitleCount = 0
    offCanvasCount = 0
    emptyTextCount = 0
    fontIssueCount = 0
    report = "PPT 版式体检报告" & vbCrLf & String(24, "=") & vbCrLf
    For Each slide In pres.Slides
        issueText = CheckSlideLayout(slide, pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight, offCanvasCount, emptyTextCount, slideFontCount)
        If Not SlideHasTitle(slide) Then
            missingTitleCount = missingTitleCount + 1
            issueText = issueText & "缺少标题；"
        End If
        If slideFontCount > 3 Then
            fontIssueCount = fontIssueCount + 1
            issueText = issueText & "字体种类较多(" & CStr(slideFontCount) & ")；"
        End If
        If Len(issueText) > 0 Then
            report = report & "第 " & CStr(slide.SlideIndex) & " 页: " & issueText & vbCrLf
        End If
        If Err.Number <> 0 Then Err.Clear
    Next
    If missingTitleCount + offCanvasCount + emptyTextCount + fontIssueCount = 0 Then
        report = report & "未发现明显版式问题。" & vbCrLf
    End If
    BuildLayoutReport = report
End Function

Function CheckSlideLayout(slide, slideWidth, slideHeight, ByRef offCanvasCount, ByRef emptyTextCount, ByRef slideFontCount)
    On Error Resume Next
    Dim shape, issues, fonts(), fontCount
    issues = ""
    fontCount = 0
    For Each shape In slide.Shapes
        If shape.Left < -1 Or shape.Top < -1 Or shape.Left + shape.Width > slideWidth + 1 Or shape.Top + shape.Height > slideHeight + 1 Then
            offCanvasCount = offCanvasCount + 1
            issues = issues & "对象越界(" & shape.Name & ")；"
        End If
        If shape.HasTextFrame = msoTrue Then
            If shape.TextFrame.HasText = msoTrue Then
                AddFont fonts, fontCount, CStr(shape.TextFrame.TextRange.Font.Name)
                If Len(Trim(Replace(Replace(shape.TextFrame.TextRange.Text, vbCr, ""), vbLf, ""))) = 0 Then
                    emptyTextCount = emptyTextCount + 1
                    issues = issues & "空文本框(" & shape.Name & ")；"
                End If
            Else
                emptyTextCount = emptyTextCount + 1
                issues = issues & "空文本框(" & shape.Name & ")；"
            End If
        End If
        If Err.Number <> 0 Then Err.Clear
    Next
    slideFontCount = fontCount
    CheckSlideLayout = issues
End Function

Function SlideHasTitle(slide)
    On Error Resume Next
    Dim titleShape
    SlideHasTitle = False
    Set titleShape = slide.Shapes.Title
    If Err.Number = 0 Then
        If titleShape.HasTextFrame = msoTrue Then
            SlideHasTitle = (Len(Trim(titleShape.TextFrame.TextRange.Text)) > 0)
        End If
    End If
    Err.Clear
End Function

Sub AddFont(ByRef fonts, ByRef fontCount, fontName)
    Dim i, normalized
    normalized = Trim(CStr(fontName))
    If Len(normalized) = 0 Then Exit Sub
    For i = 0 To fontCount - 1
        If StrComp(fonts(i), normalized, vbTextCompare) = 0 Then Exit Sub
    Next
    ReDim Preserve fonts(fontCount)
    fonts(fontCount) = normalized
    fontCount = fontCount + 1
End Sub

Function SafeHostText(methodName)
    On Error Resume Next
    If methodName = "GetPptSlideSummary" Then
        SafeHostText = Host.GetPptSlideSummary()
    ElseIf methodName = "GetPptLayoutIssueSummary" Then
        SafeHostText = Host.GetPptLayoutIssueSummary()
    Else
        SafeHostText = ""
    End If
    If Err.Number <> 0 Then
        SafeHostText = ""
        Err.Clear
    End If
End Function

Function SafeResolveOutputPlan(defaultName)
    On Error Resume Next
    SafeResolveOutputPlan = Host.ResolveOutputPlan(defaultName, "overwrite")
    If Err.Number <> 0 Then
        SafeResolveOutputPlan = ""
        Err.Clear
    End If
End Function

Function ExtractJsonString(jsonText, fieldName)
    Dim pattern, startAt, colonAt, quoteAt, i, ch, result
    pattern = """" & fieldName & """"
    startAt = InStr(1, jsonText, pattern, vbTextCompare)
    If startAt <= 0 Then
        ExtractJsonString = ""
        Exit Function
    End If
    colonAt = InStr(startAt + Len(pattern), jsonText, ":")
    quoteAt = InStr(colonAt + 1, jsonText, """")
    If colonAt <= 0 Or quoteAt <= 0 Then
        ExtractJsonString = ""
        Exit Function
    End If
    result = ""
    For i = quoteAt + 1 To Len(jsonText)
        ch = Mid(jsonText, i, 1)
        If ch = """" Then Exit For
        If ch = "\" And i < Len(jsonText) Then
            i = i + 1
            ch = Mid(jsonText, i, 1)
            If ch = "n" Then
                result = result & vbCrLf
            Else
                result = result & ch
            End If
        Else
            result = result & ch
        End If
    Next
    ExtractJsonString = result
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
