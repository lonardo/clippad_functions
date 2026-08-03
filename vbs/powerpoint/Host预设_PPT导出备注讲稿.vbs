' 函数名: HostPptExportSpeakerNotes
' 描述: 导出当前 PowerPoint 演示文稿每页标题和备注讲稿到文本文件，并复制摘要；适合演讲稿整理、复盘和培训材料沉淀
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

    Main = HostPptExportSpeakerNotes(appObj)
End Function

Function HostPptExportSpeakerNotes(appObj)
    On Error Resume Next

    Dim pres
    Set pres = appObj.ActivePresentation
    If Err.Number <> 0 Or TypeName(pres) = "Empty" Or TypeName(pres) = "Nothing" Then
        Err.Clear
        HostPptExportSpeakerNotes = "{""ok"":false,""code"":""E_NO_PRESENTATION"",""message"":""当前没有活动演示文稿""}"
        Exit Function
    End If

    Dim pptContext, outputPlan, outputPath
    pptContext = SafeHostText("GetPptContextInfo")
    outputPlan = SafeResolveOutputPlan("ppt_speaker_notes.txt")
    outputPath = ExtractJsonString(outputPlan, "path")
    If Len(outputPath) = 0 Then outputPath = Host.ResolveTempPath("ppt_speaker_notes.txt")

    Dim notesText, noteCount
    notesText = BuildNotesText(pres, noteCount)
    Host.WriteTextFile outputPath, notesText, True

    Dim summary
    summary = "PPT 备注讲稿导出完成；页数=" & CStr(pres.Slides.Count) & "；含备注页=" & CStr(noteCount) & _
        "；输出=" & outputPath & vbCrLf & _
        "Host.GetPptContextInfo: " & pptContext & vbCrLf & _
        "Host.ResolveOutputPlan: " & outputPlan
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostPptExportSpeakerNotes = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""slideCount"":" & CStr(pres.Slides.Count) & _
        ",""noteCount"":" & CStr(noteCount) & ",""outputPath"":""" & EscapeJson(outputPath) & """}"

End Function

Function BuildNotesText(pres, ByRef noteCount)
    On Error Resume Next
    Dim i, slide, titleText, notesText, result
    result = "PPT 备注讲稿导出" & vbCrLf & String(24, "=") & vbCrLf
    noteCount = 0
    For i = 1 To pres.Slides.Count
        Set slide = pres.Slides(i)
        titleText = Trim(GetSlideTitle(slide))
        If Len(titleText) = 0 Then titleText = "未命名幻灯片"
        notesText = Trim(GetSlideNotes(slide))
        If Len(notesText) > 0 Then noteCount = noteCount + 1
        result = result & vbCrLf & "第 " & CStr(i) & " 页｜" & titleText & vbCrLf & String(24, "-") & vbCrLf
        If Len(notesText) > 0 Then
            result = result & notesText & vbCrLf
        Else
            result = result & "(无备注)" & vbCrLf
        End If
    Next
    BuildNotesText = result
End Function

Function GetSlideTitle(slide)
    On Error Resume Next
    Dim titleShape
    GetSlideTitle = ""
    Set titleShape = slide.Shapes.Title
    If Err.Number = 0 Then
        If titleShape.HasTextFrame = msoTrue Then
            GetSlideTitle = titleShape.TextFrame.TextRange.Text
        End If
    End If
    Err.Clear
End Function

Function GetSlideNotes(slide)
    On Error Resume Next
    Dim shape, textValue
    GetSlideNotes = ""
    For Each shape In slide.NotesPage.Shapes
        If shape.HasTextFrame = msoTrue Then
            If shape.TextFrame.HasText = msoTrue Then
                textValue = Trim(shape.TextFrame.TextRange.Text)
                If Len(textValue) > 0 Then
                    If InStr(1, textValue, "Click to add notes", vbTextCompare) = 0 Then
                        If Len(GetSlideNotes) > 0 Then GetSlideNotes = GetSlideNotes & vbCrLf
                        GetSlideNotes = GetSlideNotes & textValue
                    End If
                End If
            End If
        End If
    Next
    If Err.Number <> 0 Then
        GetSlideNotes = ""
        Err.Clear
    End If
End Function

Function SafeHostText(methodName)
    On Error Resume Next
    If methodName = "GetPptContextInfo" Then
        SafeHostText = Host.GetPptContextInfo()
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
    If colonAt <= 0 Then
        ExtractJsonString = ""
        Exit Function
    End If
    quoteAt = InStr(colonAt + 1, jsonText, """")
    If quoteAt <= 0 Then
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
