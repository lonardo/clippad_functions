' 函数名: HostWordExtractAllImages
' 描述: 将当前 Word 文档图片通过 HTML 中转导出到选定文件夹；对标稻壳/公文工具中的图片提取
' 适用应用: Word
' 搜索范围: 全文
' 搜索对象: 无

Option Explicit

Const wdFormatHTML = 8
Const wdDoNotSaveChanges = 0

Function Main()
    On Error Resume Next
    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = "{""ok"":false,""code"":""E_NO_WORD_APP"",""message"":""未取得 Word 应用，请在 Word 中运行该预设""}"
        Exit Function
    End If
    Main = HostWordExtractAllImages(appObj)
End Function

Function HostWordExtractAllImages(appObj)
    On Error Resume Next
    Dim doc
    Set doc = appObj.ActiveDocument
    If Err.Number <> 0 Or TypeName(doc) = "Empty" Or TypeName(doc) = "Nothing" Then
        Err.Clear
        HostWordExtractAllImages = "{""ok"":false,""code"":""E_NO_DOCUMENT"",""message"":""当前没有活动 Word 文档""}"
        Exit Function
    End If

    Dim outputFolder
    outputFolder = Host.SelectFolder("请选择图片导出文件夹")
    If Err.Number <> 0 Or Len(CStr(outputFolder)) = 0 Then
        Err.Clear
        HostWordExtractAllImages = "{""ok"":false,""code"":""E_NO_FOLDER"",""message"":""未选择导出文件夹""}"
        Exit Function
    End If

    Dim wordSummary, planId, previewJson, inlineCount, shapeCount
    wordSummary = SafeHostText("GetWordContextInfo")
    inlineCount = doc.InlineShapes.Count
    shapeCount = doc.Shapes.Count
    planId = SafeBeginWritePlan()
    SafeRecordWrite "word_extract_all_images", "{""inline"":" & CStr(inlineCount) & ",""shapes"":" & CStr(shapeCount) & "}", "office.word.image.export"
    previewJson = SafePreviewWritePlan()

    Dim tempHtml, newDoc, filesJson, exported
    tempHtml = Host.CombinePath(outputFolder, "HostWordImageExport_" & CStr(Year(Now)) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2) & ".html")
    Set newDoc = appObj.Documents.Add
    doc.Content.Copy
    newDoc.Content.Paste
    newDoc.SaveAs tempHtml, wdFormatHTML
    If Err.Number <> 0 Then
        Dim errMsg
        errMsg = Err.Description
        Err.Clear
        newDoc.Close wdDoNotSaveChanges
        SafeCloseWritePlan planId
        HostWordExtractAllImages = "{""ok"":false,""code"":""E_HTML_EXPORT"",""message"":""HTML 中转导出失败: " & EscapeJson(errMsg) & """}"
        Exit Function
    End If
    newDoc.Close wdDoNotSaveChanges

    filesJson = Host.EnumerateFiles(outputFolder, "*.*", True)
    exported = CountImageMentions(filesJson)
    SafeCloseWritePlan planId

    Dim summary
    summary = "Word 图片提取完成；内嵌图=" & CStr(inlineCount) & "；浮动图=" & CStr(shapeCount) & _
        "；导出目录=" & outputFolder & "；检测到图片文件线索=" & CStr(exported) & vbCrLf & _
        "说明: 图片通常位于 HTML 同级 .files/_files 文件夹；请手工删除临时 HTML 中转文件" & vbCrLf & _
        "Host.GetWordContextInfo: " & wordSummary & vbCrLf & _
        "Host.EnumerateFiles: " & filesJson & vbCrLf & _
        "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostWordExtractAllImages = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""inlineShapes"":" & CStr(inlineCount) & _
        ",""shapes"":" & CStr(shapeCount) & ",""imageHints"":" & CStr(exported) & ",""folder"":""" & EscapeJson(outputFolder) & """,""planId"":""" & EscapeJson(planId) & """}"
End Function

Function CountImageMentions(filesJson)
    On Error Resume Next
    Dim text, count, tokens, i, token
    text = LCase(CStr(filesJson))
    count = 0
    tokens = Array(".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".wmf", ".emf")
    For i = 0 To UBound(tokens)
        token = tokens(i)
        Dim pos, n
        pos = 1
        n = 0
        Do
            pos = InStr(pos, text, token, vbTextCompare)
            If pos = 0 Then Exit Do
            n = n + 1
            pos = pos + Len(token)
        Loop
        count = count + n
    Next
    CountImageMentions = count
End Function

Function SafeHostText(methodName)
    On Error Resume Next
    If methodName = "GetWordContextInfo" Then
        SafeHostText = Host.GetWordContextInfo()
    Else
        SafeHostText = ""
    End If
    If Err.Number <> 0 Then
        SafeHostText = ""
        Err.Clear
    End If
End Function
Function SafePrompt(message, defaultValue)
    On Error Resume Next
    SafePrompt = Host.Prompt(message, defaultValue)
    If Err.Number <> 0 Then
        SafePrompt = defaultValue
        Err.Clear
    End If
End Function

Function SafeBeginWritePlan()
    On Error Resume Next
    SafeBeginWritePlan = Host.BeginWritePlan("")
    If Err.Number <> 0 Then SafeBeginWritePlan = ""
    Err.Clear
End Function

Sub SafeRecordWrite(actionId, paramsJson, capabilityId)
    On Error Resume Next
    Host.RecordWrite actionId, paramsJson, "office", "", "", capabilityId
    Err.Clear
End Sub

Function SafePreviewWritePlan()
    On Error Resume Next
    SafePreviewWritePlan = Host.PreviewWritePlan()
    If Err.Number <> 0 Then SafePreviewWritePlan = ""
    Err.Clear
End Function

Sub SafeCloseWritePlan(planId)
    On Error Resume Next
    Host.RollbackWritePlan planId
    Err.Clear
End Sub

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