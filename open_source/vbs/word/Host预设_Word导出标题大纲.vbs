' 函数名: HostWordExportHeadingOutline
' 描述: 提取当前 Word 文档标题段落为大纲，复制到剪贴板并写入文本文件；验证 Word 上下文、ResolveOutputPlan 和 WriteTextFile
' 适用应用: Word
' 搜索范围: 全文
' 搜索对象: 无

Option Explicit

Function Main()
    On Error Resume Next

    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = "{""ok"":false,""code"":""E_NO_WORD_APP"",""message"":""未取得 Word 应用，请在 Word 中运行该预设""}"
        Exit Function
    End If

    Main = HostWordExportHeadingOutline(appObj)
End Function

Function HostWordExportHeadingOutline(appObj)
    On Error Resume Next

    Dim doc
    Set doc = appObj.ActiveDocument
    If Err.Number <> 0 Or TypeName(doc) = "Empty" Or TypeName(doc) = "Nothing" Then
        Err.Clear
        HostWordExportHeadingOutline = "{""ok"":false,""code"":""E_NO_DOCUMENT"",""message"":""当前没有活动 Word 文档""}"
        Exit Function
    End If

    Dim wordContext, outline
    wordContext = SafeHostText("GetWordContextInfo")
    outline = BuildHeadingOutline(doc, wordContext)

    Dim outputJson, outputPath, written
    outputJson = Host.ResolveOutputPlan("word_heading_outline.txt", "avoid")
    outputPath = JsonStringValue(outputJson, "path")
    If Len(outputPath) = 0 Then outputPath = Host.ResolveTempPath("word_heading_outline.txt")
    written = Host.WriteTextFile(outputPath, outline, True)
    Host.WriteClipboard outline
    SafeWriteLog "Word heading outline exported: " & outputPath

    HostWordExportHeadingOutline = "{""ok"":" & JsonBool(written) & ",""message"":""Word 标题大纲已导出"",""outputPath"":""" & EscapeJson(outputPath) & """}"

End Function

Function BuildHeadingOutline(doc, wordContext)
    On Error Resume Next
    Dim text, para, titleText, levelText, outlineLevel
    text = "Word 标题大纲" & vbCrLf & _
        "Document: " & doc.Name & vbCrLf & _
        "Host.GetWordContextInfo: " & wordContext & vbCrLf & vbCrLf

    For Each para In doc.Paragraphs
        titleText = Trim(Replace(para.Range.Text, vbCr, ""))
        If Len(titleText) > 0 Then
            outlineLevel = 10
            Err.Clear
            outlineLevel = para.OutlineLevel
            Err.Clear
            If outlineLevel >= 1 And outlineLevel <= 9 Then
                levelText = RepeatText("#", outlineLevel)
                text = text & levelText & " " & titleText & vbCrLf
            End If
        End If
    Next

    If Right(text, 2) = vbCrLf And InStr(text, "# ") = 0 Then
        text = text & "(未检测到标题级别段落)" & vbCrLf
    End If
    BuildHeadingOutline = text
End Function

Function RepeatText(token, count)
    Dim i, text
    text = ""
    For i = 1 To count
        text = text & token
    Next
    RepeatText = text
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
