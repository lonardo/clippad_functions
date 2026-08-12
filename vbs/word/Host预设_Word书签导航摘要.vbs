' 函数名: HostWordBookmarkNavSummary
' 描述: 只读汇总当前文档全部书签名称、位置样例和数量；可选按名称跳转，不修改文档内容
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
        Main = FailureJson("E_NO_WORD_APP", "未取得 Word 应用，请在 Word 中运行该预设")
        Exit Function
    End If
    Main = HostWordBookmarkNavSummary(appObj)
End Function

Function HostWordBookmarkNavSummary(appObj)
    On Error Resume Next

    Dim doc
    Set doc = appObj.ActiveDocument
    If Err.Number <> 0 Or TypeName(doc) = "Empty" Or TypeName(doc) = "Nothing" Then
        Err.Clear
        HostWordBookmarkNavSummary = FailureJson("E_NO_DOCUMENT", "当前没有活动 Word 文档")
        Exit Function
    End If

    Dim bmCount, hiddenCount, sampleLines, names(), nameCount
    bmCount = 0
    hiddenCount = 0
    nameCount = 0
    sampleLines = ""
    CollectBookmarks doc, bmCount, hiddenCount, sampleLines, names, nameCount

    Dim wordSummary, previewText
    wordSummary = SafeHostText("GetWordContextInfo")
    previewText = "Word 书签导航摘要（只读，尚未写入/不修改文档）" & vbCrLf & _
        "书签总数=" & CStr(bmCount) & "；隐藏/系统书签约=" & CStr(hiddenCount) & "；可导航=" & CStr(nameCount) & vbCrLf & _
        "样例：" & vbCrLf & sampleLines & vbCrLf & _
        "Host.GetWordContextInfo: " & wordSummary

    Dim jumpName, jumped
    jumpName = Trim(SafePrompt(previewText & vbCrLf & vbCrLf & "如需跳转请输入书签名称；直接确定则只复制摘要", ""))
    jumped = False
    If Len(jumpName) > 0 Then
        jumped = TryGotoBookmark(doc, jumpName)
        If Not jumped Then
            Host.WriteClipboard previewText & vbCrLf & "跳转失败：未找到书签 " & jumpName
            SafeWriteLog previewText
            HostWordBookmarkNavSummary = FailureJson("E_BOOKMARK_NOT_FOUND", "未找到书签：" & jumpName & "；摘要已复制，文档未改")
            Exit Function
        End If
    End If

    Dim summary
    summary = previewText
    If jumped Then summary = summary & vbCrLf & "已跳转到书签=" & jumpName
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostWordBookmarkNavSummary = "{""ok"":true,""readonly"":true,""bookmarkCount"":" & CStr(bmCount) & ",""navigableCount"":" & CStr(nameCount) & ",""jumped"":" & LCase(CStr(jumped)) & ",""message"":""" & EscapeJson(summary) & """}"
End Function

Sub CollectBookmarks(doc, ByRef bmCount, ByRef hiddenCount, ByRef sampleLines, ByRef names, ByRef nameCount)
    On Error Resume Next
    Dim bm, nm, startPos, snippet, line, shown
    bmCount = 0
    hiddenCount = 0
    nameCount = 0
    sampleLines = ""
    shown = 0
    For Each bm In doc.Bookmarks
        bmCount = bmCount + 1
        nm = ""
        nm = CStr(bm.Name)
        If Err.Number <> 0 Then
            Err.Clear
            nm = ""
        End If
        If Len(nm) = 0 Then
            hiddenCount = hiddenCount + 1
        ElseIf Left(nm, 1) = "_" Then
            hiddenCount = hiddenCount + 1
        Else
            ReDim Preserve names(nameCount)
            names(nameCount) = nm
            nameCount = nameCount + 1
            If shown < 12 Then
                startPos = 0
                snippet = ""
                On Error Resume Next
                startPos = bm.Range.Start
                snippet = Host.LimitText(NormalizeSnippet(bm.Range.Text), 24, "...")
                Err.Clear
                line = "- " & nm & " @" & CStr(startPos) & " " & snippet
                If Len(sampleLines) = 0 Then
                    sampleLines = line
                Else
                    sampleLines = sampleLines & vbCrLf & line
                End If
                shown = shown + 1
            End If
        End If
    Next
    If Len(sampleLines) = 0 Then sampleLines = "（无用户书签）"
End Sub

Function TryGotoBookmark(doc, bookmarkName)
    On Error Resume Next
    TryGotoBookmark = False
    If doc.Bookmarks.Exists(bookmarkName) Then
        doc.Bookmarks(bookmarkName).Range.Select
        If Err.Number = 0 Then TryGotoBookmark = True
    End If
    Err.Clear
End Function

Function NormalizeSnippet(value)
    Dim text
    text = CStr(value)
    text = Replace(text, vbCr, " ")
    text = Replace(text, vbLf, " ")
    text = Replace(text, vbTab, " ")
    text = Replace(text, Chr(7), "")
    NormalizeSnippet = Trim(text)
End Function

Function SafeHostText(methodName)
    On Error Resume Next
    If methodName = "GetWordContextInfo" Then
        SafeHostText = Host.GetWordContextInfo()
    Else
        SafeHostText = ""
    End If
    If Err.Number <> 0 Then SafeHostText = ""
    Err.Clear
End Function

Function FailureJson(code, message)
    FailureJson = "{""ok"":false,""code"":""" & EscapeJson(code) & """,""message"":""" & EscapeJson(message) & """}"
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

Function SafePrompt(message, defaultValue)
    On Error Resume Next
    SafePrompt = Host.Prompt(message, defaultValue)
    If Err.Number <> 0 Then SafePrompt = defaultValue
    Err.Clear
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

Sub SafeRollbackWritePlan(planId)
    On Error Resume Next
    Host.RollbackWritePlan planId
    Err.Clear
End Sub

Sub SafeWriteLog(message)
    On Error Resume Next
    Host.WriteLog message
    Err.Clear
End Sub
