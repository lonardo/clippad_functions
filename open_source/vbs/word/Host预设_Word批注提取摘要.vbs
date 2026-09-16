' 函数名: HostWordCommentExtractSummary
' 描述: 只读提取当前文档批注作者、位置与摘录，可按作者过滤并复制摘要，不修改文档
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
    Main = HostWordCommentExtractSummary(appObj)
End Function

Function HostWordCommentExtractSummary(appObj)
    On Error Resume Next

    Dim doc
    Set doc = appObj.ActiveDocument
    If Err.Number <> 0 Or TypeName(doc) = "Empty" Or TypeName(doc) = "Nothing" Then
        Err.Clear
        HostWordCommentExtractSummary = FailureJson("E_NO_DOCUMENT", "当前没有活动 Word 文档")
        Exit Function
    End If

    Dim authorFilter, maxLinesText, maxLines
    authorFilter = Trim(SafePrompt("仅看该作者（留空=全部）", ""))
    maxLinesText = Trim(SafePrompt("摘要最多条数", "30"))
    If IsNumeric(maxLinesText) Then
        maxLines = CLng(maxLinesText)
    Else
        maxLines = 30
    End If
    If maxLines < 1 Then maxLines = 1
    If maxLines > 200 Then maxLines = 200

    Dim totalCount, matchedCount, sampleLines, authorsSample
    totalCount = 0
    matchedCount = 0
    sampleLines = ""
    authorsSample = ""
    CollectComments doc, authorFilter, maxLines, totalCount, matchedCount, sampleLines, authorsSample

    Dim wordSummary, reviewSummary, previewText
    wordSummary = SafeHostText("GetWordContextInfo")
    reviewSummary = SafeHostText("GetWordReviewSummary")
    previewText = "Word 批注提取摘要（只读，尚未写入/不修改文档）" & vbCrLf & _
        "批注总数=" & CStr(totalCount) & "；匹配=" & CStr(matchedCount) & _
        "；作者过滤=" & FilterLabel(authorFilter) & vbCrLf & _
        "作者样例=" & authorsSample & vbCrLf & _
        "条目：" & vbCrLf & sampleLines & vbCrLf & _
        "Host.GetWordContextInfo: " & wordSummary & vbCrLf & _
        "Host.GetWordReviewSummary: " & reviewSummary

    Host.WriteClipboard previewText
    SafeWriteLog previewText
    HostWordCommentExtractSummary = "{""ok"":true,""readonly"":true,""commentCount"":" & CStr(totalCount) & _
        ",""matchedCount"":" & CStr(matchedCount) & ",""message"":""" & EscapeJson(previewText) & """}"
End Function

Sub CollectComments(doc, authorFilter, maxLines, ByRef totalCount, ByRef matchedCount, ByRef sampleLines, ByRef authorsSample)
    On Error Resume Next
    Dim cmt, author, body, scopeText, pageNo, line, shown
    Dim authorKeys(), authorCount, a
    totalCount = 0
    matchedCount = 0
    shown = 0
    sampleLines = "（无批注）"
    authorsSample = "（无）"
    authorCount = 0

    For Each cmt In doc.Comments
        totalCount = totalCount + 1
        author = ""
        body = ""
        scopeText = ""
        pageNo = "?"
        author = Trim(CStr(cmt.Author))
        body = NormalizeCommentText(cmt.Range.Text)
        scopeText = NormalizeCommentText(cmt.Scope.Text)
        On Error Resume Next
        pageNo = CStr(cmt.Scope.Information(3))
        If Err.Number <> 0 Then
            Err.Clear
            pageNo = "?"
        End If

        If Len(author) > 0 Then
            If Not AuthorSeen(authorKeys, authorCount, author) Then
                ReDim Preserve authorKeys(authorCount)
                authorKeys(authorCount) = author
                authorCount = authorCount + 1
            End If
        End If

        If Len(authorFilter) = 0 Or InStr(1, author, authorFilter, vbTextCompare) > 0 Then
            matchedCount = matchedCount + 1
            If shown < maxLines Then
                line = CStr(matchedCount) & ". [" & author & "] p" & pageNo & " " & _
                    Host.LimitText(body, 40, "...") & " | 原文:" & Host.LimitText(scopeText, 24, "...")
                If shown = 0 Then
                    sampleLines = line
                Else
                    sampleLines = sampleLines & vbCrLf & line
                End If
                shown = shown + 1
            End If
        End If
    Next

    If authorCount > 0 Then
        authorsSample = ""
        For a = 0 To authorCount - 1
            If a > 0 Then authorsSample = authorsSample & "; "
            authorsSample = authorsSample & authorKeys(a)
            If a >= 7 Then
                authorsSample = authorsSample & "; ..."
                Exit For
            End If
        Next
    End If

    If matchedCount = 0 And totalCount > 0 Then
        sampleLines = "（无匹配批注）"
    End If
End Sub

Function AuthorSeen(ByRef authorKeys, authorCount, author)
    Dim i
    AuthorSeen = False
    For i = 0 To authorCount - 1
        If StrComp(authorKeys(i), author, vbTextCompare) = 0 Then
            AuthorSeen = True
            Exit Function
        End If
    Next
End Function

Function NormalizeCommentText(value)
    Dim text
    text = CStr(value)
    text = Replace(text, vbCr, " ")
    text = Replace(text, vbLf, " ")
    text = Replace(text, vbTab, " ")
    text = Replace(text, Chr(7), "")
    Do While InStr(text, "  ") > 0
        text = Replace(text, "  ", " ")
    Loop
    NormalizeCommentText = Trim(text)
End Function

Function FilterLabel(authorFilter)
    If Len(Trim(CStr(authorFilter))) = 0 Then
        FilterLabel = "全部"
    Else
        FilterLabel = authorFilter
    End If
End Function

Function SafeHostText(methodName)
    On Error Resume Next
    If methodName = "GetWordContextInfo" Then
        SafeHostText = Host.GetWordContextInfo()
    ElseIf methodName = "GetWordReviewSummary" Then
        SafeHostText = Host.GetWordReviewSummary()
    ElseIf methodName = "GetWordFormatIssueSummary" Then
        SafeHostText = Host.GetWordFormatIssueSummary()
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
