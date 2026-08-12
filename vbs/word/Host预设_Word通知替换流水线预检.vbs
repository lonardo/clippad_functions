' 函数名: HostWordNoticeReplacePipelinePreflight
' 描述: 通知/公文替换流水线：预检常见占位符与待替换词，支持最多3组查找替换预览，确认词“替换”后写回；并推荐清理包/页眉信息下一步
' 适用应用: Word
' 搜索范围: 全文
' 搜索对象: 无

Option Explicit

Const wdReplaceAll = 2
Const wdFindContinue = 1

Function Main()
    On Error Resume Next
    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = FailureJson("E_NO_WORD_APP", "未取得 Word 应用，请在 Word 中运行该预设")
        Exit Function
    End If
    Main = HostWordNoticeReplacePipelinePreflight(appObj)
End Function

Function HostWordNoticeReplacePipelinePreflight(appObj)
    On Error Resume Next

    Dim doc
    Set doc = appObj.ActiveDocument
    If Err.Number <> 0 Or TypeName(doc) = "Empty" Or TypeName(doc) = "Nothing" Then
        Err.Clear
        HostWordNoticeReplacePipelinePreflight = FailureJson("E_NO_DOCUMENT", "当前没有活动 Word 文档")
        Exit Function
    End If

    Dim bodyText, placeholderCount, samplePlaceholders
    bodyText = ""
    On Error Resume Next
    bodyText = doc.Content.Text
    Err.Clear
    placeholderCount = 0
    samplePlaceholders = ""
    CountPlaceholders bodyText, placeholderCount, samplePlaceholders

    Dim finds(3), repls(3), hits(3), pairCount, idx
    pairCount = 0
    For idx = 1 To 3
        If idx = 1 Then
            finds(idx) = SafePrompt("替换对" & CStr(idx) & " 查找文本（留空结束）", "XX")
        Else
            finds(idx) = SafePrompt("替换对" & CStr(idx) & " 查找文本（留空结束）", "")
        End If
        If Len(Trim(finds(idx))) = 0 Then Exit For
        repls(idx) = SafePrompt("替换对" & CStr(idx) & " 替换为", "")
        hits(idx) = CountFindHits(doc, finds(idx))
        pairCount = pairCount + 1
    Next

    ' VBScript has no IIf - fix below after string built

    Dim totalHits, pairSummary
    totalHits = 0
    pairSummary = ""
    For idx = 1 To pairCount
        totalHits = totalHits + hits(idx)
        If Len(pairSummary) > 0 Then pairSummary = pairSummary & "；"
        pairSummary = pairSummary & finds(idx) & "=>" & repls(idx) & " x" & CStr(hits(idx))
    Next

    Dim nextSteps, formatSummary, reviewSummary, previewText
    nextSteps = "下一步推荐：Host预设_Word清理包预检.vbs → Host预设_Word页眉文件信息预检.vbs → Host预设_Word书签导航摘要.vbs"
    formatSummary = SafeHostText("GetWordFormatIssueSummary")
    reviewSummary = SafeHostText("GetWordReviewSummary")
    previewText = "Word 通知替换流水线预检（尚未写入）" & vbCrLf & _
        "疑似占位符=" & CStr(placeholderCount) & "；样例=" & samplePlaceholders & vbCrLf & _
        "替换对数=" & CStr(pairCount) & "；预计命中=" & CStr(totalHits) & vbCrLf & _
        "替换对：" & pairSummary & vbCrLf & _
        "Host.GetWordFormatIssueSummary: " & formatSummary & vbCrLf & _
        "Host.GetWordReviewSummary: " & reviewSummary & vbCrLf & _
        nextSteps

    If pairCount = 0 Then
        Host.WriteClipboard previewText & vbCrLf & "未配置替换对，仅输出预检"
        SafeWriteLog previewText
        HostWordNoticeReplacePipelinePreflight = "{""ok"":true,""changed"":false,""message"":""" & EscapeJson(previewText) & """,""placeholderCount"":" & CStr(placeholderCount) & "}"
        Exit Function
    End If

    Dim planId, planPreview
    planId = SafeBeginWritePlan()
    SafeRecordWrite "word_notice_replace_pipeline", "{""pairs"":" & CStr(pairCount) & ",""hits"":" & CStr(totalHits) & "}", "office.word.pipeline.replace"
    planPreview = SafePreviewWritePlan()

    Dim confirmText
    confirmText = Trim(SafePrompt(previewText & vbCrLf & vbCrLf & "以上尚未写入。确认执行通知替换，请输入：替换", ""))
    If StrComp(confirmText, "替换", vbTextCompare) <> 0 Then
        SafeRollbackWritePlan planId
        HostWordNoticeReplacePipelinePreflight = FailureJson("E_CONFIRM_REQUIRED", "未输入“替换”，已取消且未修改文档")
        Exit Function
    End If

    Dim replacedTotal, replacedOne
    replacedTotal = 0
    For idx = 1 To pairCount
        replacedOne = ApplyReplace(doc, finds(idx), repls(idx))
        replacedTotal = replacedTotal + replacedOne
    Next
    SafeRollbackWritePlan planId

    Dim summary
    summary = "Word 通知替换流水线完成；替换对=" & CStr(pairCount) & "；替换次数≈" & CStr(replacedTotal) & vbCrLf & _
        pairSummary & vbCrLf & nextSteps & vbCrLf & "Preview: " & planPreview
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostWordNoticeReplacePipelinePreflight = "{""ok"":true,""changed"":true,""pairCount"":" & CStr(pairCount) & ",""replacedTotal"":" & CStr(replacedTotal) & ",""message"":""" & EscapeJson(summary) & """}"
End Function

Sub CountPlaceholders(bodyText, ByRef placeholderCount, ByRef samplePlaceholders)
    Dim patterns, p, idx, countOne, shown
    patterns = Array("{", "}", "【", "】", "XX", "xx", "___", "（待填）", "[姓名]", "[日期]", "XXX")
    placeholderCount = 0
    samplePlaceholders = ""
    shown = 0
    For idx = 0 To UBound(patterns)
        p = patterns(idx)
        countOne = CountToken(bodyText, p)
        If countOne > 0 Then
            placeholderCount = placeholderCount + countOne
            If shown < 8 Then
                If Len(samplePlaceholders) > 0 Then samplePlaceholders = samplePlaceholders & "; "
                samplePlaceholders = samplePlaceholders & p & "x" & CStr(countOne)
                shown = shown + 1
            End If
        End If
    Next
    If Len(samplePlaceholders) = 0 Then samplePlaceholders = "（未发现常见占位）"
End Sub

Function CountToken(text, token)
    Dim t, p, n, startPos
    t = CStr(text)
    p = CStr(token)
    n = 0
    startPos = 1
    If Len(p) = 0 Then
        CountToken = 0
        Exit Function
    End If
    Do
        startPos = InStr(startPos, t, p, vbTextCompare)
        If startPos = 0 Then Exit Do
        n = n + 1
        startPos = startPos + Len(p)
    Loop
    CountToken = n
End Function

Function CountFindHits(doc, findText)
    On Error Resume Next
    Dim rng, n
    n = 0
    Set rng = doc.Content.Duplicate
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
    End With
    Do While rng.Find.Execute
        n = n + 1
        If n > 5000 Then Exit Do
    Loop
    CountFindHits = n
    Err.Clear
End Function

Function ApplyReplace(doc, findText, replText)
    On Error Resume Next
    Dim rng
    Set rng = doc.Content
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Replacement.Text = replText
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .Execute findText, False, False, False, False, False, True, wdFindContinue, False, replText, wdReplaceAll
    End With
    ' Word doesn't return count reliably; approximate by pre-count style 0
    ApplyReplace = 0
    If Err.Number = 0 Then ApplyReplace = 1
    Err.Clear
End Function

Function SafeHostText(methodName)
    On Error Resume Next
    If methodName = "GetWordFormatIssueSummary" Then
        SafeHostText = Host.GetWordFormatIssueSummary()
    ElseIf methodName = "GetWordReviewSummary" Then
        SafeHostText = Host.GetWordReviewSummary()
    ElseIf methodName = "GetWordContextInfo" Then
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
