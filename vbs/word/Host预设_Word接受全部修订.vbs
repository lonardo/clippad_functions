' 函数名: HostWordAcceptAllRevisions
' 描述: 接受当前 Word 文档全部修订并删除批注；对标公文/审阅工具中的痕迹一键清零
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
    Main = HostWordAcceptAllRevisions(appObj)
End Function

Function HostWordAcceptAllRevisions(appObj)
    On Error Resume Next
    Dim doc
    Set doc = appObj.ActiveDocument
    If Err.Number <> 0 Or TypeName(doc) = "Empty" Or TypeName(doc) = "Nothing" Then
        Err.Clear
        HostWordAcceptAllRevisions = "{""ok"":false,""code"":""E_NO_DOCUMENT"",""message"":""当前没有活动 Word 文档""}"
        Exit Function
    End If

    Dim confirmation
    confirmation = SafePrompt("将接受全部修订并删除全部批注。请输入 确认 继续", "")
    If confirmation <> "确认" Then
        HostWordAcceptAllRevisions = "{""ok"":false,""code"":""E_CONFIRM_REQUIRED"",""message"":""已取消接受修订""}"
        Exit Function
    End If

    Dim reviewSummary, planId, previewJson, revBefore, cmtBefore
    reviewSummary = SafeHostText("GetWordReviewSummary")
    revBefore = doc.Revisions.Count
    cmtBefore = doc.Comments.Count
    planId = SafeBeginWritePlan()
    SafeRecordWrite "word_accept_all_revisions", "{""revisions"":" & CStr(revBefore) & ",""comments"":" & CStr(cmtBefore) & "}", "office.word.review.cleanup"
    previewJson = SafePreviewWritePlan()

    doc.AcceptAllRevisions
    Dim i
    For i = doc.Comments.Count To 1 Step -1
        doc.Comments(i).Delete
        Err.Clear
    Next
    SafeCloseWritePlan planId

    Dim summary
    summary = "Word 修订批注清零完成；原修订=" & CStr(revBefore) & "；原批注=" & CStr(cmtBefore) & _
        "；剩余修订=" & CStr(doc.Revisions.Count) & "；剩余批注=" & CStr(doc.Comments.Count) & vbCrLf & _
        "Host.GetWordReviewSummary: " & reviewSummary & vbCrLf & "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostWordAcceptAllRevisions = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""revisionsBefore"":" & CStr(revBefore) & _
        ",""commentsBefore"":" & CStr(cmtBefore) & ",""planId"":""" & EscapeJson(planId) & """}"
End Function

Function SafeHostText(methodName)
    On Error Resume Next
    If methodName = "GetWordReviewSummary" Then
        SafeHostText = Host.GetWordReviewSummary()
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