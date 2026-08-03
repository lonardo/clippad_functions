' 函数名: HostWordReviewCleanup
' 描述: 清理当前 Word 文档批注并按确认词接受修订；验证 Host.Prompt、Word 审阅上下文、写入计划诊断和剪贴板摘要
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

    Main = HostWordReviewCleanup(appObj)
End Function

Function HostWordReviewCleanup(appObj)
    On Error Resume Next

    Dim doc
    Set doc = appObj.ActiveDocument
    If Err.Number <> 0 Or TypeName(doc) = "Empty" Or TypeName(doc) = "Nothing" Then
        Err.Clear
        HostWordReviewCleanup = "{""ok"":false,""code"":""E_NO_DOCUMENT"",""message"":""当前没有活动 Word 文档""}"
        Exit Function
    End If

    Dim reviewBefore, confirmText
    reviewBefore = SafeHostText("GetWordReviewSummary")
    confirmText = Trim(SafePrompt("该操作会删除批注并接受所有修订。请输入 CONFIRM 继续。", ""))
    If UCase(confirmText) <> "CONFIRM" Then
        HostWordReviewCleanup = "{""ok"":false,""code"":""E_NOT_CONFIRMED"",""message"":""用户未确认清理审阅痕迹"",""reviewSummary"":""" & EscapeJson(reviewBefore) & """}"
        Exit Function
    End If

    Dim planId, previewJson, validationJson
    planId = SafeBeginWritePlan()
    SafeRecordWrite "word_review_cleanup", "{""scope"":""activeDocument"",""requiresConfirmation"":true}", "office.word.review.cleanup"
    validationJson = SafeValidateWritePlan(planId)
    previewJson = SafePreviewWritePlan()

    Dim revisionCount, commentCount
    revisionCount = 0
    commentCount = 0
    Err.Clear
    revisionCount = doc.Revisions.Count
    commentCount = doc.Comments.Count
    Err.Clear

    Dim i
    For i = doc.Comments.Count To 1 Step -1
        doc.Comments(i).Delete
    Next

    Err.Clear
    If doc.Revisions.Count > 0 Then doc.AcceptAllRevisions
    Err.Clear

    SafeCloseWritePlan planId

    Dim reviewAfter, summary
    reviewAfter = SafeHostText("GetWordReviewSummary")
    summary = "Word 审阅痕迹清理完成；原修订=" & CStr(revisionCount) & "，原批注=" & CStr(commentCount) & vbCrLf & _
        "Before: " & reviewBefore & vbCrLf & "After: " & reviewAfter
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostWordReviewCleanup = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""revisionCount"":" & CStr(revisionCount) & _
        ",""commentCount"":" & CStr(commentCount) & ",""planId"":""" & EscapeJson(planId) & _
        """,""preview"":""" & EscapeJson(previewJson) & """,""validation"":""" & EscapeJson(validationJson) & """}"

End Function

Function SafePrompt(message, defaultValue)
    On Error Resume Next
    SafePrompt = Host.Prompt(message, defaultValue)
    If Err.Number <> 0 Then
        SafePrompt = defaultValue
        Err.Clear
    End If
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

Function SafeValidateWritePlan(planId)
    On Error Resume Next
    SafeValidateWritePlan = Host.ValidateWritePlan(planId, False)
    If Err.Number <> 0 Then SafeValidateWritePlan = ""
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
