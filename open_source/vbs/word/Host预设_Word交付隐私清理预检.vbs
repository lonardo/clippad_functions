' 函数名: HostWordDeliveryPrivacyCleanupPreflight
' 描述: 交付前预检隐私痕迹后一键确认清理；默认安全包（批注+属性，不含修订），可改预览
' 适用应用: Word
' 搜索范围: 全文
' 搜索对象: 无
'
' 实现约定（供后续维护/生成使用）：
' - 先统计批注/修订/内置属性，默认动作为“预览”；只有明确输入“清理”才执行安全包：批注=是、属性=是、修订=否。
' - 单轮覆盖：预览 / 清理|修订 / 清理|批注 / 清理|属性 / 清理|全部 / 清理|安全。不自动另存副本。
' - 接受修订会改正文显示；属性清理基本不可回滚。RollbackWritePlan 只关闭 Host 计划记录。

Option Explicit

Const wdRDIComments = 1
Const wdRDIRevisions = 2
Const wdRDIVersions = 3
Const wdRDIRemovePersonalInformation = 4
Const wdRDIDocumentProperties = 8

Function Main()
    On Error Resume Next
    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = FailureJson("E_NO_WORD_APP", "未取得 Word 应用，请在 Word 中运行该预设")
        Exit Function
    End If
    Main = HostWordDeliveryPrivacyCleanupPreflight(appObj)
End Function

Function HostWordDeliveryPrivacyCleanupPreflight(appObj)
    On Error Resume Next

    Dim doc
    Set doc = appObj.ActiveDocument
    If Err.Number <> 0 Or TypeName(doc) = "Empty" Or TypeName(doc) = "Nothing" Then
        Err.Clear
        HostWordDeliveryPrivacyCleanupPreflight = FailureJson("E_NO_DOCUMENT", "当前没有活动 Word 文档")
        Exit Function
    End If

    Dim commentCount, revisionCount, author, lastAuthor, company, title, subject, keywords, remarks
    commentCount = 0
    revisionCount = 0
    author = ""
    lastAuthor = ""
    company = ""
    title = ""
    subject = ""
    keywords = ""
    remarks = ""
    CollectPrivacyStats doc, commentCount, revisionCount, author, lastAuthor, company, title, subject, keywords, remarks

    Dim doComments, doRevisions, doProps
    doComments = True
    doRevisions = False
    doProps = True

    Dim reviewSummary, wordSummary, previewText
    reviewSummary = SafeHostText("GetWordReviewSummary")
    wordSummary = SafeHostText("GetWordContextInfo")
    previewText = "Word 交付隐私清理场景包（尚未写入）" & vbCrLf & _
        "批注=" & CStr(commentCount) & "；修订=" & CStr(revisionCount) & vbCrLf & _
        "作者=" & author & "；最后编辑者=" & lastAuthor & "；公司=" & company & vbCrLf & _
        "标题=" & title & "；主题=" & subject & "；关键词=" & keywords & vbCrLf & _
        "备注=" & remarks & vbCrLf & _
        "默认安全包：批注=是；接受修订=否；属性/个人信息=是" & vbCrLf & _
        "风险：接受修订会改变正文显示；属性清理通常不可回滚；请先自行另存副本" & vbCrLf & _
        "Host.GetWordReviewSummary: " & reviewSummary & vbCrLf & _
        "Host.GetWordContextInfo: " & wordSummary & vbCrLf & vbCrLf & _
        "确认清理：请输入 清理；仅预览可直接确认或输入 预览。" & vbCrLf & _
        "仅预览：输入 预览" & vbCrLf & _
        "覆盖写法：清理|修订 / 清理|批注 / 清理|属性 / 清理|全部 / 清理|安全"

    Dim confirmRaw, doClean
    confirmRaw = SafePrompt(previewText, "预览")
    doClean = False
    If Not ParsePrivacyConfirm(confirmRaw, doClean, doComments, doRevisions, doProps) Then
        HostWordDeliveryPrivacyCleanupPreflight = FailureJson("E_CONFIRM_REQUIRED", "未确认清理/预览，已取消")
        Exit Function
    End If

    previewText = "Word 交付隐私清理场景包（尚未写入）" & vbCrLf & _
        "批注=" & CStr(commentCount) & "；修订=" & CStr(revisionCount) & vbCrLf & _
        "作者=" & author & "；最后编辑者=" & lastAuthor & "；公司=" & company & vbCrLf & _
        "将执行：批注=" & YesNoLabel(doComments) & "；接受修订=" & YesNoLabel(doRevisions) & "；属性/个人信息=" & YesNoLabel(doProps) & vbCrLf & _
        "风险：接受修订会改变正文显示；属性清理通常不可回滚"

    If Not doClean Then
        Host.WriteClipboard previewText
        SafeWriteLog previewText
        HostWordDeliveryPrivacyCleanupPreflight = "{""ok"":true,""readonly"":true,""commentCount"":" & CStr(commentCount) & _
            ",""revisionCount"":" & CStr(revisionCount) & ",""message"":""" & EscapeJson(previewText) & """}"
        Exit Function
    End If

    If Not (doComments Or doRevisions Or doProps) Then
        HostWordDeliveryPrivacyCleanupPreflight = FailureJson("E_NO_ACTION", "未选择任何清理项，已取消")
        Exit Function
    End If

    Dim planId, planPreview
    planId = SafeBeginWritePlan()
    SafeRecordWrite "word_delivery_privacy_scene", "{""comments"":" & LCase(CStr(doComments)) & ",""revisions"":" & LCase(CStr(doRevisions)) & ",""properties"":" & LCase(CStr(doProps)) & "}", "office.word.privacy.cleanup"
    planPreview = SafePreviewWritePlan()

    Dim clearedComments, acceptedRevisions, clearedProps
    clearedComments = 0
    acceptedRevisions = 0
    clearedProps = 0
    ApplyPrivacyCleanup doc, doComments, doRevisions, doProps, clearedComments, acceptedRevisions, clearedProps
    CloseWritePlan planId

    Dim summary
    summary = "Word 交付隐私清理场景包完成；删除批注=" & CStr(clearedComments) & _
        "；接受修订=" & CStr(acceptedRevisions) & "；属性清理标记=" & CStr(clearedProps) & _
        "；包=批注" & YesNoLabel(doComments) & "/修订" & YesNoLabel(doRevisions) & "/属性" & YesNoLabel(doProps) & _
        vbCrLf & "Preview: " & planPreview
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostWordDeliveryPrivacyCleanupPreflight = "{""ok"":true,""changed"":true,""clearedComments"":" & CStr(clearedComments) & _
        ",""acceptedRevisions"":" & CStr(acceptedRevisions) & ",""clearedProps"":" & CStr(clearedProps) & _
        ",""doComments"":" & LCase(CStr(doComments)) & ",""doRevisions"":" & LCase(CStr(doRevisions)) & ",""doProps"":" & LCase(CStr(doProps)) & _
        ",""writePlanPreview"":""" & EscapeJson(planPreview) & """,""message"":""" & EscapeJson(summary) & """}"
End Function

Function ParsePrivacyConfirm(rawText, ByRef doClean, ByRef doComments, ByRef doRevisions, ByRef doProps)
    Dim text, parts, head, i, token, sawOverride
    doClean = False
    ' defaults already set by caller to safe pack; keep unless override present
    text = Trim(CStr(rawText))
    If Len(text) = 0 Then
        ParsePrivacyConfirm = False
        Exit Function
    End If

    sawOverride = False
    If InStr(1, text, "|", vbBinaryCompare) > 0 Then
        parts = Split(text, "|")
        head = Trim(CStr(parts(0)))
    Else
        head = text
        parts = Split(text, "|")
    End If

    If StrComp(head, "预览", vbTextCompare) = 0 Or StrComp(head, "preview", vbTextCompare) = 0 Then
        doClean = False
        ParsePrivacyConfirm = True
        Exit Function
    End If

    If StrComp(head, "清理", vbTextCompare) = 0 Or StrComp(head, "clean", vbTextCompare) = 0 Or StrComp(head, "clear", vbTextCompare) = 0 Then
        doClean = True
        If UBound(parts) >= 1 Then
            ' reset then apply explicit pack tokens
            For i = 1 To UBound(parts)
                token = LCase(Trim(CStr(parts(i))))
                If Len(token) = 0 Then
                    ' skip
                ElseIf token = "安全" Or token = "safe" Then
                    doComments = True
                    doRevisions = False
                    doProps = True
                    sawOverride = True
                ElseIf token = "全部" Or token = "all" Then
                    doComments = True
                    doRevisions = True
                    doProps = True
                    sawOverride = True
                ElseIf token = "修订" Or token = "revisions" Or token = "revision" Then
                    If Not sawOverride Then
                        doComments = False
                        doRevisions = False
                        doProps = False
                        sawOverride = True
                    End If
                    doRevisions = True
                ElseIf token = "批注" Or token = "comments" Or token = "comment" Then
                    If Not sawOverride Then
                        doComments = False
                        doRevisions = False
                        doProps = False
                        sawOverride = True
                    End If
                    doComments = True
                ElseIf token = "属性" Or token = "props" Or token = "properties" Or token = "property" Then
                    If Not sawOverride Then
                        doComments = False
                        doRevisions = False
                        doProps = False
                        sawOverride = True
                    End If
                    doProps = True
                Else
                    ParsePrivacyConfirm = False
                    Exit Function
                End If
            Next
        End If
        ParsePrivacyConfirm = True
        Exit Function
    End If

    ParsePrivacyConfirm = False
End Function

Sub CollectPrivacyStats(doc, ByRef commentCount, ByRef revisionCount, ByRef author, ByRef lastAuthor, ByRef company, ByRef title, ByRef subject, ByRef keywords, ByRef remarks)
    On Error Resume Next
    commentCount = doc.Comments.Count
    If Err.Number <> 0 Then commentCount = 0
    Err.Clear
    revisionCount = doc.Revisions.Count
    If Err.Number <> 0 Then revisionCount = 0
    Err.Clear

    author = SafeDocProp(doc, "Author")
    lastAuthor = SafeDocProp(doc, "Last author")
    If lastAuthor = "(空)" Then lastAuthor = SafeDocProp(doc, "Last Author")
    company = SafeDocProp(doc, "Company")
    title = SafeDocProp(doc, "Title")
    subject = SafeDocProp(doc, "Subject")
    keywords = SafeDocProp(doc, "Keywords")
    remarks = SafeDocProp(doc, "Comments")
End Sub

Function SafeDocProp(doc, propName)
    On Error Resume Next
    Dim v
    v = ""
    v = CStr(doc.BuiltInDocumentProperties(propName).Value)
    If Err.Number <> 0 Then
        Err.Clear
        v = ""
    End If
    If Len(Trim(v)) = 0 Then v = "(空)"
    SafeDocProp = Left(v, 80)
End Function

Sub ApplyPrivacyCleanup(doc, doComments, doRevisions, doProps, ByRef clearedComments, ByRef acceptedRevisions, ByRef clearedProps)
    On Error Resume Next
    Dim i, beforeRev
    clearedComments = 0
    acceptedRevisions = 0
    clearedProps = 0

    If doComments Then
        clearedComments = doc.Comments.Count
        If Err.Number <> 0 Then clearedComments = 0
        Err.Clear
        For i = doc.Comments.Count To 1 Step -1
            doc.Comments(i).Delete
            Err.Clear
        Next
        doc.RemoveDocumentInformation wdRDIComments
        Err.Clear
    End If

    If doRevisions Then
        beforeRev = doc.Revisions.Count
        If Err.Number <> 0 Then beforeRev = 0
        Err.Clear
        If beforeRev > 0 Then
            doc.AcceptAllRevisions
            If Err.Number = 0 Then acceptedRevisions = beforeRev
            Err.Clear
        End If
        doc.RemoveDocumentInformation wdRDIRevisions
        Err.Clear
    End If

    If doProps Then
        ClearDocProp doc, "Author"
        ClearDocProp doc, "Last author"
        ClearDocProp doc, "Last Author"
        ClearDocProp doc, "Company"
        ClearDocProp doc, "Title"
        ClearDocProp doc, "Subject"
        ClearDocProp doc, "Keywords"
        ClearDocProp doc, "Comments"
        doc.RemovePersonalInformation = True
        Err.Clear
        doc.RemoveDocumentInformation wdRDIDocumentProperties
        If Err.Number = 0 Then clearedProps = clearedProps + 1
        Err.Clear
        doc.RemoveDocumentInformation wdRDIRemovePersonalInformation
        If Err.Number = 0 Then clearedProps = clearedProps + 1
        Err.Clear
        doc.RemoveDocumentInformation wdRDIVersions
        If Err.Number = 0 Then clearedProps = clearedProps + 1
        Err.Clear
    End If
End Sub

Sub ClearDocProp(doc, propName)
    On Error Resume Next
    doc.BuiltInDocumentProperties(propName).Value = ""
    Err.Clear
End Sub

Function SafeHostText(methodName)
    On Error Resume Next
    If methodName = "GetWordReviewSummary" Then
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

Sub CloseWritePlan(planId)
    On Error Resume Next
    Host.RollbackWritePlan planId
    Err.Clear
End Sub

Sub SafeWriteLog(message)
    On Error Resume Next
    Host.WriteLog message
    Err.Clear
End Sub

Function YesNoLabel(flag)
    If flag Then
        YesNoLabel = "是"
    Else
        YesNoLabel = "否"
    End If
End Function
