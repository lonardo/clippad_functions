' 函数名: HostWordFindReplaceWithPreflight
' 描述: 对当前 Word 文档执行查找替换前预检并记录写入计划；验证 Host.Prompt、GetRunPreflightPlan、写入计划和剪贴板摘要
' 适用应用: Word
' 搜索范围: 全文
' 搜索对象: 无

Option Explicit

Const wdFindContinue = 1
Const wdReplaceAll = 2

Function Main()
    On Error Resume Next

    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = "{""ok"":false,""code"":""E_NO_WORD_APP"",""message"":""未取得 Word 应用，请在 Word 中运行该预设""}"
        Exit Function
    End If

    Main = HostWordFindReplaceWithPreflight(appObj)
End Function

Function HostWordFindReplaceWithPreflight(appObj)
    On Error Resume Next

    Dim doc
    Set doc = appObj.ActiveDocument
    If Err.Number <> 0 Or TypeName(doc) = "Empty" Or TypeName(doc) = "Nothing" Then
        Err.Clear
        HostWordFindReplaceWithPreflight = "{""ok"":false,""code"":""E_NO_DOCUMENT"",""message"":""当前没有活动 Word 文档""}"
        Exit Function
    End If

    Dim findText, replaceText
    findText = SafePrompt("请输入要查找的文本", "")
    If Len(findText) = 0 Then
        HostWordFindReplaceWithPreflight = "{""ok"":false,""code"":""E_EMPTY_FIND"",""message"":""查找文本为空，已取消""}"
        Exit Function
    End If
    replaceText = SafePrompt("请输入替换后的文本", "")

    Dim preflightJson, planId, previewJson, validationJson
    preflightJson = Host.GetRunPreflightPlan("Word", True, False, False)
    planId = SafeBeginWritePlan()
    SafeRecordWrite "word_find_replace", "{""scope"":""activeDocument"",""find"":""" & EscapeJson(findText) & """,""replace"":""" & EscapeJson(replaceText) & """}", "office.word.findReplace"
    validationJson = SafeValidateWritePlan(planId)
    previewJson = SafePreviewWritePlan()

    Dim beforeCount, replaced
    beforeCount = CountOccurrences(doc.Content.Text, findText)
    replaced = ApplyFindReplace(doc, findText, replaceText)
    SafeCloseWritePlan planId

    Dim summary
    summary = "Word 查找替换完成；查找=" & findText & "；替换为=" & replaceText & "；估算命中=" & CStr(beforeCount) & vbCrLf & _
        "Host.GetRunPreflightPlan: " & preflightJson & vbCrLf & _
        "Preview: " & previewJson & vbCrLf & _
        "Validation: " & validationJson
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostWordFindReplaceWithPreflight = "{""ok"":" & JsonBool(replaced) & ",""message"":""" & EscapeJson(summary) & """,""estimatedMatches"":" & CStr(beforeCount) & _
        ",""planId"":""" & EscapeJson(planId) & """}"

End Function

Function ApplyFindReplace(doc, findText, replaceText)
    On Error Resume Next
    Dim rng, findObj
    Set rng = doc.Content
    Set findObj = rng.Find
    findObj.ClearFormatting
    findObj.Replacement.ClearFormatting
    findObj.Text = findText
    findObj.Replacement.Text = replaceText
    findObj.Forward = True
    findObj.Wrap = wdFindContinue
    findObj.Format = False
    findObj.MatchCase = False
    findObj.MatchWholeWord = False
    findObj.Execute findText, False, False, False, False, False, True, wdFindContinue, False, replaceText, wdReplaceAll
    ApplyFindReplace = (Err.Number = 0)
    Err.Clear
End Function

Function CountOccurrences(text, needle)
    Dim count, startAt, pos
    count = 0
    startAt = 1
    If Len(needle) = 0 Then
        CountOccurrences = 0
        Exit Function
    End If
    Do
        pos = InStr(startAt, text, needle, vbTextCompare)
        If pos <= 0 Then Exit Do
        count = count + 1
        startAt = pos + Len(needle)
    Loop
    CountOccurrences = count
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
