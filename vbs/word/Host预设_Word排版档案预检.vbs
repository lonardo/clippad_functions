' 函数名: HostWordFormatProfilePreflight
' 描述: 只读检查当前 Word 文档的结构、排版和审阅风险，为按样本套版或专业格式档案应用生成预检摘要；不会修改文档
' 适用应用: Word
' 搜索范围: 全文
' 搜索对象: 无

Option Explicit

Const wdNoProtection = -1

Function Main()
    On Error Resume Next

    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = "{""ok"":false,""code"":""E_NO_WORD_APP"",""message"":""未取得 Word 应用，请在 Word 中运行该预设""}"
        Exit Function
    End If

    Main = HostWordFormatProfilePreflight(appObj)
End Function

Function HostWordFormatProfilePreflight(appObj)
    On Error Resume Next

    Dim doc
    Set doc = appObj.ActiveDocument
    If Err.Number <> 0 Or TypeName(doc) = "Empty" Or TypeName(doc) = "Nothing" Then
        Err.Clear
        HostWordFormatProfilePreflight = "{""ok"":false,""code"":""E_NO_DOCUMENT"",""message"":""当前没有活动 Word 文档""}"
        Exit Function
    End If

    Dim formatProfile, wordContext, formatIssues, reviewSummary
    formatProfile = NormalizeProfile(SafePrompt("拟使用的格式档案（如：通用报告、用户确认的法律文书模板）；本次只做预检，不会修改文档", "通用报告"))
    wordContext = SafeHostText("GetWordContextInfo")
    formatIssues = SafeHostText("GetWordFormatIssueSummary")
    reviewSummary = SafeHostText("GetWordReviewSummary")

    Dim paragraphCount, tableCount, sectionCount, isProtected, hasRevisions, hasComments, savedDocument
    paragraphCount = SafeCount(doc.Paragraphs.Count)
    tableCount = SafeCount(doc.Tables.Count)
    sectionCount = SafeCount(doc.Sections.Count)
    isProtected = IsDocumentProtected(doc)
    hasRevisions = SummaryValueIsPositive(reviewSummary, "revisions")
    hasComments = SummaryValueIsPositive(reviewSummary, "comments")
    savedDocument = IsDocumentSaved(doc)

    Dim caution, nextStep
    caution = ""
    If isProtected Then caution = AppendCaution(caution, "文档受保护，不能直接套版")
    If hasRevisions Then caution = AppendCaution(caution, "存在修订，应用格式前需确认是否保留修订")
    If hasComments Then caution = AppendCaution(caution, "存在批注，格式应用不应删除批注")
    If Not savedDocument Then caution = AppendCaution(caution, "文档尚未保存，正式套版前应先保存或创建备份")
    If Len(caution) = 0 Then caution = "未发现必须阻止套版的基础风险；正式应用仍需先预览并确认"

    nextStep = "下一步请选择用户确认的样本文档或格式档案，仅映射标题、正文、表格、页眉页脚和页码等已勾选格式规则。"
    Dim summary
    summary = "Word 排版档案预检完成；格式档案=" & formatProfile & _
        "；段落=" & CStr(paragraphCount) & "；表格=" & CStr(tableCount) & _
        "；节=" & CStr(sectionCount) & "；只读=true" & vbCrLf & _
        "风险提示：" & caution & vbCrLf & _
        "Host.GetWordFormatIssueSummary: " & formatIssues & vbCrLf & _
        "Host.GetWordReviewSummary: " & reviewSummary
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostWordFormatProfilePreflight = "{""ok"":true,""readOnly"":true,""formatProfile"":""" & EscapeJson(formatProfile) & _
        """,""paragraphs"":" & CStr(paragraphCount) & ",""tables"":" & CStr(tableCount) & _
        ",""sections"":" & CStr(sectionCount) & ",""protected"":" & JsonBool(isProtected) & _
        ",""hasRevisions"":" & JsonBool(hasRevisions) & ",""hasComments"":" & JsonBool(hasComments) & _
        ",""saved"":" & JsonBool(savedDocument) & ",""requiresConfirmation"":true,""nextStep"":""" & EscapeJson(nextStep) & _
        """,""message"":""" & EscapeJson(summary) & """,""wordContext"":""" & EscapeJson(wordContext) & _
        """,""formatIssues"":""" & EscapeJson(formatIssues) & """,""reviewSummary"":""" & EscapeJson(reviewSummary) & """}"
End Function

Function NormalizeProfile(value)
    Dim text
    text = Trim(CStr(value))
    If Len(text) = 0 Then text = "通用报告"
    NormalizeProfile = Left(text, 80)
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
    If methodName = "GetWordContextInfo" Then
        SafeHostText = Host.GetWordContextInfo()
    ElseIf methodName = "GetWordFormatIssueSummary" Then
        SafeHostText = Host.GetWordFormatIssueSummary()
    ElseIf methodName = "GetWordReviewSummary" Then
        SafeHostText = Host.GetWordReviewSummary()
    Else
        SafeHostText = ""
    End If
    If Err.Number <> 0 Then
        SafeHostText = ""
        Err.Clear
    End If
End Function

Function SafeCount(value)
    On Error Resume Next
    SafeCount = CLng(value)
    If Err.Number <> 0 Then
        SafeCount = 0
        Err.Clear
    End If
End Function

Function IsDocumentProtected(doc)
    On Error Resume Next
    IsDocumentProtected = (CLng(doc.ProtectionType) <> wdNoProtection)
    If Err.Number <> 0 Then
        IsDocumentProtected = False
        Err.Clear
    End If
End Function

Function IsDocumentSaved(doc)
    On Error Resume Next
    IsDocumentSaved = (Len(Trim(CStr(doc.Path))) > 0)
    If Err.Number <> 0 Then
        IsDocumentSaved = False
        Err.Clear
    End If
End Function

Function SummaryValueIsPositive(summary, key)
    Dim marker, startAt, endAt, rawValue
    SummaryValueIsPositive = False
    marker = LCase(CStr(key)) & "="
    startAt = InStr(1, LCase(CStr(summary)), marker, vbTextCompare)
    If startAt <= 0 Then Exit Function
    startAt = startAt + Len(marker)
    endAt = InStr(startAt, summary, ";")
    If endAt <= 0 Then endAt = Len(summary) + 1
    rawValue = Trim(Mid(summary, startAt, endAt - startAt))
    If IsNumeric(rawValue) Then SummaryValueIsPositive = (CLng(rawValue) > 0)
End Function

Function AppendCaution(currentValue, newValue)
    If Len(currentValue) = 0 Then
        AppendCaution = newValue
    Else
        AppendCaution = currentValue & "；" & newValue
    End If
End Function

Function JsonBool(value)
    If value Then
        JsonBool = "true"
    Else
        JsonBool = "false"
    End If
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
