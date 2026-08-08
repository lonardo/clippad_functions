' 函数名: HostWordTocCaptionPreflight
' 描述: 预检 Word 目录/题注目录后一键确认更新；缺目录时可创建基础 TOC，可改预览
' 适用应用: Word
' 搜索范围: 全文
' 搜索对象: 无
'
' 实现约定（供后续维护/生成使用）：
' - 先统计 TOC/TOF/题注域与标题样例，默认动作为“预览”；只有明确输入“更新”才执行写入。
' - 默认智能范围：有 TOC/TOF 则更新已有项；都没有时在文首创建基础目录；仅在有题注域或显式 tof/创建 时才创建题注目录。
' - 不批量自动给图/表插题注。写入计划仅展示意图；RollbackWritePlan 不能撤销 Word COM 更新。

Option Explicit

Const wdCollapseEnd = 0
Const wdStyleHeading1 = -2

Function Main()
    On Error Resume Next
    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = FailureJson("E_NO_WORD_APP", "未取得 Word 应用，请在 Word 中运行该预设")
        Exit Function
    End If
    Main = HostWordTocCaptionPreflight(appObj)
End Function

Function HostWordTocCaptionPreflight(appObj)
    On Error Resume Next

    Dim doc
    Set doc = appObj.ActiveDocument
    If Err.Number <> 0 Or TypeName(doc) = "Empty" Or TypeName(doc) = "Nothing" Then
        Err.Clear
        HostWordTocCaptionPreflight = FailureJson("E_NO_DOCUMENT", "当前没有活动 Word 文档")
        Exit Function
    End If

    Dim tocCount, tofCount, captionFieldCount, headingSample, tocSample
    tocCount = 0
    tofCount = 0
    captionFieldCount = 0
    headingSample = ""
    tocSample = ""
    CollectTocCaptionStats doc, tocCount, tofCount, captionFieldCount, headingSample, tocSample

    Dim defaultScope, defaultLabel
    defaultScope = InferDefaultTocScope(tocCount, tofCount, captionFieldCount)
    defaultLabel = TocScopeLabel(defaultScope)

    Dim wordSummary, formatSummary, previewText
    wordSummary = SafeHostText("GetWordContextInfo")
    formatSummary = SafeHostText("GetWordFormatIssueSummary")
    previewText = "Word 目录题注场景包（尚未写入）" & vbCrLf & _
        "目录=" & CStr(tocCount) & "；图表/题注目录=" & CStr(tofCount) & "；题注相关域=" & CStr(captionFieldCount) & vbCrLf & _
        "默认动作=预览（智能写入范围：" & defaultLabel & "）" & vbCrLf & _
        "目录样例=" & tocSample & vbCrLf & _
        "标题样例=" & headingSample & vbCrLf & _
        "说明：会刷新已有目录页码；缺目录时可创建基础 TOC；不自动批量插题注" & vbCrLf & _
        "风险：页码/目录域显示会变；RollbackWritePlan 不能撤销 Word COM" & vbCrLf & _
        "Host.GetWordContextInfo: " & wordSummary & vbCrLf & _
        "Host.GetWordFormatIssueSummary: " & formatSummary & vbCrLf & vbCrLf & _
        "确认更新：请输入 更新；仅预览可直接确认或输入 预览。" & vbCrLf & _
        "仅预览：输入 预览" & vbCrLf & _
        "覆盖写法：更新|toc / 更新|tof / 更新|toc_tof / 更新|创建"

    Dim confirmRaw, doUpdate, scopeCode
    confirmRaw = SafePrompt(previewText, "预览")
    doUpdate = False
    scopeCode = defaultScope
    If Not ParseTocConfirm(confirmRaw, doUpdate, scopeCode, defaultScope) Then
        HostWordTocCaptionPreflight = FailureJson("E_CONFIRM_REQUIRED", "未确认更新/预览，已取消")
        Exit Function
    End If

    If Not doUpdate Then
        Host.WriteClipboard previewText
        SafeWriteLog previewText
        HostWordTocCaptionPreflight = "{""ok"":true,""readonly"":true,""tocCount"":" & CStr(tocCount) & _
            ",""tofCount"":" & CStr(tofCount) & ",""captionFields"":" & CStr(captionFieldCount) & _
            ",""defaultScope"":""" & EscapeJson(defaultScope) & """,""message"":""" & EscapeJson(previewText) & """}"
        Exit Function
    End If

    Dim planId, planPreview
    planId = SafeBeginWritePlan()
    SafeRecordWrite "word_toc_caption_scene", "{""scope"":""" & EscapeJson(scopeCode) & """,""toc"":" & CStr(tocCount) & ",""tof"":" & CStr(tofCount) & ",""captions"":" & CStr(captionFieldCount) & "}", "office.word.toc"
    planPreview = SafePreviewWritePlan()

    Dim createdToc, createdTof, updatedToc, updatedTof, noteText
    createdToc = 0
    createdTof = 0
    updatedToc = 0
    updatedTof = 0
    noteText = ""

    If scopeCode = "create" Or ((scopeCode = "toc" Or scopeCode = "toc_tof") And tocCount = 0) Then
        If CreateBasicToc(doc) Then
            createdToc = 1
            tocCount = tocCount + 1
        Else
            noteText = AppendNote(noteText, "基础目录创建失败或跳过")
        End If
    End If

    If scopeCode = "create" Or scopeCode = "tof" Or scopeCode = "toc_tof" Then
        If tofCount = 0 Then
            If captionFieldCount > 0 Or scopeCode = "create" Or scopeCode = "tof" Then
                If CreateBasicTof(doc) Then
                    createdTof = 1
                    tofCount = tofCount + 1
                Else
                    noteText = AppendNote(noteText, "题注目录创建失败或跳过")
                End If
            Else
                noteText = AppendNote(noteText, "无题注域，未自动创建题注目录")
            End If
        End If
    End If

    If scopeCode = "toc" Or scopeCode = "toc_tof" Or scopeCode = "create" Then
        updatedToc = UpdateTablesOfContents(doc)
    End If
    If scopeCode = "tof" Or scopeCode = "toc_tof" Or scopeCode = "create" Then
        updatedTof = UpdateTablesOfFigures(doc)
    End If

    If createdToc = 0 And createdTof = 0 And updatedToc = 0 And updatedTof = 0 Then
        CloseWritePlan planId
        HostWordTocCaptionPreflight = FailureJson("E_NO_TOC_ACTION", "没有可更新/可创建的目录项" & IIf(Len(noteText) > 0, "：" & noteText, ""))
        Exit Function
    End If

    CloseWritePlan planId

    Dim summary, noteShow
    If Len(noteText) = 0 Then
        noteShow = "无"
    Else
        noteShow = noteText
    End If
    summary = "Word 目录题注场景包完成；范围=" & TocScopeLabel(scopeCode) & _
        "；新建目录=" & CStr(createdToc) & "；新建题注目录=" & CStr(createdTof) & _
        "；更新目录=" & CStr(updatedToc) & "；更新题注目录=" & CStr(updatedTof) & _
        "；未批量插题注" & vbCrLf & "说明=" & noteShow & vbCrLf & "Preview: " & planPreview
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostWordTocCaptionPreflight = "{""ok"":true,""changed"":true,""scope"":""" & EscapeJson(scopeCode) & """,""createdToc"":" & CStr(createdToc) & _
        ",""createdTof"":" & CStr(createdTof) & ",""updatedToc"":" & CStr(updatedToc) & ",""updatedTof"":" & CStr(updatedTof) & _
        ",""writePlanPreview"":""" & EscapeJson(planPreview) & """,""message"":""" & EscapeJson(summary) & """}"
End Function

Function InferDefaultTocScope(tocCount, tofCount, captionFieldCount)
    If tocCount > 0 And tofCount > 0 Then
        InferDefaultTocScope = "toc_tof"
    ElseIf tocCount > 0 And tofCount = 0 Then
        If captionFieldCount > 0 Then
            InferDefaultTocScope = "toc_tof"
        Else
            InferDefaultTocScope = "toc"
        End If
    ElseIf tocCount = 0 And tofCount > 0 Then
        InferDefaultTocScope = "tof"
    Else
        InferDefaultTocScope = "create"
    End If
End Function

Function ParseTocConfirm(rawText, ByRef doUpdate, ByRef scopeCode, defaultScope)
    Dim text, parts, head, overrideText
    doUpdate = False
    scopeCode = defaultScope
    text = Trim(CStr(rawText))
    If Len(text) = 0 Then
        ParseTocConfirm = False
        Exit Function
    End If

    overrideText = ""
    If InStr(1, text, "|", vbBinaryCompare) > 0 Then
        parts = Split(text, "|")
        head = Trim(CStr(parts(0)))
        If UBound(parts) >= 1 Then overrideText = Trim(CStr(parts(1)))
    Else
        head = text
    End If

    If StrComp(head, "预览", vbTextCompare) = 0 Or StrComp(head, "preview", vbTextCompare) = 0 Then
        doUpdate = False
        ParseTocConfirm = True
        Exit Function
    End If

    If StrComp(head, "更新", vbTextCompare) = 0 Or StrComp(head, "update", vbTextCompare) = 0 Then
        doUpdate = True
        If Len(overrideText) > 0 Then
            scopeCode = NormalizeTocScope(overrideText)
            If Len(scopeCode) = 0 Then
                ParseTocConfirm = False
                Exit Function
            End If
        Else
            scopeCode = defaultScope
        End If
        ParseTocConfirm = True
        Exit Function
    End If

    scopeCode = NormalizeTocScope(head)
    If Len(scopeCode) > 0 Then
        doUpdate = True
        ParseTocConfirm = True
    Else
        ParseTocConfirm = False
    End If
End Function

Function NormalizeTocScope(text)
    Dim t
    t = LCase(Trim(CStr(text)))
    If t = "toc" Or t = "目录" Or t = "更新目录" Then
        NormalizeTocScope = "toc"
    ElseIf t = "tof" Or t = "题注目录" Or t = "图表目录" Then
        NormalizeTocScope = "tof"
    ElseIf t = "toc_tof" Or t = "toc+tof" Or t = "all" Or t = "目录+题注目录" Or t = "更新目录+题注目录" Then
        NormalizeTocScope = "toc_tof"
    ElseIf t = "create" Or t = "创建" Or t = "新建" Then
        NormalizeTocScope = "create"
    Else
        NormalizeTocScope = ""
    End If
End Function

Function TocScopeLabel(code)
    If code = "toc" Then
        TocScopeLabel = "仅目录"
    ElseIf code = "tof" Then
        TocScopeLabel = "仅题注目录"
    ElseIf code = "toc_tof" Then
        TocScopeLabel = "目录+题注目录"
    ElseIf code = "create" Then
        TocScopeLabel = "缺则创建并更新"
    Else
        TocScopeLabel = CStr(code)
    End If
End Function

Function CreateBasicToc(doc)
    On Error Resume Next
    Dim rng, toc
    CreateBasicToc = False
    Set rng = doc.Range(0, 0)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    rng.InsertBefore "目录" & vbCr
    Err.Clear
    Set rng = doc.Paragraphs(1).Range
    rng.Style = wdStyleHeading1
    Err.Clear
    Set rng = doc.Paragraphs(1).Range
    rng.Collapse wdCollapseEnd
    Set toc = doc.TablesOfContents.Add(rng, True, 1, 3)
    If Err.Number = 0 Then
        CreateBasicToc = True
    End If
    Err.Clear
End Function

Function CreateBasicTof(doc)
    On Error Resume Next
    Dim rng, tof
    CreateBasicTof = False
    If doc.TablesOfContents.Count > 0 Then
        Set rng = doc.TablesOfContents(1).Range
        rng.Collapse wdCollapseEnd
    Else
        Set rng = doc.Range(0, 0)
    End If
    If Err.Number <> 0 Then
        Err.Clear
        Set rng = doc.Range(0, 0)
    End If
    rng.InsertAfter vbCr & "图表目录" & vbCr
    Err.Clear
    Set rng = doc.Range(rng.End, rng.End)
    Set tof = doc.TablesOfFigures.Add(rng)
    If Err.Number = 0 Then
        CreateBasicTof = True
    End If
    Err.Clear
End Function

Sub CollectTocCaptionStats(doc, ByRef tocCount, ByRef tofCount, ByRef captionFieldCount, ByRef headingSample, ByRef tocSample)
    On Error Resume Next
    Dim toc, fld, i, txt, styleName, para
    tocCount = 0
    tofCount = 0
    captionFieldCount = 0
    headingSample = ""
    tocSample = ""

    tocCount = doc.TablesOfContents.Count
    If Err.Number <> 0 Then tocCount = 0
    Err.Clear
    tofCount = doc.TablesOfFigures.Count
    If Err.Number <> 0 Then tofCount = 0
    Err.Clear

    i = 0
    For Each toc In doc.TablesOfContents
        txt = ""
        txt = Left(Replace(Replace(CStr(toc.Range.Text), vbCr, " "), vbLf, " "), 40)
        If Len(txt) > 0 Then
            If Len(tocSample) > 0 Then tocSample = tocSample & " | "
            tocSample = tocSample & txt
            i = i + 1
            If i >= 3 Then Exit For
        End If
        Err.Clear
    Next
    If Len(tocSample) = 0 Then tocSample = "无"

    For Each fld In doc.Fields
        txt = LCase(CStr(fld.Code.Text))
        If InStr(1, txt, "seq ", vbTextCompare) > 0 Or InStr(1, txt, "styleref", vbTextCompare) > 0 Or InStr(1, txt, "toc \\c", vbTextCompare) > 0 Then
            If InStr(1, txt, "图", vbTextCompare) > 0 Or InStr(1, txt, "表", vbTextCompare) > 0 Or InStr(1, txt, "figure", vbTextCompare) > 0 Or InStr(1, txt, "table", vbTextCompare) > 0 Then
                captionFieldCount = captionFieldCount + 1
            End If
        End If
        Err.Clear
    Next

    i = 0
    For Each para In doc.Paragraphs
        styleName = ""
        styleName = CStr(para.Style.NameLocal)
        If Err.Number <> 0 Then
            Err.Clear
            styleName = ""
        End If
        If InStr(1, styleName, "标题 1", vbTextCompare) > 0 Or InStr(1, styleName, "Heading 1", vbTextCompare) > 0 Then
            txt = Trim(Replace(Replace(CStr(para.Range.Text), vbCr, ""), vbLf, ""))
            If Len(txt) > 0 Then
                If Len(headingSample) > 0 Then headingSample = headingSample & " | "
                headingSample = headingSample & Left(txt, 30)
                i = i + 1
                If i >= 5 Then Exit For
            End If
        End If
        Err.Clear
    Next
    If Len(headingSample) = 0 Then headingSample = "无"
End Sub

Function UpdateTablesOfContents(doc)
    On Error Resume Next
    Dim toc, n
    n = 0
    For Each toc In doc.TablesOfContents
        toc.Update
        If Err.Number = 0 Then n = n + 1
        Err.Clear
    Next
    UpdateTablesOfContents = n
End Function

Function UpdateTablesOfFigures(doc)
    On Error Resume Next
    Dim tof, n
    n = 0
    For Each tof In doc.TablesOfFigures
        tof.Update
        If Err.Number = 0 Then n = n + 1
        Err.Clear
    Next
    UpdateTablesOfFigures = n
End Function

Function AppendNote(baseText, addText)
    If Len(Trim(CStr(baseText))) = 0 Then
        AppendNote = CStr(addText)
    Else
        AppendNote = CStr(baseText) & "；" & CStr(addText)
    End If
End Function

Function IIf(cond, a, b)
    If cond Then
        IIf = a
    Else
        IIf = b
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
