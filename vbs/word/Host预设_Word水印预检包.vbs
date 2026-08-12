' 函数名: HostWordWatermarkPreflightPack
' 描述: 预检文档已有水印/页眉艺术字痕迹，可添加草稿/内部/保密文字水印或清除；确认后写入，默认不改正文
' 适用应用: Word
' 搜索范围: 全文
' 搜索对象: 无

Option Explicit

Const wdHeaderFooterPrimary = 1
Const msoTrue = -1
Const msoFalse = 0
Const msoTextOrientationHorizontal = 1
Const msoSendBehindText = 1
Const wdColorGray25 = 12632256

Function Main()
    On Error Resume Next
    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = FailureJson("E_NO_WORD_APP", "未取得 Word 应用，请在 Word 中运行该预设")
        Exit Function
    End If
    Main = HostWordWatermarkPreflightPack(appObj)
End Function

Function HostWordWatermarkPreflightPack(appObj)
    On Error Resume Next

    Dim doc
    Set doc = appObj.ActiveDocument
    If Err.Number <> 0 Or TypeName(doc) = "Empty" Or TypeName(doc) = "Nothing" Then
        Err.Clear
        HostWordWatermarkPreflightPack = FailureJson("E_NO_DOCUMENT", "当前没有活动 Word 文档")
        Exit Function
    End If

    Dim actionText, actionCode, markText
    actionText = Trim(SafePrompt("动作：预览 / 添加 / 清除", "预览"))
    actionCode = NormalizeWmAction(actionText)
    If Len(actionCode) = 0 Then
        HostWordWatermarkPreflightPack = FailureJson("E_ACTION_INVALID", "动作仅支持：预览 / 添加 / 清除")
        Exit Function
    End If

    markText = "草稿"
    If actionCode = "add" Then
        markText = Trim(SafePrompt("水印文字：草稿 / 内部 / 保密（也可自定义）", "草稿"))
        If Len(markText) = 0 Then markText = "草稿"
        markText = NormalizeWmLabel(markText)
    End If

    Dim sectionCount, wmCount, sampleText
    sectionCount = 0
    wmCount = 0
    sampleText = ""
    CollectWatermarkStats doc, sectionCount, wmCount, sampleText

    Dim wordSummary, reviewSummary, previewText, extra
    wordSummary = SafeHostText("GetWordContextInfo")
    reviewSummary = SafeHostText("GetWordReviewSummary")
    extra = ""
    If actionCode = "add" Then extra = "；将添加文字=" & markText
    previewText = "Word 水印预检包（尚未写入）" & vbCrLf & _
        "节数=" & CStr(sectionCount) & "；疑似水印形状=" & CStr(wmCount) & vbCrLf & _
        "动作=" & ActionLabel(actionCode) & extra & vbCrLf & _
        "样例=" & sampleText & vbCrLf & _
        "默认只改页眉水印层，不改正文" & vbCrLf & _
        "Host.GetWordContextInfo: " & wordSummary & vbCrLf & _
        "Host.GetWordReviewSummary: " & reviewSummary

    If actionCode = "preview" Then
        Host.WriteClipboard previewText
        SafeWriteLog previewText
        HostWordWatermarkPreflightPack = "{""ok"":true,""readonly"":true,""watermarkCount"":" & CStr(wmCount) & ",""sections"":" & CStr(sectionCount) & ",""message"":""" & EscapeJson(previewText) & """}"
        Exit Function
    End If

    Dim planId, planPreview, confirmWord, confirmText
    If actionCode = "add" Then
        confirmWord = "写入"
    Else
        confirmWord = "清理"
    End If

    planId = SafeBeginWritePlan()
    SafeRecordWrite "word_watermark_preflight_pack", "{""action"":""" & EscapeJson(actionCode) & """,""text"":""" & EscapeJson(markText) & """,""existing"":" & CStr(wmCount) & "}", "office.word.watermark"
    planPreview = SafePreviewWritePlan()

    confirmText = Trim(SafePrompt(previewText & vbCrLf & vbCrLf & "以上尚未写入。确认执行请输入：" & confirmWord, ""))
    If StrComp(confirmText, confirmWord, vbTextCompare) <> 0 Then
        SafeRollbackWritePlan planId
        HostWordWatermarkPreflightPack = FailureJson("E_CONFIRM_REQUIRED", "未输入“" & confirmWord & "”，已取消且未修改文档")
        Exit Function
    End If

    Dim changed
    changed = 0
    If actionCode = "clear" Then
        changed = ClearWatermarks(doc)
    Else
        changed = AddTextWatermark(doc, markText)
    End If
    SafeRollbackWritePlan planId

    Dim summary
    summary = "Word 水印预检包完成；动作=" & ActionLabel(actionCode) & "；影响区域/形状=" & CStr(changed) & _
        "；正文未改" & vbCrLf & "Preview: " & planPreview
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostWordWatermarkPreflightPack = "{""ok"":true,""changed"":true,""action"":""" & EscapeJson(actionCode) & """,""affected"":" & CStr(changed) & ",""message"":""" & EscapeJson(summary) & """}"
End Function

Function NormalizeWmAction(text)
    Dim t
    t = LCase(Trim(CStr(text)))
    If t = "预览" Or t = "preview" Or t = "view" Then
        NormalizeWmAction = "preview"
    ElseIf t = "添加" Or t = "add" Or t = "写入" Then
        NormalizeWmAction = "add"
    ElseIf t = "清除" Or t = "清理" Or t = "clear" Or t = "remove" Then
        NormalizeWmAction = "clear"
    Else
        NormalizeWmAction = ""
    End If
End Function

Function NormalizeWmLabel(text)
    Dim t
    t = Trim(CStr(text))
    If StrComp(t, "内部", vbTextCompare) = 0 Or StrComp(t, "internal", vbTextCompare) = 0 Then
        NormalizeWmLabel = "内部"
    ElseIf StrComp(t, "保密", vbTextCompare) = 0 Or StrComp(t, "secret", vbTextCompare) = 0 Or StrComp(t, "confidential", vbTextCompare) = 0 Then
        NormalizeWmLabel = "保密"
    ElseIf StrComp(t, "草稿", vbTextCompare) = 0 Or StrComp(t, "draft", vbTextCompare) = 0 Then
        NormalizeWmLabel = "草稿"
    Else
        NormalizeWmLabel = t
    End If
End Function

Function ActionLabel(code)
    If code = "add" Then
        ActionLabel = "添加水印"
    ElseIf code = "clear" Then
        ActionLabel = "清除水印"
    Else
        ActionLabel = "仅预览"
    End If
End Function

Sub CollectWatermarkStats(doc, ByRef sectionCount, ByRef wmCount, ByRef sampleText)
    On Error Resume Next
    Dim section, shape, nm, tx, hit
    sectionCount = 0
    wmCount = 0
    sampleText = "（无）"
    For Each section In doc.Sections
        sectionCount = sectionCount + 1
        For Each shape In section.Headers(wdHeaderFooterPrimary).Shapes
            hit = False
            nm = ""
            tx = ""
            nm = CStr(shape.Name)
            tx = ShapeText(shape)
            If InStr(1, nm, "Water", vbTextCompare) > 0 Then hit = True
            If InStr(1, nm, "水印", vbTextCompare) > 0 Then hit = True
            If InStr(1, nm, "OA_WM_", vbTextCompare) > 0 Then hit = True
            If InStr(1, tx, "草稿", vbTextCompare) > 0 Then hit = True
            If InStr(1, tx, "内部", vbTextCompare) > 0 Then hit = True
            If InStr(1, tx, "保密", vbTextCompare) > 0 Then hit = True
            If InStr(1, tx, "DRAFT", vbTextCompare) > 0 Then hit = True
            If shape.Type = 17 Then
                If Len(Trim(tx)) > 0 Then hit = True
            End If
            If hit Then
                wmCount = wmCount + 1
                If sampleText = "（无）" Then
                    sampleText = nm & "/" & Host.LimitText(tx, 20, "...")
                End If
            End If
        Next
    Next
End Sub

Function ShapeText(shape)
    On Error Resume Next
    ShapeText = ""
    ShapeText = shape.TextFrame.TextRange.Text
    If Err.Number <> 0 Then
        Err.Clear
        ShapeText = ""
        ShapeText = shape.TextEffect.Text
        If Err.Number <> 0 Then
            Err.Clear
            ShapeText = ""
        End If
    End If
    ShapeText = Replace(CStr(ShapeText), vbCr, "")
    ShapeText = Replace(ShapeText, vbLf, "")
    ShapeText = Trim(ShapeText)
End Function

Function ClearWatermarks(doc)
    On Error Resume Next
    Dim section, shape, i, removed, nm, tx, hit
    removed = 0
    For Each section In doc.Sections
        For i = section.Headers(wdHeaderFooterPrimary).Shapes.Count To 1 Step -1
            Set shape = section.Headers(wdHeaderFooterPrimary).Shapes(i)
            hit = False
            nm = CStr(shape.Name)
            tx = ShapeText(shape)
            If InStr(1, nm, "Water", vbTextCompare) > 0 Then hit = True
            If InStr(1, nm, "水印", vbTextCompare) > 0 Then hit = True
            If InStr(1, nm, "OA_WM_", vbTextCompare) > 0 Then hit = True
            If InStr(1, tx, "草稿", vbTextCompare) > 0 Then hit = True
            If InStr(1, tx, "内部", vbTextCompare) > 0 Then hit = True
            If InStr(1, tx, "保密", vbTextCompare) > 0 Then hit = True
            If InStr(1, tx, "DRAFT", vbTextCompare) > 0 Then hit = True
            If hit Then
                shape.Delete
                If Err.Number = 0 Then removed = removed + 1
                Err.Clear
            End If
        Next
    Next
    ClearWatermarks = removed
End Function

Function AddTextWatermark(doc, markText)
    On Error Resume Next
    Dim section, shape, added, width, height
    added = 0
    ClearOwnWatermarks doc
    For Each section In doc.Sections
        width = 400
        height = 100
        On Error Resume Next
        If doc.PageSetup.PageWidth > 0 Then width = doc.PageSetup.PageWidth - doc.PageSetup.LeftMargin - doc.PageSetup.RightMargin
        Set shape = section.Headers(wdHeaderFooterPrimary).Shapes.AddTextbox(msoTextOrientationHorizontal, 80, 200, width, height)
        If Err.Number = 0 Then
            If Not shape Is Nothing Then
                shape.Name = "OA_WM_" & markText
                shape.Line.Visible = msoFalse
                shape.Fill.Visible = msoFalse
                shape.TextFrame.TextRange.Text = markText
                shape.TextFrame.TextRange.Font.Size = 54
                shape.TextFrame.TextRange.Font.Bold = True
                shape.TextFrame.TextRange.Font.Color = wdColorGray25
                shape.TextFrame.TextRange.ParagraphFormat.Alignment = 1
                shape.Rotation = 315
                shape.ZOrder msoSendBehindText
                If Err.Number = 0 Then added = added + 1
                Err.Clear
            End If
        Else
            Err.Clear
        End If
    Next
    AddTextWatermark = added
End Function

Sub ClearOwnWatermarks(doc)
    On Error Resume Next
    Dim section, shape, i, nm
    For Each section In doc.Sections
        For i = section.Headers(wdHeaderFooterPrimary).Shapes.Count To 1 Step -1
            Set shape = section.Headers(wdHeaderFooterPrimary).Shapes(i)
            nm = CStr(shape.Name)
            If InStr(1, nm, "OA_WM_", vbTextCompare) > 0 Then
                shape.Delete
            End If
            Err.Clear
        Next
    Next
End Sub

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
