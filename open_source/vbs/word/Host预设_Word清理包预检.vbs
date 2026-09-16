' 函数名: HostWordCleanupPackPreflight
' 描述: 预检分页符/分节符/空段/连续空段/首尾空白/多余空格影响量后，确认词“清理”才清理空白；默认保留分页与分节
' 适用应用: Word
' 搜索范围: 全文
' 搜索对象: 无

Option Explicit

Const wdWithInTable = 12

Function Main()
    On Error Resume Next
    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = FailureJson("E_NO_WORD_APP", "未取得 Word 应用，请在 Word 中运行该预设")
        Exit Function
    End If
    Main = HostWordCleanupPackPreflight(appObj)
End Function

Function HostWordCleanupPackPreflight(appObj)
    On Error Resume Next

    Dim doc
    Set doc = appObj.ActiveDocument
    If Err.Number <> 0 Or TypeName(doc) = "Empty" Or TypeName(doc) = "Nothing" Then
        Err.Clear
        HostWordCleanupPackPreflight = FailureJson("E_NO_DOCUMENT", "当前没有活动 Word 文档")
        Exit Function
    End If

    Dim formatSummary, reviewSummary
    formatSummary = SafeHostText("GetWordFormatIssueSummary")
    reviewSummary = SafeHostText("GetWordReviewSummary")

    Dim totalParas, emptyParas, consecutiveEmptyExtras, leadingTrim, trailingTrim, multiSpaceParas
    Dim pageBreakMarkers, sectionBreakMarkers, tableSkipped
    totalParas = 0
    emptyParas = 0
    consecutiveEmptyExtras = 0
    leadingTrim = 0
    trailingTrim = 0
    multiSpaceParas = 0
    pageBreakMarkers = 0
    sectionBreakMarkers = 0
    tableSkipped = 0

    CountCleanupImpact doc, totalParas, emptyParas, consecutiveEmptyExtras, leadingTrim, trailingTrim, multiSpaceParas, pageBreakMarkers, sectionBreakMarkers, tableSkipped

    Dim impactTotal
    impactTotal = emptyParas + leadingTrim + trailingTrim + multiSpaceParas

    Dim previewText
    previewText = "Word 清理包预检（尚未写入）" & vbCrLf & _
        "扫描段落=" & CStr(totalParas) & "；可删空段=" & CStr(emptyParas) & _
        "；连续空段额外=" & CStr(consecutiveEmptyExtras) & vbCrLf & _
        "首空白可去=" & CStr(leadingTrim) & "；尾空白可去=" & CStr(trailingTrim) & _
        "；多余空格段=" & CStr(multiSpaceParas) & vbCrLf & _
        "分页符标记=" & CStr(pageBreakMarkers) & "；分节符标记=" & CStr(sectionBreakMarkers) & _
        "（默认保留，不删除）" & vbCrLf & _
        "表格内段落跳过=" & CStr(tableSkipped) & "；预计影响项=" & CStr(impactTotal) & vbCrLf & _
        "Host.GetWordFormatIssueSummary: " & formatSummary & vbCrLf & _
        "Host.GetWordReviewSummary: " & reviewSummary

    If impactTotal <= 0 Then
        Host.WriteClipboard previewText & vbCrLf & "无需清理"
        SafeWriteLog previewText
        HostWordCleanupPackPreflight = "{""ok"":true,""changed"":false,""message"":""" & EscapeJson(previewText & "；无需清理") & """,""totalParas"":" & CStr(totalParas) & "}"
        Exit Function
    End If

    Dim planId, previewJson
    planId = SafeBeginWritePlan()
    SafeRecordWrite "word_cleanup_pack", "{""scope"":""activeDocument"",""preserveBreaks"":true,""impactTotal"":" & CStr(impactTotal) & "}", "office.word.cleanup"
    previewJson = SafePreviewWritePlan()

    Dim confirmText
    confirmText = Trim(SafePrompt(previewText & vbCrLf & vbCrLf & "以上尚未写入。确认清理空白（保留分页/分节），请输入：清理", ""))
    If StrComp(confirmText, "清理", vbTextCompare) <> 0 Then
        SafeRollbackWritePlan planId
        HostWordCleanupPackPreflight = FailureJson("E_CONFIRM_REQUIRED", "未输入“清理”，已取消且未修改文档")
        Exit Function
    End If

    Dim removedEmpty, trimmedLeading, trimmedTrailing, collapsedSpaces
    removedEmpty = 0
    trimmedLeading = 0
    trimmedTrailing = 0
    collapsedSpaces = 0

    ApplyCleanupPack doc, removedEmpty, trimmedLeading, trimmedTrailing, collapsedSpaces
    SafeRollbackWritePlan planId

    Dim summary
    summary = "Word 清理包完成；删除空段=" & CStr(removedEmpty) & _
        "；去首空白=" & CStr(trimmedLeading) & _
        "；去尾空白=" & CStr(trimmedTrailing) & _
        "；压缩多空格段=" & CStr(collapsedSpaces) & vbCrLf & _
        "分页/分节默认保留" & vbCrLf & _
        "Format: " & formatSummary & vbCrLf & _
        "Review: " & reviewSummary & vbCrLf & _
        "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostWordCleanupPackPreflight = "{""ok"":true,""changed"":true,""message"":""" & EscapeJson(summary) & _
        """,""removedEmpty"":" & CStr(removedEmpty) & _
        ",""trimmedLeading"":" & CStr(trimmedLeading) & _
        ",""trimmedTrailing"":" & CStr(trimmedTrailing) & _
        ",""collapsedSpaces"":" & CStr(collapsedSpaces) & _
        ",""preserveBreaks"":true,""planId"":""" & EscapeJson(planId) & """}"
End Function

Sub CountCleanupImpact(doc, ByRef totalParas, ByRef emptyParas, ByRef consecutiveEmptyExtras, ByRef leadingTrim, ByRef trailingTrim, ByRef multiSpaceParas, ByRef pageBreakMarkers, ByRef sectionBreakMarkers, ByRef tableSkipped)
    On Error Resume Next
    Dim index, para, rawText, plainText, visibleText, inTable, prevEmpty
    prevEmpty = False
    totalParas = doc.Paragraphs.Count
    For index = 1 To doc.Paragraphs.Count
        Set para = doc.Paragraphs(index)
        rawText = ""
        Err.Clear
        rawText = CStr(para.Range.Text)
        Err.Clear

        pageBreakMarkers = pageBreakMarkers + CountChar(rawText, Chr(12))
        sectionBreakMarkers = sectionBreakMarkers + CountChar(rawText, Chr(14))

        inTable = False
        Err.Clear
        inTable = CBool(para.Range.Information(wdWithInTable))
        If Err.Number <> 0 Then
            inTable = False
            Err.Clear
        End If
        If inTable Then
            tableSkipped = tableSkipped + 1
            prevEmpty = False
        Else
            plainText = StripControlChars(rawText)
            visibleText = Trim(plainText)
            If Len(visibleText) = 0 Then
                If Not ContainsBreakMarker(rawText) Then
                    emptyParas = emptyParas + 1
                    If prevEmpty Then consecutiveEmptyExtras = consecutiveEmptyExtras + 1
                    prevEmpty = True
                Else
                    prevEmpty = False
                End If
            Else
                If Left(plainText, 1) = " " Then leadingTrim = leadingTrim + 1
                If Right(plainText, 1) = " " Then trailingTrim = trailingTrim + 1
                If InStr(plainText, "  ") > 0 Then multiSpaceParas = multiSpaceParas + 1
                prevEmpty = False
            End If
        End If
    Next
End Sub

Sub ApplyCleanupPack(doc, ByRef removedEmpty, ByRef trimmedLeading, ByRef trimmedTrailing, ByRef collapsedSpaces)
    On Error Resume Next
    Dim index, para, rawText, plainText, visibleText, newText, inTable

    For index = doc.Paragraphs.Count To 1 Step -1
        Set para = doc.Paragraphs(index)
        inTable = False
        Err.Clear
        inTable = CBool(para.Range.Information(wdWithInTable))
        If Err.Number <> 0 Then
            inTable = False
            Err.Clear
        End If
        If Not inTable Then
            rawText = ""
            Err.Clear
            rawText = CStr(para.Range.Text)
            Err.Clear
            If Not ContainsBreakMarker(rawText) Then
                plainText = StripControlChars(rawText)
                visibleText = Trim(plainText)
                If Len(visibleText) = 0 Then
                    para.Range.Delete
                    If Err.Number = 0 Then removedEmpty = removedEmpty + 1
                    Err.Clear
                Else
                    newText = plainText
                    If Left(newText, 1) = " " Then
                        Do While Left(newText, 1) = " "
                            newText = Mid(newText, 2)
                        Loop
                        trimmedLeading = trimmedLeading + 1
                    End If
                    If Len(newText) > 0 And Right(newText, 1) = " " Then
                        Do While Len(newText) > 0 And Right(newText, 1) = " "
                            newText = Left(newText, Len(newText) - 1)
                        Loop
                        trimmedTrailing = trimmedTrailing + 1
                    End If
                    If InStr(newText, "  ") > 0 Then
                        Do While InStr(newText, "  ") > 0
                            newText = Replace(newText, "  ", " ")
                        Loop
                        collapsedSpaces = collapsedSpaces + 1
                    End If
                    If newText <> plainText Then
                        ReplaceParagraphText para, newText
                    End If
                End If
            End If
        End If
    Next
End Sub

Sub ReplaceParagraphText(para, newText)
    On Error Resume Next
    Dim rng
    Set rng = para.Range
    If rng.End > rng.Start Then
        rng.End = rng.End - 1
    End If
    rng.Text = CStr(newText)
    Err.Clear
End Sub

Function ContainsBreakMarker(value)
    Dim text
    text = CStr(value)
    If InStr(text, Chr(12)) > 0 Then
        ContainsBreakMarker = True
        Exit Function
    End If
    If InStr(text, Chr(14)) > 0 Then
        ContainsBreakMarker = True
        Exit Function
    End If
    ContainsBreakMarker = False
End Function

Function StripControlChars(value)
    Dim text
    text = CStr(value)
    text = Replace(text, vbCr, "")
    text = Replace(text, vbLf, "")
    text = Replace(text, vbTab, " ")
    text = Replace(text, Chr(7), "")
    text = Replace(text, Chr(11), "")
    text = Replace(text, Chr(12), "")
    text = Replace(text, Chr(14), "")
    text = Replace(text, ChrW(160), " ")
    text = Replace(text, ChrW(12288), " ")
    StripControlChars = text
End Function

Function CountChar(value, ch)
    Dim text, i, n
    text = CStr(value)
    n = 0
    For i = 1 To Len(text)
        If Mid(text, i, 1) = ch Then n = n + 1
    Next
    CountChar = n
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

Function SafeHostText(methodName)
    On Error Resume Next
    If methodName = "GetWordFormatIssueSummary" Then
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
