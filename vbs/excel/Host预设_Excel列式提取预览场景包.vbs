' 函数名: HostExcelAiExtractColumnsPreviewPack
' 描述: 列式提取预览场景包：先预览约20行规则提取电话/单号/金额/邮箱/证件，确认后写新表，不改源表
' 适用应用: Excel
' 搜索范围: 当前范围
' 搜索对象: 无
' 作用范围: 当前范围
' 输出类型: 表

Option Explicit

Function Main()
    On Error Resume Next
    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = FailureJson("E_NO_EXCEL_APP", "未取得 Excel 应用，请在 Excel 中运行该预设")
        Exit Function
    End If
    Main = HostExcelAiExtractColumnsPreviewPack(appObj)
End Function

Function HostExcelAiExtractColumnsPreviewPack(appObj)
    On Error Resume Next
    Dim sourceRange, sourceSheet, sourceName, outputMode
    Set sourceRange = appObj.Selection.CurrentRegion
    If Err.Number <> 0 Or TypeName(sourceRange) = "Empty" Or TypeName(sourceRange) = "Nothing" Then
        Err.Clear
        HostExcelAiExtractColumnsPreviewPack = FailureJson("E_NO_RANGE", "请先选中带表头的数据区域")
        Exit Function
    End If
    If sourceRange.Rows.Count < 2 Then
        HostExcelAiExtractColumnsPreviewPack = FailureJson("E_RANGE_TOO_SMALL", "当前区域至少需要一行表头和一行数据")
        Exit Function
    End If
    Set sourceSheet = sourceRange.Worksheet
    sourceName = sourceSheet.Name
    outputMode = NormalizeOutputMode(SafePrompt("输出模式：NewSheet 或 NewWorkbook（默认 NewSheet）", "NewSheet"))

    Dim sourceField, sourceCol, extractSpec, extractKinds(), kindCount
    sourceField = Trim(SafePrompt("待提取文本列（表头名）", CStr(sourceRange.Cells(1, 1).Text)))
    sourceCol = FindHeaderColumn(sourceRange, sourceField)
    If sourceCol <= 0 Then
        HostExcelAiExtractColumnsPreviewPack = FailureJson("E_FIELD_NOT_FOUND", "未找到待提取列：" & sourceField)
        Exit Function
    End If
    extractSpec = Trim(SafePrompt("提取项（逗号分隔，可选 phone,order,amount,email,idcard）", "phone,order,amount,email,idcard"))
    kindCount = ParseExtractKinds(extractSpec, extractKinds)
    If kindCount <= 0 Then
        HostExcelAiExtractColumnsPreviewPack = FailureJson("E_NO_EXTRACT_KIND", "至少选择一个提取项，例如 phone,order,amount")
        Exit Function
    End If

    Dim dataRows, previewRows, previewText, i, rowIndex, sampleLine, cellText
    Dim hitCounts()
    dataRows = sourceRange.Rows.Count - 1
    previewRows = dataRows
    If previewRows > 20 Then previewRows = 20
    ReDim hitCounts(kindCount)
    For i = 1 To kindCount
        hitCounts(i) = 0
    Next

    previewText = "Excel 列式提取预览场景包（先预览约20行，尚未写入）" & vbCrLf & _
        "命令ID：excel.ai_extract_columns_preview" & vbCrLf & _
        "源表=" & sourceName & "；区域=" & sourceRange.Address & vbCrLf & _
        "输出模式=" & outputMode & vbCrLf & _
        "文本列=" & sourceField & "；提取项=" & JoinExtractKinds(extractKinds, kindCount) & vbCrLf & _
        "数据行=" & CStr(dataRows) & "；预览行=" & CStr(previewRows) & vbCrLf & _
        "默认写新表，不覆盖原列；sourceUnchanged=true" & vbCrLf & _
        "---- 预览样本 ----"

    For i = 1 To previewRows
        rowIndex = i + 1
        cellText = CStr(sourceRange.Cells(rowIndex, sourceCol).Text)
        sampleLine = "R" & CStr(rowIndex) & ": "
        Dim k, extracted
        For k = 1 To kindCount
            extracted = ExtractByKind(cellText, extractKinds(k))
            If Len(extracted) > 0 Then hitCounts(k) = hitCounts(k) + 1
            sampleLine = sampleLine & extractKinds(k) & "=" & extracted
            If k < kindCount Then sampleLine = sampleLine & "；"
        Next
        If i <= 8 Then previewText = previewText & vbCrLf & sampleLine
    Next
    If previewRows > 8 Then previewText = previewText & vbCrLf & "... 其余预览行确认后写入结果表"
    previewText = previewText & vbCrLf & "预览命中：" & BuildHitSummary(extractKinds, hitCounts, kindCount)
    previewText = previewText & vbCrLf & "AIGEN验收骨架：可用 =AIGEN(""从文本提取" & JoinExtractKinds(extractKinds, kindCount) & """) 对比规则结果；仅建议，不自动改源表。"

    Dim planId, planPreview
    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_ai_extract_columns_preview", "{""source"":""" & EscapeJson(sourceRange.Address) & """,""field"":""" & EscapeJson(sourceField) & """,""kinds"":""" & EscapeJson(JoinExtractKinds(extractKinds, kindCount)) & """,""previewRows"":" & CStr(previewRows) & "}", "office.excel.aiExtractColumnsPreview"
    planPreview = SafePreviewWritePlan()
    SafeRollbackWritePlan planId

    If Not SafeConfirmStep(previewText & vbCrLf & vbCrLf & "确认按规则全量提取到新表？", "excel.ai_extract_columns_preview") Then
        HostExcelAiExtractColumnsPreviewPack = FailureJson("E_CONFIRM_REQUIRED", "用户取消或未确认，未修改工作簿")
        Exit Function
    End If

    Dim createdSheets(), createdCount, resultSheet, summarySheet, outRows, fullHits()
    createdCount = 0
    ReDim createdSheets(4)
    ReDim fullHits(kindCount)
    For i = 1 To kindCount
        fullHits(i) = 0
    Next

    Set resultSheet = CreateOutputSheet(appObj, sourceName, "列式提取结果", outputMode)
    If resultSheet Is Nothing Then
        HostExcelAiExtractColumnsPreviewPack = FailureJson("E_OUTPUT_SHEET", "无法创建列式提取结果工作表")
        Exit Function
    End If
    createdCount = createdCount + 1: Set createdSheets(createdCount) = resultSheet

    If Not WriteExtractResultSheet(sourceRange, sourceCol, extractKinds, kindCount, resultSheet, outRows, fullHits) Then
        RollbackCreatedSheets appObj, createdSheets, createdCount
        HostExcelAiExtractColumnsPreviewPack = FailureJson("E_EXTRACT_WRITE", "列式提取结果写入失败，已删除未完成输出")
        Exit Function
    End If

    Set summarySheet = CreateOutputSheet(appObj, sourceName, "处理摘要", outputMode)
    If summarySheet Is Nothing Then
        RollbackCreatedSheets appObj, createdSheets, createdCount
        HostExcelAiExtractColumnsPreviewPack = FailureJson("E_OUTPUT_SHEET", "无法创建处理摘要工作表")
        Exit Function
    End If
    createdCount = createdCount + 1: Set createdSheets(createdCount) = summarySheet

    Dim summaryLines, impactJson, summary, aigenHint
    aigenHint = "AIGEN验收骨架：抽20行用 =AIGEN(""提取电话/单号/金额/邮箱/证件"") 与规则列对比；差异人工定标后可再全量。"
    summaryLines = "场景名=列式提取预览场景包" & vbCrLf & _
        "命令ID=excel.ai_extract_columns_preview" & vbCrLf & _
        "源表名称=" & sourceName & vbCrLf & _
        "源区域=" & sourceRange.Address & vbCrLf & _
        "输出模式=" & outputMode & vbCrLf & _
        "输出位置=" & resultSheet.Name & "," & summarySheet.Name & vbCrLf & _
        "处理前行数=" & CStr(sourceRange.Rows.Count) & vbCrLf & _
        "处理后行数=" & CStr(outRows) & vbCrLf & _
        "参数=sourceField=" & sourceField & ";kinds=" & JoinExtractKinds(extractKinds, kindCount) & vbCrLf & _
        "预览行=" & CStr(previewRows) & vbCrLf & _
        "全量命中=" & BuildHitSummary(extractKinds, fullHits, kindCount) & vbCrLf & _
        "AIGEN提示=" & aigenHint & vbCrLf & _
        "执行时间=" & Now & vbCrLf & _
        "结果=成功"
    impactJson = SafeHostText("ExcelWriteImpactSummary", summaryLines, summarySheet.Name)
    If Not ExtractJsonBoolean(impactJson, "ok") Then WriteSummaryFallback summarySheet, summaryLines

    summary = "Excel 列式提取完成：结果=" & resultSheet.Name & "；数据行=" & CStr(outRows) & "；" & BuildHitSummary(extractKinds, fullHits, kindCount) & "；sourceUnchanged=true"
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostExcelAiExtractColumnsPreviewPack = "{""ok"":true,""sourceUnchanged"":true,""commandId"":""excel.ai_extract_columns_preview"",""resultSheet"":""" & EscapeJson(resultSheet.Name) & """,""summarySheet"":""" & EscapeJson(summarySheet.Name) & """,""rows"":" & CStr(outRows) & ",""kinds"":""" & EscapeJson(JoinExtractKinds(extractKinds, kindCount)) & """,""writePlanPreview"":""" & EscapeJson(planPreview) & """,""message"":""" & EscapeJson(summary) & """}"
End Function

Function ParseExtractKinds(spec, ByRef kinds())
    Dim parts, i, n, token, normalized
    parts = Split(Replace(CStr(spec), "，", ","), ",")
    n = 0
    ReDim kinds(UBound(parts) + 1)
    For i = 0 To UBound(parts)
        token = LCase(Trim(CStr(parts(i))))
        normalized = NormalizeExtractKind(token)
        If Len(normalized) > 0 Then
            If Not ExtractKindExists(kinds, n, normalized) Then
                n = n + 1
                kinds(n) = normalized
            End If
        End If
    Next
    ParseExtractKinds = n
End Function

Function NormalizeExtractKind(token)
    Dim t
    t = LCase(Trim(CStr(token)))
    If t = "phone" Or t = "tel" Or t = "mobile" Or t = "电话" Or t = "手机" Or t = "手机号" Then
        NormalizeExtractKind = "phone"
    ElseIf t = "order" Or t = "orderno" Or t = "单号" Or t = "订单号" Or t = "订单编号" Then
        NormalizeExtractKind = "order"
    ElseIf t = "amount" Or t = "money" Or t = "金额" Or t = "价格" Then
        NormalizeExtractKind = "amount"
    ElseIf t = "email" Or t = "mail" Or t = "邮箱" Or t = "邮件" Then
        NormalizeExtractKind = "email"
    ElseIf t = "idcard" Or t = "id" Or t = "身份证" Or t = "证件号" Then
        NormalizeExtractKind = "idcard"
    Else
        NormalizeExtractKind = ""
    End If
End Function

Function ExtractKindExists(ByRef kinds(), kindCount, kindName)
    Dim i
    ExtractKindExists = False
    For i = 1 To kindCount
        If kinds(i) = kindName Then
            ExtractKindExists = True
            Exit Function
        End If
    Next
End Function

Function JoinExtractKinds(ByRef kinds(), kindCount)
    Dim i, buf
    buf = ""
    For i = 1 To kindCount
        If Len(buf) > 0 Then buf = buf & ","
        buf = buf & kinds(i)
    Next
    JoinExtractKinds = buf
End Function

Function BuildHitSummary(ByRef kinds(), ByRef hits(), kindCount)
    Dim i, buf
    buf = ""
    For i = 1 To kindCount
        If Len(buf) > 0 Then buf = buf & "；"
        buf = buf & kinds(i) & "命中" & CStr(hits(i))
    Next
    BuildHitSummary = buf
End Function

Function ExtractKindHeader(kindName)
    Select Case CStr(kindName)
        Case "phone": ExtractKindHeader = "提取_电话"
        Case "order": ExtractKindHeader = "提取_单号"
        Case "amount": ExtractKindHeader = "提取_金额"
        Case "email": ExtractKindHeader = "提取_邮箱"
        Case "idcard": ExtractKindHeader = "提取_证件号"
        Case Else: ExtractKindHeader = "提取_" & CStr(kindName)
    End Select
End Function

Function ExtractByKind(textValue, kindName)
    Dim t
    t = CStr(textValue)
    Select Case CStr(kindName)
        Case "phone": ExtractByKind = ExtractPhone(t)
        Case "order": ExtractByKind = ExtractOrderNo(t)
        Case "amount": ExtractByKind = ExtractAmount(t)
        Case "email": ExtractByKind = ExtractEmail(t)
        Case "idcard": ExtractByKind = ExtractIdCard(t)
        Case Else: ExtractByKind = ""
    End Select
End Function

Function ExtractPhone(textValue)
    Dim digits, i, cand
    digits = DigitsOnly(textValue)
    ExtractPhone = ""
    If Len(digits) < 11 Then Exit Function
    For i = 1 To Len(digits) - 10
        cand = Mid(digits, i, 11)
        If Left(cand, 1) = "1" Then
            ExtractPhone = cand
            Exit Function
        End If
    Next
    cand = Right(digits, 11)
    If Left(cand, 1) = "1" Then ExtractPhone = cand
End Function

Function ExtractOrderNo(textValue)
    Dim t, i, ch, buf, best
    t = UCase(CStr(textValue))
    buf = ""
    best = ""
    For i = 1 To Len(t)
        ch = Mid(t, i, 1)
        If (ch >= "0" And ch <= "9") Or (ch >= "A" And ch <= "Z") Or ch = "-" Or ch = "_" Then
            buf = buf & ch
        Else
            If IsLikelyOrderToken(buf) Then
                If Len(buf) > Len(best) Then best = buf
            End If
            buf = ""
        End If
    Next
    If IsLikelyOrderToken(buf) Then
        If Len(buf) > Len(best) Then best = buf
    End If
    ExtractOrderNo = best
End Function

Function IsLikelyOrderToken(token)
    Dim t, hasDigit, hasLetter, i, ch
    t = CStr(token)
    IsLikelyOrderToken = False
    If Len(t) < 6 Or Len(t) > 32 Then Exit Function
    hasDigit = False
    hasLetter = False
    For i = 1 To Len(t)
        ch = Mid(t, i, 1)
        If ch >= "0" And ch <= "9" Then hasDigit = True
        If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Then hasLetter = True
    Next
    If hasDigit And (hasLetter Or Len(t) >= 8) Then IsLikelyOrderToken = True
End Function

Function ExtractAmount(textValue)
    Dim t, i, ch, buf, best, inNum
    t = Replace(Replace(CStr(textValue), ",", ""), "￥", "")
    t = Replace(Replace(t, "元", " "), "¥", "")
    buf = ""
    best = ""
    inNum = False
    For i = 1 To Len(t)
        ch = Mid(t, i, 1)
        If (ch >= "0" And ch <= "9") Or (ch = "." And inNum) Then
            buf = buf & ch
            inNum = True
        Else
            If inNum Then
                If IsValidAmountToken(buf) Then
                    If Len(buf) >= Len(best) Then best = buf
                End If
                buf = ""
                inNum = False
            End If
        End If
    Next
    If inNum And IsValidAmountToken(buf) Then
        If Len(buf) >= Len(best) Then best = buf
    End If
    ExtractAmount = best
End Function

Function IsValidAmountToken(token)
    Dim t, parts
    t = CStr(token)
    IsValidAmountToken = False
    If Len(t) = 0 Then Exit Function
    If Left(t, 1) = "." Or Right(t, 1) = "." Then Exit Function
    parts = Split(t, ".")
    If UBound(parts) > 1 Then Exit Function
    If UBound(parts) = 1 Then
        If Len(parts(1)) > 2 Then Exit Function
    End If
    IsValidAmountToken = True
End Function

Function ExtractEmail(textValue)
    Dim t, i, ch, buf, best
    t = LCase(CStr(textValue))
    buf = ""
    best = ""
    For i = 1 To Len(t)
        ch = Mid(t, i, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Or ch = "." Or ch = "_" Or ch = "%" Or ch = "+" Or ch = "-" Or ch = "@" Then
            buf = buf & ch
        Else
            If IsValidEmailToken(buf) Then best = buf
            buf = ""
        End If
    Next
    If IsValidEmailToken(buf) Then best = buf
    ExtractEmail = best
End Function

Function IsValidEmailToken(token)
    Dim t, atPos, dotPos
    t = CStr(token)
    IsValidEmailToken = False
    atPos = InStr(t, "@")
    If atPos <= 1 Then Exit Function
    If InStr(atPos + 1, t, "@") > 0 Then Exit Function
    dotPos = InStr(atPos + 1, t, ".")
    If dotPos <= atPos + 1 Then Exit Function
    If dotPos >= Len(t) Then Exit Function
    IsValidEmailToken = True
End Function

Function ExtractIdCard(textValue)
    Dim t, i, ch, buf, best
    t = UCase(CStr(textValue))
    buf = ""
    best = ""
    For i = 1 To Len(t)
        ch = Mid(t, i, 1)
        If (ch >= "0" And ch <= "9") Or ch = "X" Then
            buf = buf & ch
        Else
            If IsValidIdToken(buf) Then best = buf
            buf = ""
        End If
    Next
    If IsValidIdToken(buf) Then best = buf
    ExtractIdCard = best
End Function

Function IsValidIdToken(token)
    Dim t
    t = UCase(CStr(token))
    IsValidIdToken = False
    If Len(t) = 15 Then
        If DigitsOnly(t) = t Then IsValidIdToken = True
    ElseIf Len(t) = 18 Then
        If DigitsOnly(Left(t, 17)) = Left(t, 17) Then
            If (Right(t, 1) >= "0" And Right(t, 1) <= "9") Or Right(t, 1) = "X" Then IsValidIdToken = True
        End If
    End If
End Function

Function WriteExtractResultSheet(sourceRange, sourceCol, ByRef kinds(), kindCount, sheetObj, ByRef outRows, ByRef hits())
    On Error Resume Next
    Dim c, r, k, srcCols, cellText, extracted, note
    WriteExtractResultSheet = False
    outRows = 0
    srcCols = sourceRange.Columns.Count

    sheetObj.Cells(1, 1).Value = "_来源行"
    For c = 1 To srcCols
        sheetObj.Cells(1, c + 1).Value = sourceRange.Cells(1, c).Text
    Next
    For k = 1 To kindCount
        sheetObj.Cells(1, srcCols + 1 + k).Value = ExtractKindHeader(kinds(k))
    Next
    sheetObj.Cells(1, srcCols + kindCount + 2).Value = "_提取说明"

    For r = 2 To sourceRange.Rows.Count
        outRows = outRows + 1
        sheetObj.Cells(outRows + 1, 1).Value = sourceRange.Row + r - 1
        For c = 1 To srcCols
            sheetObj.Cells(outRows + 1, c + 1).Value = sourceRange.Cells(r, c).Value
        Next
        cellText = CStr(sourceRange.Cells(r, sourceCol).Text)
        note = ""
        For k = 1 To kindCount
            extracted = ExtractByKind(cellText, kinds(k))
            sheetObj.Cells(outRows + 1, srcCols + 1 + k).Value = extracted
            If Len(extracted) > 0 Then
                hits(k) = hits(k) + 1
                If Len(note) > 0 Then note = note & "；"
                note = note & kinds(k) & "命中"
            End If
        Next
        If Len(note) = 0 Then note = "未命中规则"
        sheetObj.Cells(outRows + 1, srcCols + kindCount + 2).Value = note
        If Err.Number <> 0 Then
            Err.Clear
            Exit Function
        End If
    Next

    sheetObj.Rows(1).Font.Bold = True
    sheetObj.Columns.AutoFit
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    WriteExtractResultSheet = True
End Function

Function ExtractJsonBoolean(jsonText, keyName)
    Dim marker, keyPos, colonPos, valueText
    ExtractJsonBoolean = False
    marker = """" & CStr(keyName) & """"
    keyPos = InStr(1, CStr(jsonText), marker, vbTextCompare)
    If keyPos <= 0 Then Exit Function
    colonPos = InStr(keyPos + Len(marker), CStr(jsonText), ":")
    If colonPos <= 0 Then Exit Function
    valueText = LCase(LTrim(Mid(CStr(jsonText), colonPos + 1)))
    ExtractJsonBoolean = (Left(valueText, 4) = "true")
End Function

Function ExtractJsonString(jsonText, keyName)
    Dim marker, keyPos, colonPos, quotePos, i, ch, nextCh, result
    ExtractJsonString = ""
    marker = """" & CStr(keyName) & """"
    keyPos = InStr(1, CStr(jsonText), marker, vbTextCompare)
    If keyPos <= 0 Then Exit Function
    colonPos = InStr(keyPos + Len(marker), CStr(jsonText), ":")
    If colonPos <= 0 Then Exit Function
    quotePos = InStr(colonPos + 1, CStr(jsonText), """")
    If quotePos <= 0 Then Exit Function
    result = ""
    i = quotePos + 1
    Do While i <= Len(jsonText)
        ch = Mid(jsonText, i, 1)
        If ch = """" Then Exit Do
        If ch = "\" And i < Len(jsonText) Then
            nextCh = Mid(jsonText, i + 1, 1)
            result = result & nextCh
            i = i + 2
        Else
            result = result & ch
            i = i + 1
        End If
    Loop
    ExtractJsonString = result
End Function

Function ExtractJsonNumber(jsonText, keyName)
    Dim marker, keyPos, colonPos, valueText, ch, i, buf
    ExtractJsonNumber = 0
    marker = """" & CStr(keyName) & """"
    keyPos = InStr(1, CStr(jsonText), marker, vbTextCompare)
    If keyPos <= 0 Then Exit Function
    colonPos = InStr(keyPos + Len(marker), CStr(jsonText), ":")
    If colonPos <= 0 Then Exit Function
    valueText = LTrim(Mid(CStr(jsonText), colonPos + 1))
    buf = ""
    For i = 1 To Len(valueText)
        ch = Mid(valueText, i, 1)
        If (ch >= "0" And ch <= "9") Or ch = "-" Or ch = "." Then
            buf = buf & ch
        ElseIf Len(buf) > 0 Then
            Exit For
        End If
    Next
    If Len(buf) > 0 Then ExtractJsonNumber = CDbl(buf)
End Function

Function ParseJsonIntArray(jsonText, keyName)
    Dim marker, keyPos, bracketStart, bracketEnd, body, parts, i, n, values()
    ParseJsonIntArray = Array()
    marker = """" & CStr(keyName) & """"
    keyPos = InStr(1, CStr(jsonText), marker, vbTextCompare)
    If keyPos <= 0 Then Exit Function
    bracketStart = InStr(keyPos, CStr(jsonText), "[")
    If bracketStart <= 0 Then Exit Function
    bracketEnd = InStr(bracketStart + 1, CStr(jsonText), "]")
    If bracketEnd <= bracketStart Then Exit Function
    body = Trim(Mid(CStr(jsonText), bracketStart + 1, bracketEnd - bracketStart - 1))
    If Len(body) = 0 Then Exit Function
    parts = Split(body, ",")
    n = -1
    ReDim values(UBound(parts))
    For i = 0 To UBound(parts)
        If Len(Trim(parts(i))) > 0 Then
            n = n + 1
            values(n) = CLng(Trim(parts(i)))
        End If
    Next
    If n >= 0 Then
        ReDim Preserve values(n)
        ParseJsonIntArray = values
    End If
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
    If Err.Number <> 0 Then
        SafePrompt = CStr(defaultValue)
        Err.Clear
    End If
End Function

Function SafeConfirmStep(message, stepInfo)
    On Error Resume Next
    Dim confirmed, answer
    confirmed = Host.ConfirmStep(message, stepInfo)
    If Err.Number = 0 Then
        SafeConfirmStep = CBool(confirmed)
        Exit Function
    End If
    Err.Clear
    answer = Trim(SafePrompt(message & vbCrLf & vbCrLf & "确认继续请输入：确认", ""))
    SafeConfirmStep = (StrComp(answer, "确认", vbTextCompare) = 0)
End Function

Function SafeBeginWritePlan()
    On Error Resume Next
    SafeBeginWritePlan = Host.BeginWritePlan("")
    If Err.Number <> 0 Then SafeBeginWritePlan = ""
    Err.Clear
End Function

Sub SafeRecordWrite(actionId, paramsJson, capabilityId)
    On Error Resume Next
    Host.RecordWrite actionId, paramsJson, "office", "excel.sheet.delete", "{}", capabilityId
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

Function SafeHostText(methodName, arg1, arg2, arg3, arg4)
    On Error Resume Next
    SafeHostText = ""
    If methodName = "ExcelGetHeaderMap" Then
        SafeHostText = Host.ExcelGetHeaderMap(arg1)
    ElseIf methodName = "ExcelCreateOutputSheet" Then
        SafeHostText = Host.ExcelCreateOutputSheet(arg1, arg2, arg3, arg4)
    ElseIf methodName = "ExcelWriteImpactSummary" Then
        SafeHostText = Host.ExcelWriteImpactSummary(arg1, arg2)
    ElseIf methodName = "ExcelPreviewRowSamples" Then
        SafeHostText = Host.ExcelPreviewRowSamples(arg1, arg2)
    ElseIf methodName = "ExcelDedupePlan" Then
        SafeHostText = Host.ExcelDedupePlan(arg1, arg2)
    ElseIf methodName = "ExcelValidatePlan" Then
        SafeHostText = Host.ExcelValidatePlan(arg1)
    End If
    If Err.Number <> 0 Then SafeHostText = ""
    Err.Clear
End Function

Function NormalizeCellText(value)
    Dim text
    text = CStr(value)
    text = Replace(text, ChrW(160), " ")
    text = Replace(text, ChrW(12288), " ")
    NormalizeCellText = Trim(text)
End Function

Function NormalizeKey(value)
    NormalizeKey = LCase(NormalizeCellText(value))
End Function

Function FindHeaderColumn(sourceRange, fieldName)
    Dim colIndex
    FindHeaderColumn = 0
    For colIndex = 1 To sourceRange.Columns.Count
        If StrComp(NormalizeCellText(sourceRange.Cells(1, colIndex).Text), NormalizeCellText(fieldName), vbTextCompare) = 0 Then
            FindHeaderColumn = colIndex
            Exit Function
        End If
    Next
End Function

Function SplitCsvFields(text)
    Dim cleaned, parts, i, n, values()
    cleaned = Replace(Replace(Trim(CStr(text)), "，", ","), "；", ",")
    If Len(cleaned) = 0 Then
        SplitCsvFields = Array()
        Exit Function
    End If
    parts = Split(cleaned, ",")
    n = -1
    ReDim values(UBound(parts))
    For i = 0 To UBound(parts)
        If Len(Trim(parts(i))) > 0 Then
            n = n + 1
            values(n) = Trim(parts(i))
        End If
    Next
    If n < 0 Then
        SplitCsvFields = Array()
    Else
        ReDim Preserve values(n)
        SplitCsvFields = values
    End If
End Function

Function ResolveHeaderColumns(sourceRange, fieldNames)
    Dim i, colIndex, n, cols()
    ResolveHeaderColumns = Array()
    If Not IsArray(fieldNames) Then Exit Function
    On Error Resume Next
    If UBound(fieldNames) < 0 Then Exit Function
    Err.Clear
    n = -1
    ReDim cols(UBound(fieldNames))
    For i = 0 To UBound(fieldNames)
        colIndex = FindHeaderColumn(sourceRange, fieldNames(i))
        If colIndex <= 0 Then
            ResolveHeaderColumns = Array()
            Exit Function
        End If
        n = n + 1
        cols(n) = colIndex
    Next
    If n >= 0 Then
        ReDim Preserve cols(n)
        ResolveHeaderColumns = cols
    End If
End Function

Function UniqueSheetName(workbook, baseName)
    Dim candidate, index, exists, ws
    index = 1
    Do
        If index = 1 Then
            candidate = Left(CStr(baseName), 31)
        Else
            candidate = Left(CStr(baseName), 28) & CStr(index)
        End If
        exists = False
        For Each ws In workbook.Worksheets
            If StrComp(ws.Name, candidate, vbTextCompare) = 0 Then exists = True
        Next
        If Not exists Then
            UniqueSheetName = candidate
            Exit Function
        End If
        index = index + 1
    Loop
End Function

Function BuildTimestamp()
    Dim d
    d = Now
    BuildTimestamp = Year(d) & Right("0" & Month(d), 2) & Right("0" & Day(d), 2) & "_" & Right("0" & Hour(d), 2) & Right("0" & Minute(d), 2) & Right("0" & Second(d), 2)
End Function

Function BuildOutputSheetName(sourceSheetName, functionName)
    Dim baseName
    baseName = CStr(sourceSheetName) & "_" & CStr(functionName) & "_" & BuildTimestamp()
    baseName = Replace(baseName, "[", "(")
    baseName = Replace(baseName, "]", ")")
    baseName = Replace(baseName, ":", "-")
    baseName = Replace(baseName, "\\", "-")
    baseName = Replace(baseName, "/", "-")
    baseName = Replace(baseName, "?", "")
    baseName = Replace(baseName, "*", "")
    If Len(baseName) > 31 Then baseName = Left(baseName, 31)
    BuildOutputSheetName = baseName
End Function

Function CreateOutputSheet(appObj, sourceSheetName, functionName, outputMode)
    On Error Resume Next
    Dim hostJson, sheetName, workbook, sheetObj, modeText
    Set CreateOutputSheet = Nothing
    modeText = NormalizeOutputMode(outputMode)
    hostJson = SafeHostText("ExcelCreateOutputSheet", functionName, modeText, sourceSheetName, True)
    sheetName = ExtractJsonString(hostJson, "sheetName")
    If Len(sheetName) = 0 Then sheetName = ExtractJsonString(hostJson, "name")
    If Len(sheetName) > 0 Then
        Set sheetObj = Nothing
        Set sheetObj = appObj.ActiveWorkbook.Worksheets(sheetName)
        If Not sheetObj Is Nothing Then
            Set CreateOutputSheet = sheetObj
            Exit Function
        End If
    End If

    Set workbook = appObj.ActiveWorkbook
    If modeText = "NewWorkbook" Then
        Set workbook = appObj.Workbooks.Add
    End If
    sheetName = UniqueSheetName(workbook, BuildOutputSheetName(sourceSheetName, functionName))
    Set sheetObj = workbook.Worksheets.Add
    sheetObj.Name = sheetName
    If Err.Number <> 0 Then
        Err.Clear
        Set CreateOutputSheet = Nothing
        Exit Function
    End If
    Set CreateOutputSheet = sheetObj
End Function

Sub SafeDeleteSheet(appObj, sheetObj)
    On Error Resume Next
    Dim prev
    If sheetObj Is Nothing Then Exit Sub
    prev = appObj.DisplayAlerts
    appObj.DisplayAlerts = False
    sheetObj.Delete
    appObj.DisplayAlerts = prev
    Err.Clear
End Sub

Function NormalizeOutputMode(text)
    Dim t
    t = LCase(Trim(CStr(text)))
    If t = "newworkbook" Or t = "workbook" Or t = "新工作簿" Then
        NormalizeOutputMode = "NewWorkbook"
    Else
        NormalizeOutputMode = "NewSheet"
    End If
End Function

Sub WriteSummaryFallback(sheetObj, summaryLines)
    Dim lines, i
    lines = Split(CStr(summaryLines), vbCrLf)
    sheetObj.Cells(1, 1).Value = "项目"
    sheetObj.Cells(1, 2).Value = "内容"
    For i = 0 To UBound(lines)
        sheetObj.Cells(i + 2, 1).Value = "行" & CStr(i + 1)
        sheetObj.Cells(i + 2, 2).Value = lines(i)
    Next
    sheetObj.Rows(1).Font.Bold = True
    sheetObj.Columns.AutoFit
End Sub

Sub RollbackCreatedSheets(appObj, createdSheets, createdCount)
    Dim i
    For i = createdCount To 1 Step -1
        SafeDeleteSheet appObj, createdSheets(i)
        Set createdSheets(i) = Nothing
    Next
End Sub

Function IsBlankCell(value)
    IsBlankCell = (Len(NormalizeCellText(value)) = 0)
End Function

Function FindWorksheet(workbook, sheetName)
    On Error Resume Next
    Dim ws
    Set FindWorksheet = Nothing
    For Each ws In workbook.Worksheets
        If StrComp(ws.Name, Trim(CStr(sheetName)), vbTextCompare) = 0 Then
            Set FindWorksheet = ws
            Exit Function
        End If
    Next
End Function

Function FindOtherSheetName(workbook, currentName)
    Dim ws
    FindOtherSheetName = ""
    For Each ws In workbook.Worksheets
        If StrComp(ws.Name, currentName, vbTextCompare) <> 0 Then
            FindOtherSheetName = ws.Name
            Exit Function
        End If
    Next
End Function

Function FindKeyRow(dataRange, keyColumn, keyText)
    Dim rowIndex
    FindKeyRow = 0
    For rowIndex = 2 To dataRange.Rows.Count
        If NormalizeKey(dataRange.Cells(rowIndex, keyColumn).Text) = keyText Then
            FindKeyRow = rowIndex
            Exit Function
        End If
    Next
End Function

Function CountKeyOccurrences(dataRange, keyColumn, keyText)
    Dim rowIndex, total
    total = 0
    For rowIndex = 2 To dataRange.Rows.Count
        If NormalizeKey(dataRange.Cells(rowIndex, keyColumn).Text) = keyText Then total = total + 1
    Next
    CountKeyOccurrences = total
End Function

Function TargetAbsoluteRow(targetRange, relativeRow)
    If relativeRow > 0 Then TargetAbsoluteRow = targetRange.Row + relativeRow - 1 Else TargetAbsoluteRow = 0
End Function
Function DigitsOnly(value)
    Dim i, ch, buf
    buf = ""
    For i = 1 To Len(CStr(value))
        ch = Mid(CStr(value), i, 1)
        If ch >= "0" And ch <= "9" Then buf = buf & ch
    Next
    DigitsOnly = buf
End Function
