' 函数名: HostExcelSceneOrdersPack
' 描述: 订单行业场景包：字段别名候选+用户确认映射，输出映射表与质检问题报告，不改源表
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
    Main = HostExcelSceneOrdersPack(appObj)
End Function

Function HostExcelSceneOrdersPack(appObj)
    On Error Resume Next
    Dim sourceRange, sourceSheet, sourceName, outputMode
    Set sourceRange = appObj.Selection.CurrentRegion
    If Err.Number <> 0 Or TypeName(sourceRange) = "Empty" Or TypeName(sourceRange) = "Nothing" Then
        Err.Clear
        HostExcelSceneOrdersPack = FailureJson("E_NO_RANGE", "请先选中带表头的订单数据区域")
        Exit Function
    End If
    If sourceRange.Rows.Count < 2 Then
        HostExcelSceneOrdersPack = FailureJson("E_RANGE_TOO_SMALL", "当前区域至少需要一行表头和一行数据")
        Exit Function
    End If
    Set sourceSheet = sourceRange.Worksheet
    sourceName = sourceSheet.Name
    outputMode = NormalizeOutputMode(SafePrompt("输出模式：NewSheet 或 NewWorkbook（默认 NewSheet）", "NewSheet"))

    Dim roleIds(), roleLabels(), roleAliases(), roleCount
    roleCount = BuildOrderRoleCatalog(roleIds, roleLabels, roleAliases)

    Dim mapCols(), mapLabels(), i, detectedCol, detectedName, userField, finalCol
    ReDim mapCols(roleCount)
    ReDim mapLabels(roleCount)

    Dim mappingPreview
    mappingPreview = "候选字段映射（别名仅供参考，需你确认）：" & vbCrLf
    For i = 1 To roleCount
        detectedCol = DetectAliasColumn(sourceRange, roleAliases(i))
        If detectedCol > 0 Then
            detectedName = CStr(sourceRange.Cells(1, detectedCol).Text)
        Else
            detectedName = ""
        End If
        userField = Trim(SafePrompt("字段[" & roleLabels(i) & "] 对应表头（可改；留空跳过）", detectedName))
        If Len(userField) = 0 Then
            mapCols(i) = 0
            mapLabels(i) = ""
            mappingPreview = mappingPreview & roleLabels(i) & " => （跳过）" & vbCrLf
        Else
            finalCol = FindHeaderColumn(sourceRange, userField)
            If finalCol <= 0 Then
                HostExcelSceneOrdersPack = FailureJson("E_FIELD_NOT_FOUND", "未找到字段：" & userField & "（角色=" & roleLabels(i) & "）")
                Exit Function
            End If
            mapCols(i) = finalCol
            mapLabels(i) = CStr(sourceRange.Cells(1, finalCol).Text)
            mappingPreview = mappingPreview & roleLabels(i) & " => " & mapLabels(i) & vbCrLf
        End If
    Next

    If mapCols(1) <= 0 Then
        HostExcelSceneOrdersPack = FailureJson("E_ORDER_NO_REQUIRED", "订单号字段必须确认映射，不能跳过")
        Exit Function
    End If

    Dim issueCount, blankKeyCount, dupKeyCount, badPhoneCount, badAmountCount, blankStatusCount
    Dim issueRows(), issueFields(), issueProblems(), issueValues()
    ScanOrderIssues sourceRange, mapCols, roleLabels, roleCount, issueCount, blankKeyCount, dupKeyCount, badPhoneCount, badAmountCount, blankStatusCount, issueRows, issueFields, issueProblems, issueValues

    Dim previewText, planId, planPreview, dataRows
    dataRows = sourceRange.Rows.Count - 1
    previewText = "Excel 订单行业场景包预览（尚未写入）" & vbCrLf & _
        "命令ID：excel.scene_orders_pack" & vbCrLf & _
        "源表=" & sourceName & "；区域=" & sourceRange.Address & vbCrLf & _
        "输出模式=" & outputMode & vbCrLf & _
        mappingPreview & _
        "数据行=" & CStr(dataRows) & "；问题数=" & CStr(issueCount) & vbCrLf & _
        "空白订单号=" & CStr(blankKeyCount) & "；重复订单号=" & CStr(dupKeyCount) & _
        "；异常电话=" & CStr(badPhoneCount) & "；异常金额=" & CStr(badAmountCount) & "；空白状态=" & CStr(blankStatusCount) & vbCrLf & _
        "将生成：字段映射确认表 + 订单质检问题 + 处理摘要；sourceUnchanged=true" & vbCrLf & _
        "说明：别名识别只给候选，不假装自动看懂任意表。" & vbCrLf & _
        "AIGEN验收骨架：可用 =AIGEN(""检查订单关键字段完整性与格式"") 对照问题清单。"

    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_scene_orders_pack", "{""source"":""" & EscapeJson(sourceRange.Address) & """,""issues"":" & CStr(issueCount) & ",""rows"":" & CStr(dataRows) & "}", "office.excel.sceneOrdersPack"
    planPreview = SafePreviewWritePlan()
    SafeRollbackWritePlan planId

    If Not SafeConfirmStep(previewText & vbCrLf & vbCrLf & "确认生成订单场景质检报告？", "excel.scene_orders_pack") Then
        HostExcelSceneOrdersPack = FailureJson("E_CONFIRM_REQUIRED", "用户取消或未确认，未修改工作簿")
        Exit Function
    End If

    Dim createdSheets(), createdCount, mapSheet, issueSheet, summarySheet
    createdCount = 0
    ReDim createdSheets(6)

    Set mapSheet = CreateOutputSheet(appObj, sourceName, "订单字段映射", outputMode)
    If mapSheet Is Nothing Then
        HostExcelSceneOrdersPack = FailureJson("E_OUTPUT_SHEET", "无法创建订单字段映射工作表")
        Exit Function
    End If
    createdCount = createdCount + 1: Set createdSheets(createdCount) = mapSheet
    If Not WriteOrderMappingSheet(mapSheet, roleLabels, mapLabels, mapCols, roleCount) Then
        RollbackCreatedSheets appObj, createdSheets, createdCount
        HostExcelSceneOrdersPack = FailureJson("E_MAP_WRITE", "字段映射表写入失败，已删除未完成输出")
        Exit Function
    End If

    Set issueSheet = CreateOutputSheet(appObj, sourceName, "订单质检问题", outputMode)
    If issueSheet Is Nothing Then
        RollbackCreatedSheets appObj, createdSheets, createdCount
        HostExcelSceneOrdersPack = FailureJson("E_OUTPUT_SHEET", "无法创建订单质检问题工作表")
        Exit Function
    End If
    createdCount = createdCount + 1: Set createdSheets(createdCount) = issueSheet
    If Not WriteOrderIssueSheet(issueSheet, issueRows, issueFields, issueProblems, issueValues, issueCount) Then
        RollbackCreatedSheets appObj, createdSheets, createdCount
        HostExcelSceneOrdersPack = FailureJson("E_ISSUE_WRITE", "质检问题表写入失败，已删除未完成输出")
        Exit Function
    End If

    Set summarySheet = CreateOutputSheet(appObj, sourceName, "处理摘要", outputMode)
    If summarySheet Is Nothing Then
        RollbackCreatedSheets appObj, createdSheets, createdCount
        HostExcelSceneOrdersPack = FailureJson("E_OUTPUT_SHEET", "无法创建处理摘要工作表")
        Exit Function
    End If
    createdCount = createdCount + 1: Set createdSheets(createdCount) = summarySheet

    Dim summaryLines, impactJson, summary, aigenHint, mappedCount
    mappedCount = 0
    For i = 1 To roleCount
        If mapCols(i) > 0 Then mappedCount = mappedCount + 1
    Next
    aigenHint = "AIGEN验收骨架：对问题清单前20条解释业务风险，并给出是否允许发货/对账的判断；不要直接改源表。"
    summaryLines = "场景名=订单行业场景包" & vbCrLf & _
        "命令ID=excel.scene_orders_pack" & vbCrLf & _
        "源表名称=" & sourceName & vbCrLf & _
        "源区域=" & sourceRange.Address & vbCrLf & _
        "输出模式=" & outputMode & vbCrLf & _
        "输出位置=" & mapSheet.Name & "," & issueSheet.Name & "," & summarySheet.Name & vbCrLf & _
        "处理前行数=" & CStr(sourceRange.Rows.Count) & vbCrLf & _
        "处理后行数=" & CStr(sourceRange.Rows.Count) & vbCrLf & _
        "参数=mappedFields=" & CStr(mappedCount) & "/" & CStr(roleCount) & vbCrLf & _
        "问题数=" & CStr(issueCount) & vbCrLf & _
        "空白订单号=" & CStr(blankKeyCount) & "；重复订单号=" & CStr(dupKeyCount) & vbCrLf & _
        "异常电话=" & CStr(badPhoneCount) & "；异常金额=" & CStr(badAmountCount) & "；空白状态=" & CStr(blankStatusCount) & vbCrLf & _
        "AIGEN提示=" & aigenHint & vbCrLf & _
        "执行时间=" & Now & vbCrLf & _
        "结果=成功"
    impactJson = SafeHostText("ExcelWriteImpactSummary", summaryLines, summarySheet.Name)
    If Not ExtractJsonBoolean(impactJson, "ok") Then WriteSummaryFallback summarySheet, summaryLines

    summary = "Excel 订单行业场景包完成：映射=" & mapSheet.Name & "；问题=" & CStr(issueCount) & "；sourceUnchanged=true"
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostExcelSceneOrdersPack = "{""ok"":true,""sourceUnchanged"":true,""commandId"":""excel.scene_orders_pack"",""mappingSheet"":""" & EscapeJson(mapSheet.Name) & """,""issueSheet"":""" & EscapeJson(issueSheet.Name) & """,""summarySheet"":""" & EscapeJson(summarySheet.Name) & """,""issueCount"":" & CStr(issueCount) & ",""writePlanPreview"":""" & EscapeJson(planPreview) & """,""message"":""" & EscapeJson(summary) & """}"
End Function

Function BuildOrderRoleCatalog(ByRef roleIds(), ByRef roleLabels(), ByRef roleAliases())
    Dim n
    n = 6
    ReDim roleIds(n)
    ReDim roleLabels(n)
    ReDim roleAliases(n)
    roleIds(1) = "order_no": roleLabels(1) = "订单号": roleAliases(1) = "订单号|订单编号|单号|order|order_no|订单ID"
    roleIds(2) = "customer": roleLabels(2) = "客户": roleAliases(2) = "客户|客户名|收货人|买家|customer|客户名称"
    roleIds(3) = "phone": roleLabels(3) = "电话": roleAliases(3) = "电话|手机|手机号|联系电话|phone|mobile"
    roleIds(4) = "amount": roleLabels(4) = "金额": roleAliases(4) = "金额|订单金额|应付|实付|amount|price|总额"
    roleIds(5) = "status": roleLabels(5) = "状态": roleAliases(5) = "状态|订单状态|status|履约状态"
    roleIds(6) = "order_date": roleLabels(6) = "下单日期": roleAliases(6) = "下单日期|订单日期|日期|创建|order_date|创建时间"
    BuildOrderRoleCatalog = n
End Function

Function DetectAliasColumn(sourceRange, aliasText)
    Dim aliases, i, col
    DetectAliasColumn = 0
    aliases = Split(Replace(Replace(CStr(aliasText), "，", "|"), ",", "|"), "|")
    For i = 0 To UBound(aliases)
        If Len(Trim(CStr(aliases(i)))) > 0 Then
            col = FindHeaderColumn(sourceRange, Trim(CStr(aliases(i))))
            If col > 0 Then
                DetectAliasColumn = col
                Exit Function
            End If
        End If
    Next
    ' fuzzy contains match on header
    Dim c, header, a
    For c = 1 To sourceRange.Columns.Count
        header = CStr(sourceRange.Cells(1, c).Text)
        For i = 0 To UBound(aliases)
            a = Trim(CStr(aliases(i)))
            If Len(a) > 0 Then
                If InStr(1, header, a, vbTextCompare) > 0 Or InStr(1, a, header, vbTextCompare) > 0 Then
                    DetectAliasColumn = c
                    Exit Function
                End If
            End If
        Next
    Next
End Function

Sub ScanOrderIssues(sourceRange, ByRef mapCols(), ByRef roleLabels(), roleCount, ByRef issueCount, ByRef blankKeyCount, ByRef dupKeyCount, ByRef badPhoneCount, ByRef badAmountCount, ByRef blankStatusCount, ByRef issueRows(), ByRef issueFields(), ByRef issueProblems(), ByRef issueValues())
    On Error Resume Next
    Dim cap, r, keyCol, phoneCol, amountCol, statusCol, keyText, phoneText, amountText, statusText
    Dim seenKeys(), seenCount, si, isDup
    cap = 5000
    issueCount = 0
    blankKeyCount = 0
    dupKeyCount = 0
    badPhoneCount = 0
    badAmountCount = 0
    blankStatusCount = 0
    ReDim issueRows(cap)
    ReDim issueFields(cap)
    ReDim issueProblems(cap)
    ReDim issueValues(cap)
    ReDim seenKeys(sourceRange.Rows.Count)
    seenCount = 0

    keyCol = mapCols(1)
    phoneCol = 0: amountCol = 0: statusCol = 0
    If roleCount >= 3 Then phoneCol = mapCols(3)
    If roleCount >= 4 Then amountCol = mapCols(4)
    If roleCount >= 5 Then statusCol = mapCols(5)

    For r = 2 To sourceRange.Rows.Count
        keyText = ""
        If keyCol > 0 Then keyText = Trim(CStr(sourceRange.Cells(r, keyCol).Text))
        If Len(keyText) = 0 Then
            blankKeyCount = blankKeyCount + 1
            AppendOrderIssue issueRows, issueFields, issueProblems, issueValues, issueCount, cap, sourceRange.Row + r - 1, roleLabels(1), "订单号为空", ""
        Else
            isDup = False
            For si = 1 To seenCount
                If StrComp(seenKeys(si), keyText, vbTextCompare) = 0 Then
                    isDup = True
                    Exit For
                End If
            Next
            If isDup Then
                dupKeyCount = dupKeyCount + 1
                AppendOrderIssue issueRows, issueFields, issueProblems, issueValues, issueCount, cap, sourceRange.Row + r - 1, roleLabels(1), "订单号重复", keyText
            Else
                seenCount = seenCount + 1
                seenKeys(seenCount) = keyText
            End If
        End If

        If phoneCol > 0 Then
            phoneText = Trim(CStr(sourceRange.Cells(r, phoneCol).Text))
            If Len(phoneText) > 0 Then
                If Not IsValidOrderPhone(phoneText) Then
                    badPhoneCount = badPhoneCount + 1
                    AppendOrderIssue issueRows, issueFields, issueProblems, issueValues, issueCount, cap, sourceRange.Row + r - 1, roleLabels(3), "电话格式异常", phoneText
                End If
            End If
        End If

        If amountCol > 0 Then
            amountText = Trim(CStr(sourceRange.Cells(r, amountCol).Text))
            If Len(amountText) > 0 Then
                If Not IsValidOrderAmount(amountText) Then
                    badAmountCount = badAmountCount + 1
                    AppendOrderIssue issueRows, issueFields, issueProblems, issueValues, issueCount, cap, sourceRange.Row + r - 1, roleLabels(4), "金额格式异常", amountText
                End If
            End If
        End If

        If statusCol > 0 Then
            statusText = Trim(CStr(sourceRange.Cells(r, statusCol).Text))
            If Len(statusText) = 0 Then
                blankStatusCount = blankStatusCount + 1
                AppendOrderIssue issueRows, issueFields, issueProblems, issueValues, issueCount, cap, sourceRange.Row + r - 1, roleLabels(5), "状态为空", ""
            End If
        End If
        If Err.Number <> 0 Then Err.Clear
    Next
End Sub

Sub AppendOrderIssue(ByRef issueRows(), ByRef issueFields(), ByRef issueProblems(), ByRef issueValues(), ByRef issueCount, cap, rowNo, fieldName, problem, valueText)
    If issueCount >= cap Then Exit Sub
    issueRows(issueCount) = rowNo
    issueFields(issueCount) = fieldName
    issueProblems(issueCount) = problem
    issueValues(issueCount) = valueText
    issueCount = issueCount + 1
End Sub

Function IsValidOrderPhone(textValue)
    Dim digits
    digits = DigitsOnly(textValue)
    IsValidOrderPhone = False
    If Len(digits) = 11 And Left(digits, 1) = "1" Then
        IsValidOrderPhone = True
    ElseIf Len(digits) >= 7 And Len(digits) <= 12 Then
        ' allow simple landline digits
        IsValidOrderPhone = True
    End If
End Function

Function IsValidOrderAmount(textValue)
    Dim t
    t = Replace(Replace(Replace(Replace(CStr(textValue), ",", ""), "￥", ""), "¥", ""), "元", "")
    t = Trim(t)
    IsValidOrderAmount = False
    If Len(t) = 0 Then Exit Function
    If IsNumeric(t) Then
        If CDbl(t) >= 0 Then IsValidOrderAmount = True
    End If
End Function

Function WriteOrderMappingSheet(sheetObj, ByRef roleLabels(), ByRef mapLabels(), ByRef mapCols(), roleCount)
    On Error Resume Next
    Dim i
    WriteOrderMappingSheet = False
    sheetObj.Cells(1, 1).Value = "角色"
    sheetObj.Cells(1, 2).Value = "映射表头"
    sheetObj.Cells(1, 3).Value = "列号"
    sheetObj.Cells(1, 4).Value = "状态"
    For i = 1 To roleCount
        sheetObj.Cells(i + 1, 1).Value = roleLabels(i)
        If mapCols(i) > 0 Then
            sheetObj.Cells(i + 1, 2).Value = mapLabels(i)
            sheetObj.Cells(i + 1, 3).Value = mapCols(i)
            sheetObj.Cells(i + 1, 4).Value = "已确认"
        Else
            sheetObj.Cells(i + 1, 2).Value = ""
            sheetObj.Cells(i + 1, 3).Value = ""
            sheetObj.Cells(i + 1, 4).Value = "已跳过"
        End If
    Next
    sheetObj.Rows(1).Font.Bold = True
    sheetObj.Columns.AutoFit
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    WriteOrderMappingSheet = True
End Function

Function WriteOrderIssueSheet(sheetObj, ByRef issueRows(), ByRef issueFields(), ByRef issueProblems(), ByRef issueValues(), issueCount)
    On Error Resume Next
    Dim i
    WriteOrderIssueSheet = False
    sheetObj.Cells(1, 1).Value = "来源行"
    sheetObj.Cells(1, 2).Value = "字段"
    sheetObj.Cells(1, 3).Value = "问题"
    sheetObj.Cells(1, 4).Value = "原值"
    sheetObj.Cells(1, 5).Value = "建议"
    If issueCount = 0 Then
        sheetObj.Cells(2, 1).Value = "(无)"
        sheetObj.Cells(2, 2).Value = "-"
        sheetObj.Cells(2, 3).Value = "未发现问题"
        sheetObj.Cells(2, 4).Value = ""
        sheetObj.Cells(2, 5).Value = "可继续小样本验收后交付"
    Else
        For i = 0 To issueCount - 1
            sheetObj.Cells(i + 2, 1).Value = issueRows(i)
            sheetObj.Cells(i + 2, 2).Value = issueFields(i)
            sheetObj.Cells(i + 2, 3).Value = issueProblems(i)
            sheetObj.Cells(i + 2, 4).Value = issueValues(i)
            sheetObj.Cells(i + 2, 5).Value = SuggestForOrderIssue(issueProblems(i))
        Next
    End If
    sheetObj.Rows(1).Font.Bold = True
    sheetObj.Columns.AutoFit
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    WriteOrderIssueSheet = True
End Function

Function SuggestForOrderIssue(problemText)
    Dim t
    t = CStr(problemText)
    If InStr(1, t, "为空", vbTextCompare) > 0 Then
        SuggestForOrderIssue = "回源表补全后再质检"
    ElseIf InStr(1, t, "重复", vbTextCompare) > 0 Then
        SuggestForOrderIssue = "确认是否拆单/重导，保留唯一主键"
    ElseIf InStr(1, t, "电话", vbTextCompare) > 0 Then
        SuggestForOrderIssue = "规范为11位手机号或有效固话"
    ElseIf InStr(1, t, "金额", vbTextCompare) > 0 Then
        SuggestForOrderIssue = "去掉币符/中文后保留数值"
    Else
        SuggestForOrderIssue = "人工复核"
    End If
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
