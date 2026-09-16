' 函数名: HostExcelNormalizePhoneNumbers
' 描述: 清理当前 Excel 区域中的电话号码格式，去除空格、横线和括号并尽量补齐 11 位手机号显示；对标名单清洗插件
' 适用应用: Excel
' 搜索范围: 当前范围
' 搜索对象: 无

Option Explicit

Function Main()
    On Error Resume Next
    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = "{""ok"":false,""code"":""E_NO_EXCEL_APP"",""message"":""未取得 Excel 应用，请在 Excel 中运行该预设""}"
        Exit Function
    End If
    Main = HostExcelNormalizePhoneNumbers(appObj)
End Function

Function HostExcelNormalizePhoneNumbers(appObj)
    On Error Resume Next
    Dim ws, selectedRange, dataRange
    Set ws = appObj.ActiveSheet
    Set selectedRange = appObj.Selection
    If Err.Number <> 0 Or TypeName(ws) = "Empty" Or TypeName(selectedRange) = "Empty" Then
        Err.Clear
        HostExcelNormalizePhoneNumbers = "{""ok"":false,""code"":""E_NO_SELECTION"",""message"":""请先选择 Excel 数据区域中的一个单元格""}"
        Exit Function
    End If
    Set dataRange = selectedRange.CurrentRegion
    If Err.Number <> 0 Or dataRange.Rows.Count < 1 Then
        Err.Clear
        HostExcelNormalizePhoneNumbers = "{""ok"":false,""code"":""E_EMPTY_RANGE"",""message"":""当前区域为空""}"
        Exit Function
    End If

    Dim rangeSummary, planId, previewJson, cell, changed, scanned, rawText, cleaned
    rangeSummary = SafeHostText("GetRangeSummary")
    planId = SafeBeginWritePlan()
    SafeRecordWrite "excel_normalize_phone_numbers", "{""scope"":""currentRegion""}", "office.excel.cell.cleanup"
    previewJson = SafePreviewWritePlan()

    changed = 0
    scanned = 0
    For Each cell In dataRange.Cells
        scanned = scanned + 1
        If Not IsError(cell.Value) Then
            rawText = CStr(cell.Text)
            If LooksLikePhoneText(rawText) Then
                cleaned = NormalizePhone(rawText)
                If IsNormalizedPhone(cleaned) Then
                    cell.NumberFormat = "@"
                    cell.Value = cleaned
                    If Err.Number = 0 Then changed = changed + 1
                    Err.Clear
                End If
            End If
        End If
    Next
    SafeCloseWritePlan planId

    Dim summary
    summary = "Excel 电话号码规范化完成；扫描=" & CStr(scanned) & "；修改=" & CStr(changed) & _
        "；区域=" & dataRange.Address(False, False) & vbCrLf & _
        "Host.GetRangeSummary: " & rangeSummary & vbCrLf & "Preview: " & previewJson
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostExcelNormalizePhoneNumbers = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""scanned"":" & CStr(scanned) & _
        ",""changed"":" & CStr(changed) & ",""planId"":""" & EscapeJson(planId) & """}"
End Function

Function NormalizePhone(value)
    Dim text, i, ch, digits
    text = CStr(value)
    text = Replace(text, ChrW(160), "")
    text = Replace(text, ChrW(12288), "")
    text = Replace(text, " ", "")
    text = Replace(text, "-", "")
    text = Replace(text, "—", "")
    text = Replace(text, "(", "")
    text = Replace(text, ")", "")
    text = Replace(text, "（", "")
    text = Replace(text, "）", "")
    text = Replace(text, "+", "")
    digits = ""
    For i = 1 To Len(text)
        ch = Mid(text, i, 1)
        If ch >= "0" And ch <= "9" Then digits = digits & ch
    Next
    If Left(digits, 2) = "86" And Len(digits) > 11 Then digits = Mid(digits, 3)
    NormalizePhone = digits
End Function


Function LooksLikePhoneText(value)
    Dim text, digits, i, ch, hasPhoneSep
    text = Trim(CStr(value))
    LooksLikePhoneText = False
    If Len(text) < 7 Or Len(text) > 24 Then Exit Function
    digits = 0
    hasPhoneSep = False
    For i = 1 To Len(text)
        ch = Mid(text, i, 1)
        If ch >= "0" And ch <= "9" Then
            digits = digits + 1
        ElseIf ch = "-" Or ch = " " Or ch = "(" Or ch = ")" Or ch = "（" Or ch = "）" Or ch = "+" Or ch = "—" Then
            hasPhoneSep = True
        ElseIf ch = ChrW(160) Or ch = ChrW(12288) Then
            hasPhoneSep = True
        Else
            Exit Function
        End If
    Next
    If digits < 7 Or digits > 15 Then Exit Function
    ' pure integer-like values without phone separators are skipped to avoid order/id false positives
    If (Not hasPhoneSep) And digits <> 11 Then Exit Function
    If digits = 11 And Left(NormalizePhone(text), 1) <> "1" Then Exit Function
    LooksLikePhoneText = True
End Function

Function IsNormalizedPhone(value)
    Dim text
    text = CStr(value)
    IsNormalizedPhone = False
    If Len(text) = 11 And Left(text, 1) = "1" Then
        IsNormalizedPhone = True
    ElseIf Len(text) >= 7 And Len(text) <= 12 Then
        IsNormalizedPhone = True
    End If
End Function

Function NormalizePhoneKeepCompare(value)
    NormalizePhoneKeepCompare = NormalizePhone(value)
End Function
Function SafeHostText(methodName)
    On Error Resume Next
    If methodName = "GetRangeSummary" Then
        SafeHostText = Host.GetRangeSummary()
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