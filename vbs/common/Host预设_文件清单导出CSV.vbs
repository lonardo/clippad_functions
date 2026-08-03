' 函数名: HostExportFolderFileListCsv
' 描述: 选择文件夹，枚举文件并导出 CSV 清单；验证 Host.SelectFolder、EnumerateFiles、ResolveOutputPlan、CsvEscape、WriteTextFile
' 适用应用: Word|Excel|PowerPoint
' 搜索范围: 无
' 搜索对象: 无

Option Explicit

Function Main()
    On Error Resume Next

    Dim appObj
    Set appObj = Nothing
    Err.Clear
    Set appObj = Host.GetApplication()
    Err.Clear

    Main = HostExportFolderFileListCsv(appObj)
End Function

Function HostExportFolderFileListCsv(appObj)
    On Error Resume Next

    Dim folderPath
    folderPath = Trim(SafeSelectFolder("选择要生成文件清单的文件夹"))
    If Len(folderPath) = 0 Then
        HostExportFolderFileListCsv = "{""ok"":false,""code"":""E_FOLDER_CANCELLED"",""message"":""未选择文件夹""}"
        Exit Function
    End If

    Dim payload, csvText, count
    payload = Host.EnumerateFiles(folderPath, "*.*", False)
    csvText = BuildCsvFromPayload(payload, count)

    Dim outputJson, outputPath
    outputJson = Host.ResolveOutputPlan("office_file_list.csv", "avoid")
    outputPath = JsonStringValue(outputJson, "path")
    If Len(outputPath) = 0 Then outputPath = Host.ResolveTempPath("office_file_list.csv")

    Dim written
    written = Host.WriteTextFile(outputPath, csvText, True)

    Dim summary
    summary = "文件清单导出完成；文件数=" & CStr(count) & "；输出=" & outputPath
    Host.WriteClipboard summary
    SafeWriteLog summary

    HostExportFolderFileListCsv = "{""ok"":" & JsonBool(written) & ",""message"":""" & EscapeJson(summary) & _
        """,""count"":" & CStr(count) & ",""outputPath"":""" & EscapeJson(outputPath) & """}"

End Function

Function BuildCsvFromPayload(payload, ByRef count)
    Dim csv, filesPos, arrayStart, pos, nextPath
    count = 0
    csv = "index,path,name" & vbCrLf
    filesPos = InStr(1, payload, """files"":[", vbTextCompare)
    If filesPos <= 0 Then
        BuildCsvFromPayload = csv
        Exit Function
    End If
    arrayStart = InStr(filesPos, payload, "[")
    If arrayStart <= 0 Then
        BuildCsvFromPayload = csv
        Exit Function
    End If
    pos = arrayStart + 1
    Do
        nextPath = NextJsonString(payload, pos)
        If Len(nextPath) = 0 Then Exit Do
        nextPath = JsonUnescape(nextPath)
        count = count + 1
        csv = csv & CStr(count) & "," & Host.CsvEscape(nextPath) & "," & Host.CsvEscape(FileNameFromPath(nextPath)) & vbCrLf
    Loop
    BuildCsvFromPayload = csv
End Function

Function NextJsonString(text, ByRef pos)
    Dim q1, q2, ch, escaped, raw
    q1 = InStr(pos, text, """")
    If q1 <= 0 Then
        NextJsonString = ""
        Exit Function
    End If
    If InStr(pos, text, "]") > 0 And InStr(pos, text, "]") < q1 Then
        NextJsonString = ""
        Exit Function
    End If
    q2 = q1 + 1
    escaped = False
    Do While q2 <= Len(text)
        ch = Mid(text, q2, 1)
        If ch = "\" And Not escaped Then
            escaped = True
        ElseIf ch = """" And Not escaped Then
            Exit Do
        Else
            escaped = False
        End If
        q2 = q2 + 1
    Loop
    If q2 > Len(text) Then
        NextJsonString = ""
        Exit Function
    End If
    raw = Mid(text, q1 + 1, q2 - q1 - 1)
    pos = q2 + 1
    NextJsonString = raw
End Function

Function JsonStringValue(jsonText, key)
    Dim pattern, keyPos, colonPos, q1, q2
    pattern = """" & key & """"
    keyPos = InStr(1, jsonText, pattern, vbTextCompare)
    If keyPos <= 0 Then
        JsonStringValue = ""
        Exit Function
    End If
    colonPos = InStr(keyPos + Len(pattern), jsonText, ":")
    q1 = InStr(colonPos + 1, jsonText, """")
    q2 = InStr(q1 + 1, jsonText, """")
    If q1 <= 0 Or q2 <= q1 Then
        JsonStringValue = ""
    Else
        JsonStringValue = JsonUnescape(Mid(jsonText, q1 + 1, q2 - q1 - 1))
    End If
End Function

Function JsonUnescape(text)
    Dim t
    t = CStr(text)
    t = Replace(t, "\\", "\")
    t = Replace(t, "\""", """")
    t = Replace(t, "\r", vbCr)
    t = Replace(t, "\n", vbLf)
    t = Replace(t, "\t", vbTab)
    JsonUnescape = t
End Function

Function FileNameFromPath(pathText)
    Dim p
    p = InStrRev(pathText, "\")
    If p > 0 Then
        FileNameFromPath = Mid(pathText, p + 1)
    Else
        FileNameFromPath = pathText
    End If
End Function

Function SafeSelectFolder(prompt)
    On Error Resume Next
    SafeSelectFolder = Host.SelectFolder(prompt)
    If Err.Number <> 0 Then
        SafeSelectFolder = ""
        Err.Clear
    End If
End Function

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
