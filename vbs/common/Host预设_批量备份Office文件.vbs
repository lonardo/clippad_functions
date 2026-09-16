' 函数名: HostBatchBackupOfficeFiles
' 描述: 选择源目录和目标目录，枚举常见 Office 文件并复制到时间戳备份目录；验证 Host.SelectFolder、EnumerateFiles、GetBatchPlan、CopyFile、WriteTextFile
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

    Main = HostBatchBackupOfficeFiles(appObj)
End Function

Function HostBatchBackupOfficeFiles(appObj)
    On Error Resume Next

    Dim sourceFolder, targetRoot
    sourceFolder = Trim(SafeSelectFolder("选择需要备份 Office 文件的源文件夹"))
    If Len(sourceFolder) = 0 Then
        HostBatchBackupOfficeFiles = "{""ok"":false,""code"":""E_SOURCE_CANCELLED"",""message"":""未选择源文件夹""}"
        Exit Function
    End If

    targetRoot = Trim(SafeSelectFolder("选择备份输出文件夹"))
    If Len(targetRoot) = 0 Then
        HostBatchBackupOfficeFiles = "{""ok"":false,""code"":""E_TARGET_CANCELLED"",""message"":""未选择备份输出文件夹""}"
        Exit Function
    End If

    Dim backupFolder
    backupFolder = Host.CombinePath(targetRoot, "OfficeBackup_" & TimestampText())
    If Not Host.CreateFolder(backupFolder) Then
        HostBatchBackupOfficeFiles = "{""ok"":false,""code"":""E_CREATE_FOLDER"",""message"":""无法创建备份目录"",""folder"":""" & EscapeJson(backupFolder) & """}"
        Exit Function
    End If

    Dim paths(), pathCount
    pathCount = 0
    CollectPattern sourceFolder, "*.doc", paths, pathCount
    CollectPattern sourceFolder, "*.docx", paths, pathCount
    CollectPattern sourceFolder, "*.xls", paths, pathCount
    CollectPattern sourceFolder, "*.xlsx", paths, pathCount
    CollectPattern sourceFolder, "*.ppt", paths, pathCount
    CollectPattern sourceFolder, "*.pptx", paths, pathCount

    Dim batchPlan
    batchPlan = SafeBatchPlan(pathCount)

    Dim copied, failed, i, src, dest, opPlan, ok
    Dim manifest
    copied = 0
    failed = 0
    manifest = "OfficeAddin Host batch backup" & vbCrLf & _
        "Source: " & sourceFolder & vbCrLf & _
        "Backup: " & backupFolder & vbCrLf & _
        "BatchPlan: " & batchPlan & vbCrLf

    For i = 0 To pathCount - 1
        src = paths(i)
        dest = Host.CombinePath(backupFolder, FileNameFromPath(src))
        opPlan = Host.GetFileOperationPlan("copy", src, dest, "avoid")
        ok = Host.CopyFile(src, dest, False)
        If ok Then
            copied = copied + 1
            manifest = manifest & "[OK] " & src & " => " & dest & vbCrLf
        Else
            failed = failed + 1
            manifest = manifest & "[FAIL] " & src & " => " & dest & " plan=" & opPlan & vbCrLf
        End If
    Next

    Dim manifestPath, copiedManifest
    manifestPath = Host.CombinePath(backupFolder, "backup_manifest.txt")
    copiedManifest = Host.WriteTextFile(manifestPath, manifest, True)
    Host.WriteClipboard manifest
    SafeWriteLog manifest

    HostBatchBackupOfficeFiles = "{""ok"":true,""message"":""Office 文件备份完成"",""source"":""" & EscapeJson(sourceFolder) & _
        """,""backupFolder"":""" & EscapeJson(backupFolder) & """,""found"":" & CStr(pathCount) & _
        ",""copied"":" & CStr(copied) & ",""failed"":" & CStr(failed) & _
        ",""manifestPath"":""" & EscapeJson(manifestPath) & """,""manifestWritten"":" & JsonBool(copiedManifest) & "}"

End Function

Sub CollectPattern(folderPath, pattern, ByRef paths, ByRef pathCount)
    On Error Resume Next
    Dim payload
    payload = Host.EnumerateFiles(folderPath, pattern, False)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    CollectFilesFromPayload payload, paths, pathCount
End Sub

Sub CollectFilesFromPayload(payload, ByRef paths, ByRef pathCount)
    Dim filesPos, arrayStart, pos, nextPath
    filesPos = InStr(1, payload, """files"":[", vbTextCompare)
    If filesPos <= 0 Then Exit Sub
    arrayStart = InStr(filesPos, payload, "[")
    If arrayStart <= 0 Then Exit Sub
    pos = arrayStart + 1
    Do
        nextPath = NextJsonString(payload, pos)
        If Len(nextPath) = 0 Then Exit Do
        ReDim Preserve paths(pathCount)
        paths(pathCount) = JsonUnescape(nextPath)
        pathCount = pathCount + 1
    Loop
End Sub

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
    Dim p1, p2, p
    p1 = InStrRev(pathText, "\")
    p2 = InStrRev(pathText, "/")
    p = p1
    If p2 > p Then p = p2
    If p > 0 Then
        FileNameFromPath = Mid(pathText, p + 1)
    Else
        FileNameFromPath = pathText
    End If
End Function

Function TimestampText()
    Dim d
    d = Now
    TimestampText = CStr(Year(d)) & Right("0" & CStr(Month(d)), 2) & Right("0" & CStr(Day(d)), 2) & "_" & _
        Right("0" & CStr(Hour(d)), 2) & Right("0" & CStr(Minute(d)), 2) & Right("0" & CStr(Second(d)), 2)
End Function

Function SafeSelectFolder(prompt)
    On Error Resume Next
    SafeSelectFolder = Host.SelectFolder(prompt)
    If Err.Number <> 0 Then
        SafeSelectFolder = ""
        Err.Clear
    End If
End Function

Function SafeBatchPlan(itemCount)
    On Error Resume Next
    SafeBatchPlan = Host.GetBatchPlan(itemCount, 50, 1000, "copy")
    If Err.Number <> 0 Then
        SafeBatchPlan = ""
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
