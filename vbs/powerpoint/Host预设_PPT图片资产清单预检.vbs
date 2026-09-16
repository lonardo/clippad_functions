' 函数名: HostPptImageAssetInventoryPreflight
' 描述: 预检 PPT 图片资产后一键确认输出清单；可改导出 PNG，不改版式
' 适用应用: PowerPoint
' 搜索范围: 全文
' 搜索对象: 无
'
' 实现约定（供后续维护/生成使用）：
' - 先统计图片/链接图/隐藏页，默认动作为“预览”；输入“清单”才写 CSV，输入“导出”=清单+best-effort PNG。
' - 默认 includeHidden=false；可用“清单|隐藏”“导出|隐藏”覆盖。选文件夹后写入，不修改演示文稿版式。
' - Shape.Export 为 best-effort（先枚举再 "PNG"）。RollbackWritePlan 不能撤销已写 CSV/PNG。

Option Explicit

Const msoPicture = 13
Const msoLinkedPicture = 11
Const msoPlaceholder = 14
Const ppPlaceholderBitmap = 9
Const ppPlaceholderPicture = 18
Const ppShapeFormatPNG = 2

Function Main()
    On Error Resume Next
    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = FailureJson("E_NO_PPT_APP", "未取得 PowerPoint 应用，请在 PowerPoint 中运行该预设")
        Exit Function
    End If
    Main = HostPptImageAssetInventoryPreflight(appObj)
End Function

Function HostPptImageAssetInventoryPreflight(appObj)
    On Error Resume Next

    Dim pres
    Set pres = appObj.ActivePresentation
    If Err.Number <> 0 Or TypeName(pres) = "Empty" Or TypeName(pres) = "Nothing" Then
        Err.Clear
        HostPptImageAssetInventoryPreflight = FailureJson("E_NO_PRESENTATION", "当前没有活动演示文稿")
        Exit Function
    End If

    Dim slideCount, imageCount, linkedCount, hiddenSlideCount, sampleText, manifestText
    Dim includeHiddenDefault
    includeHiddenDefault = False
    slideCount = pres.Slides.Count
    imageCount = 0
    linkedCount = 0
    hiddenSlideCount = 0
    sampleText = ""
    manifestText = ""
    ' temporary collect with default hidden policy for preview counts
    CollectImageAssets pres, includeHiddenDefault, imageCount, linkedCount, hiddenSlideCount, sampleText, manifestText

    Dim slideSummary, layoutSummary, previewText
    slideSummary = SafeHostText("GetPptSlideSummary")
    layoutSummary = SafeHostText("GetPptLayoutIssueSummary")
    previewText = "PPT 图片资产场景包（尚未写入）" & vbCrLf & _
        "幻灯片=" & CStr(slideCount) & "；图片=" & CStr(imageCount) & "；链接图=" & CStr(linkedCount) & _
        "；隐藏页=" & CStr(hiddenSlideCount) & "；默认含隐藏页=否" & vbCrLf & _
        "默认动作=预览（不改版式、不写文件）" & vbCrLf & _
        "样例=" & sampleText & vbCrLf & _
        "说明：导出模式会 best-effort 导出 PNG；失败时仍尽量保留清单" & vbCrLf & _
        "Host.GetPptSlideSummary: " & slideSummary & vbCrLf & _
        "Host.GetPptLayoutIssueSummary: " & layoutSummary & vbCrLf & vbCrLf & _
        "确认清单：输入 清单。" & vbCrLf & _
        "清单+导出图片：输入 导出" & vbCrLf & _
        "仅预览：直接确认或输入 预览" & vbCrLf & _
        "覆盖写法：清单|隐藏 / 导出|隐藏 / 预览"

    Dim confirmRaw, actionCode, includeHidden
    confirmRaw = SafePrompt(previewText, "预览")
    actionCode = "list"
    includeHidden = includeHiddenDefault
    If Not ParseImageConfirm(confirmRaw, actionCode, includeHidden) Then
        HostPptImageAssetInventoryPreflight = FailureJson("E_CONFIRM_REQUIRED", "未确认清单/导出/预览，已取消")
        Exit Function
    End If

    ' recollect if hidden policy changed
    If includeHidden <> includeHiddenDefault Then
        imageCount = 0
        linkedCount = 0
        hiddenSlideCount = 0
        sampleText = ""
        manifestText = ""
        CollectImageAssets pres, includeHidden, imageCount, linkedCount, hiddenSlideCount, sampleText, manifestText
        previewText = "PPT 图片资产场景包（尚未写入）" & vbCrLf & _
            "幻灯片=" & CStr(slideCount) & "；图片=" & CStr(imageCount) & "；链接图=" & CStr(linkedCount) & _
            "；隐藏页=" & CStr(hiddenSlideCount) & "；含隐藏页=" & YesNoLabel(includeHidden) & vbCrLf & _
            "动作=" & ImageActionLabel(actionCode) & vbCrLf & _
            "样例=" & sampleText
    Else
        ' ensure manifest matches default collect
        If Len(manifestText) = 0 Then
            CollectImageAssets pres, includeHidden, imageCount, linkedCount, hiddenSlideCount, sampleText, manifestText
        End If
    End If

    If actionCode = "preview" Then
        Host.WriteClipboard previewText
        SafeWriteLog previewText
        HostPptImageAssetInventoryPreflight = "{""ok"":true,""readonly"":true,""imageCount"":" & CStr(imageCount) & _
            ",""linkedCount"":" & CStr(linkedCount) & ",""includeHidden"":" & LCase(CStr(includeHidden)) & _
            ",""message"":""" & EscapeJson(previewText) & """}"
        Exit Function
    End If

    If imageCount <= 0 Then
        HostPptImageAssetInventoryPreflight = FailureJson("E_NO_IMAGES", "未发现可列出的图片资产")
        Exit Function
    End If

    Dim outFolder, folderTitle
    If actionCode = "export" Then
        folderTitle = "请选择图片导出与清单文件夹"
    Else
        folderTitle = "请选择图片清单输出文件夹"
    End If
    outFolder = Host.SelectFolder(folderTitle)
    If Err.Number <> 0 Or Len(Trim(CStr(outFolder))) = 0 Then
        Err.Clear
        HostPptImageAssetInventoryPreflight = FailureJson("E_OUTPUT_CANCELLED", "未选择输出文件夹")
        Exit Function
    End If

    Dim planId, planPreview, manifestPath, exportedCount
    planId = SafeBeginWritePlan()
    SafeRecordWrite "ppt_image_asset_scene", "{""action"":""" & EscapeJson(actionCode) & """,""images"":" & CStr(imageCount) & ",""includeHidden"":" & LCase(CStr(includeHidden)) & "}", "office.ppt.image.inventory"
    planPreview = SafePreviewWritePlan()

    manifestPath = Host.CombinePath(outFolder, "ppt_image_assets_manifest.csv")
    Host.WriteTextFile manifestPath, manifestText, True
    If Err.Number <> 0 Then
        Err.Clear
        CloseWritePlan planId
        HostPptImageAssetInventoryPreflight = FailureJson("E_MANIFEST_WRITE", "图片清单写入失败")
        Exit Function
    End If

    exportedCount = 0
    If actionCode = "export" Then
        exportedCount = ExportImageAssets(pres, outFolder, includeHidden)
    End If
    CloseWritePlan planId

    Dim summary
    summary = "PPT 图片资产场景包完成；动作=" & ImageActionLabel(actionCode) & _
        "；图片=" & CStr(imageCount) & "；导出成功=" & CStr(exportedCount) & _
        "；含隐藏页=" & YesNoLabel(includeHidden) & "；清单=" & manifestPath & _
        "；版式未改" & vbCrLf & "Preview: " & planPreview
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostPptImageAssetInventoryPreflight = "{""ok"":true,""sourceUnchanged"":true,""outputWritten"":true,""action"":""" & EscapeJson(actionCode) & """,""imageCount"":" & CStr(imageCount) & _
        ",""exportedCount"":" & CStr(exportedCount) & ",""manifestPath"":""" & EscapeJson(manifestPath) & """,""includeHidden"":" & LCase(CStr(includeHidden)) & _
        ",""writePlanPreview"":""" & EscapeJson(planPreview) & """,""message"":""" & EscapeJson(summary) & """}"
End Function

Function ParseImageConfirm(rawText, ByRef actionCode, ByRef includeHidden)
    Dim text, parts, head, overrideText, i, token
    actionCode = ""
    includeHidden = False
    text = Trim(CStr(rawText))
    If Len(text) = 0 Then
        ParseImageConfirm = False
        Exit Function
    End If

    overrideText = ""
    If InStr(1, text, "|", vbBinaryCompare) > 0 Then
        parts = Split(text, "|")
        head = Trim(CStr(parts(0)))
        For i = 1 To UBound(parts)
            token = LCase(Trim(CStr(parts(i))))
            If token = "隐藏" Or token = "hidden" Or token = "含隐藏" Or token = "includehidden" Then
                includeHidden = True
            ElseIf token = "不隐藏" Or token = "nohidden" Then
                includeHidden = False
            ElseIf Len(overrideText) = 0 Then
                overrideText = token
            End If
        Next
    Else
        head = text
    End If

    If StrComp(head, "预览", vbTextCompare) = 0 Or StrComp(head, "preview", vbTextCompare) = 0 Then
        actionCode = "preview"
        ParseImageConfirm = True
        Exit Function
    End If

    If StrComp(head, "清单", vbTextCompare) = 0 Or StrComp(head, "仅清单", vbTextCompare) = 0 Or StrComp(head, "list", vbTextCompare) = 0 Then
        actionCode = "list"
        ParseImageConfirm = True
        Exit Function
    End If

    If StrComp(head, "导出", vbTextCompare) = 0 Or StrComp(head, "export", vbTextCompare) = 0 Or StrComp(head, "清单+导出图片", vbTextCompare) = 0 Then
        actionCode = "export"
        ParseImageConfirm = True
        Exit Function
    End If

    ParseImageConfirm = False
End Function

Function ImageActionLabel(code)
    If code = "list" Then
        ImageActionLabel = "仅清单"
    ElseIf code = "export" Then
        ImageActionLabel = "清单+导出图片"
    ElseIf code = "preview" Then
        ImageActionLabel = "仅预览"
    Else
        ImageActionLabel = CStr(code)
    End If
End Function

Function IsPictureShape(shape)
    On Error Resume Next
    Dim t, ph
    t = shape.Type
    If t = msoPicture Or t = msoLinkedPicture Then
        IsPictureShape = True
        Exit Function
    End If
    If t = msoPlaceholder Then
        ph = shape.PlaceholderFormat.Type
        If Err.Number = 0 Then
            If ph = ppPlaceholderBitmap Or ph = ppPlaceholderPicture Then
                IsPictureShape = True
                Exit Function
            End If
        End If
        Err.Clear
    End If
    IsPictureShape = False
End Function

Sub CollectImageAssets(pres, includeHidden, ByRef imageCount, ByRef linkedCount, ByRef hiddenSlideCount, ByRef sampleText, ByRef manifestText)
    On Error Resume Next
    Dim slide, shape, hs, sn, nm, tp, w, h, line, sampleCount
    imageCount = 0
    linkedCount = 0
    hiddenSlideCount = 0
    sampleText = ""
    sampleCount = 0
    manifestText = "slide,shape,name,type,width,height,hiddenSlide" & vbCrLf

    For Each slide In pres.Slides
        hs = False
        hs = slide.SlideShowTransition.Hidden
        If Err.Number <> 0 Then
            hs = False
            Err.Clear
        End If
        If hs Then hiddenSlideCount = hiddenSlideCount + 1
        If (Not hs) Or includeHidden Then
            For Each shape In slide.Shapes
                If IsPictureShape(shape) Then
                    imageCount = imageCount + 1
                    sn = slide.SlideIndex
                    nm = ""
                    nm = CStr(shape.Name)
                    tp = shape.Type
                    If tp = msoLinkedPicture Then linkedCount = linkedCount + 1
                    w = 0
                    h = 0
                    w = shape.Width
                    h = shape.Height
                    line = CStr(sn) & "," & CStr(shape.Id) & ",""" & CsvSafe(nm) & """," & CStr(tp) & "," & CStr(w) & "," & CStr(h) & "," & LCase(CStr(hs))
                    manifestText = manifestText & line & vbCrLf
                    If sampleCount < 6 Then
                        If sampleCount > 0 Then sampleText = sampleText & " | "
                        sampleText = sampleText & "S" & CStr(sn) & ":" & Left(nm, 20)
                        sampleCount = sampleCount + 1
                    End If
                End If
                Err.Clear
            Next
        End If
        Err.Clear
    Next
    If Len(sampleText) = 0 Then sampleText = "无"
End Sub

Function ExportImageAssets(pres, outFolder, includeHidden)
    On Error Resume Next
    Dim slide, shape, hs, sn, nm, outPath, n, ok
    n = 0
    For Each slide In pres.Slides
        hs = False
        hs = slide.SlideShowTransition.Hidden
        If Err.Number <> 0 Then
            hs = False
            Err.Clear
        End If
        If (Not hs) Or includeHidden Then
            For Each shape In slide.Shapes
                If IsPictureShape(shape) Then
                    sn = slide.SlideIndex
                    nm = Host.SanitizeFileName(CStr(shape.Name), "pic")
                    outPath = Host.CombinePath(outFolder, "Slide" & Right("000" & CStr(sn), 3) & "_" & nm & "_" & CStr(shape.Id) & ".png")
                    ok = False
                    Err.Clear
                    shape.Export outPath, ppShapeFormatPNG
                    If Err.Number = 0 Then
                        ok = True
                    Else
                        Err.Clear
                        shape.Export outPath, "PNG"
                        If Err.Number = 0 Then ok = True
                    End If
                    If ok Then n = n + 1
                    Err.Clear
                End If
                Err.Clear
            Next
        End If
        Err.Clear
    Next
    ExportImageAssets = n
End Function

Function CsvSafe(text)
    Dim t
    t = Replace(CStr(text), """", """""")
    t = Replace(t, ",", "；")
    t = Replace(t, vbCrLf, " ")
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    CsvSafe = t
End Function

Function SafeHostText(methodName)
    On Error Resume Next
    If methodName = "GetPptSlideSummary" Then
        SafeHostText = Host.GetPptSlideSummary()
    ElseIf methodName = "GetPptLayoutIssueSummary" Then
        SafeHostText = Host.GetPptLayoutIssueSummary()
    ElseIf methodName = "GetPptContextInfo" Then
        SafeHostText = Host.GetPptContextInfo()
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

Function YesNoLabel(flag)
    If flag Then
        YesNoLabel = "是"
    Else
        YesNoLabel = "否"
    End If
End Function
