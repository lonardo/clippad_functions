' 函数名: HostPptExportSlidesAsImages
' 描述: 将当前 PowerPoint 的每一页导出为 PNG 图片；适合预览图、长图排版和素材归档
' 适用应用: PowerPoint
' 搜索范围: 全文
' 搜索对象: 无

Option Explicit

Function Main()
    On Error Resume Next

    Dim appObj
    Set appObj = Host.GetApplication()
    If Err.Number <> 0 Or TypeName(appObj) = "Empty" Or TypeName(appObj) = "Nothing" Then
        Err.Clear
        Main = "{""ok"":false,""code"":""E_NO_PPT_APP"",""message"":""未取得 PowerPoint 应用，请在 PowerPoint 中运行该预设""}"
        Exit Function
    End If

    Main = HostPptExportSlidesAsImages(appObj)
End Function

Function HostPptExportSlidesAsImages(appObj)
    On Error Resume Next

    Dim pres, outputFolder, slide, outputPath, exported, planId
    Set pres = appObj.ActivePresentation
    If Err.Number <> 0 Or TypeName(pres) = "Empty" Or TypeName(pres) = "Nothing" Then
        Err.Clear
        HostPptExportSlidesAsImages = "{""ok"":false,""code"":""E_NO_PRESENTATION"",""message"":""当前没有活动演示文稿""}"
        Exit Function
    End If
    outputFolder = Host.SelectFolder("请选择 PNG 图片导出文件夹")
    If Len(Trim(outputFolder)) = 0 Then
        HostPptExportSlidesAsImages = "{""ok"":false,""code"":""E_OUTPUT_CANCELLED"",""message"":""未选择导出文件夹""}"
        Exit Function
    End If

    planId = SafeBeginWritePlan()
    SafeRecordWrite "ppt_export_slides_as_images", "{""slideCount"":" & CStr(pres.Slides.Count) & "}", "office.ppt.export"
    exported = 0
    For Each slide In pres.Slides
        outputPath = Host.CombinePath(outputFolder, "Slide_" & Right("000" & CStr(slide.SlideIndex), 3) & ".png")
        slide.Export outputPath, "PNG", 1600, 900
        If Err.Number = 0 Then exported = exported + 1
        Err.Clear
    Next
    SafeCloseWritePlan planId

    Dim summary
    summary = "PPT 幻灯片图片导出完成；导出=" & CStr(exported) & " 页；文件夹=" & outputFolder
    Host.WriteClipboard summary
    SafeWriteLog summary
    HostPptExportSlidesAsImages = "{""ok"":true,""message"":""" & EscapeJson(summary) & """,""exportedCount"":" & CStr(exported) & ",""outputFolder"":""" & EscapeJson(outputFolder) & """}"
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
    EscapeJson = text
End Function
