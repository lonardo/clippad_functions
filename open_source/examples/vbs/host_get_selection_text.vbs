' 函数名: HostGetSelectionText
' 描述: 读取当前选区文本，不弹出选择对话框 / Read current selection text without opening a selection dialog.
' 适用应用: Word/Excel/PowerPoint
' 搜索范围: 选区
' 搜索对象: 文本
' License: Apache-2.0

Function Main()
    On Error Resume Next
    Err.Clear

    Dim selectedText
    selectedText = Host.GetSelection()
    If Err.Number <> 0 Then
        Main = "{""ok"":false,""code"":""E_GET_SELECTION"",""message"":""无法读取当前选区 / Unable to read current selection""}"
        Exit Function
    End If

    If Len(selectedText) = 0 Then
        Main = "{""ok"":false,""code"":""E_EMPTY_SELECTION"",""message"":""当前没有可读取的选区文本 / No selection text is available""}"
        Exit Function
    End If

    Main = "{""ok"":true,""text"":""" & Host.JsonEscape(selectedText) & """}"
End Function
