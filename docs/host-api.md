# Host API

完整方法清单见 ../host/host-api-v1.json。该清单由 Host 方法单一事实源整理，提示词和贡献者应以它为准。

## Host.GetSelection() 的准确语义

Host.GetSelection() 已注册并有实际 COM 分发实现。它调用当前 Host 适配器的选区文本读取路径，返回 String/BSTR 形式的当前选区纯文本。

它的特点是：

- 可用；
- 只读；
- 不弹出选择框；
- 不返回 Range、Selection 或 Shape 等 COM 对象；
- 没有活动文档、选区或 Host 上下文时可能返回空字符串；
- Word 主要返回 Selection.Text，Excel 主要返回当前 Selection 的 Value2，PowerPoint 主要返回当前 ShapeRange 的文本；复杂对象和多单元格数组不保证被展开。

示例：

~~~vbscript
Function Main()
    On Error Resume Next
    Err.Clear

    Dim selectedText
    selectedText = Host.GetSelection()
    If Err.Number <> 0 Then
        Main = "{""ok"":false,""code"":""E_GET_SELECTION""}"
        Exit Function
    End If

    Main = "{""ok"":true,""text"":""" & Host.JsonEscape(selectedText) & """}"
End Function
~~~

## 选择方法的决策

| 需求 | 方法 |
|---|---|
| 读取当前选区纯文本 | Host.GetSelection() |
| 获取选区地址、类型和摘要 | Host.GetSelectionInfo() |
| 操作当前 Office 选区对象 | Host.GetSelectionObject() |
| 直接使用 Office Application 对象 | Host.GetApplication() |
| 让用户重新选择范围 | Host.SelectRange() |

如果用户已经选中了对象，不要连续调用 Host.SelectRange() 或重复读取并要求用户再次选择。需要结构化数据时优先使用 GetSelectionInfo()；需要对象级操作时使用 GetSelectionObject()。
