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

## 已注册但容易被漏掉的 Excel 方法

这些方法已经在 Host 运行时注册。生成 VBS 时必须能调用它们。`GetHeaders()` 只返回当前选区首行，用 `|` 分隔，不能代替整表映射。

| 需求 | 调用 | 返回 |
|---|---|---|
| 按表头找列 | `Host.ExcelGetHeaderMap([address])` | JSON：`ok`、`sheet`、`address`、`headers`、`headerMap`（表头到从 1 开始的列号）。地址空则用当前选区。表头按该区域第 1 行 |
| 读取一张表 | `Host.ExcelReadTable([address][, maxRows])` | TableJson。默认最多 5000 行，硬上限 20000。字段见 `docs/TablePipeline_Contract.md` |
| 结果写到新表 | `Host.ExcelCreateOutputSheet(functionName[, outputMode][, sourceSheet][, create])` | `outputMode` 默认 `NewSheet`。`create` 为 `false`/`plan` 时只返回计划名，不建表 |
| 影响摘要 | `Host.ExcelWriteImpactSummary(summaryJson[, sheetName])` | 把摘要行写到指定表 |
| 样本行 | `Host.ExcelPreviewRowSamples([sampleCount][, columnsCsv])` | 数据行样本 JSON，不是整表 |
| 去重计划 | `Host.ExcelDedupePlan(keyColumnsCsv[, keepPolicy])` | 重复键计划，不直接删行 |
| 校验计划 | `Host.ExcelValidatePlan(rulesJson)` | 校验问题计划，不直接改值 |
| 复用已有脚本 | `Host.GetVbaScriptDirectory()` / `ResolveVbaScriptPath` / `ReadVbaScriptLib` | 只读 `VBA Script` 根目录。拒绝绝对路径和 `..` |

按表头定位时先调用 `ExcelGetHeaderMap`。目标表头不在 `headerMap` 里就停止，并在结果里写出缺的表头。不要把列号写死，也不要因为 `GetHeaders()` 为空就认定工作簿没有表头。

## 还没有 Host 方法、但写脚本时必须知道边界

下面三个名字没有注册。不要生成 `Host.ExcelListSheetCatalog`、`Host.ExcelInsertNamedChart`、`Host.ExcelCreateQueryDraft` 或任何未出现在 `host-api-v1.json` 里的 `Host.*`。

当前工作簿的表名和各表首行，用已经注入的 Application 只读遍历。限制列数，结果里只放表名、地址和表头，不把单元格明细放进日志或模型上下文。

```vbscript
Set app = Host.GetApplication()
Set wb = app.ActiveWorkbook
For Each ws In wb.Worksheets
    ' 只读 ws.Name 和 UsedRange 第 1 行，列数上限 32
Next
```

图表同样还没有专用 Host 方法。需要出图时，先 `ExcelGetHeaderMap` 或 `GetAddress` 确认区域，再用 `Host.ConfirmStep` 确认类型、系列方向、对象名和字体。确认后只用当前 Application 插入命名图表。VBScript 没有 Excel 枚举名，图表类型必须用数字，并在脚本里写 Const：柱状簇 `51`，折线 `4`，饼图 `5`。系列按行是 `1`，按列是 `2`。插入后立刻设置 `ChartObjects` 的 `Name`。失败或用户取消时按这个名字删除。不要 `CreateObject` 新 Excel，不要 `Quit`。源区域不改值和格式。

Power Query 也没有 Host 包装。只有用户明确选择创建查询，并且 `Workbook.Queries` 可用时，才 `Queries.Add`。默认不 `Refresh`。集合不存在时返回明确失败码，并改走已有清洗预设。不要把刷新、隐私级别和凭据对话框包成成功。
