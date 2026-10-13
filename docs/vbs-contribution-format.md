# VBS Contribution Format

这份文档同时描述“插件会读取的字段”和“贡献审核需要的字段”。两者不要混淆：头部注释是脚本的一部分，插件会在建立命令目录时读取它。

## 1. 文件位置和扫描规则

仓库为了便于维护，把脚本按 `vbs/common`、`vbs/word`、`vbs/excel`、`vbs/powerpoint` 分类；这是公共源码布局，不是插件运行时目录。

安装到客户端时，脚本必须直接放在：

```text
<OfficeAddin 安装目录>\VBA Script\MyFunction.vbs
```

当前索引器只枚举 `VBA Script\*.vbs`，不会递归读取 `VBA Script\word\*.vbs`。公共包安装器或手工安装步骤需要把分类目录中的文件展平复制到运行时目录。文件名（不含扩展名）是命令名，因此同一目录不能有重复的基础文件名。

## 2. 必须使用 UTF-8 BOM

所有公共 VBS 文件必须是 UTF-8 with BOM（字节开头为 `EF BB BF`）。BOM 不是可有可无的装饰：插件会按 UTF-8 优先读取头部，Windows 独立 `cscript` 的编码行为又不同。不要把已有脚本另存为 ANSI、GBK 或无 BOM UTF-8。

## 3. 程序会读取的头部字段

头部必须位于所有代码之前。插件从文件头开始读取连续的单引号注释，跳过空行；遇到第一行代码后停止解析。推荐使用 ASCII 半角冒号 `:`，每个字段一行：

```vbscript
' 函数名: HostExcelNormalizeText
' 描述: 清理当前 Excel 区域的文本空白 / Normalize text whitespace in the current Excel range
' 适用应用: Excel
' 搜索范围: 当前范围
' 搜索对象: 文本
' 风险等级: preview
' 写入模式: new_output
' 所需 Host: GetApplication,GetRangeSummary,BeginWritePlan,PreviewWritePlan,CommitWritePlan
' License: Apache-2.0

Option Explicit
```

### 必须字段

| 字段 | 程序行为 |
|---|---|
| `函数名:` | `en-US` 界面优先用它作为显示名；应使用稳定、可读的英文名称。它不是 VBScript 的执行入口。 |
| `描述:` 或 `Description:` | 命令描述；如果同时存在，当前解析器优先使用中文 `描述:`。建议把中英文放在同一行。 |
| `适用应用:` | 写入 `Word`、`Excel`、`PowerPoint`；多个应用使用 `\|`，例如 `Word\|Excel\|PowerPoint`。 |

### 强烈建议字段

`搜索范围:` 和 `搜索对象:` 会进入 VBS 命令索引。命令搜索和洞察推荐都用它们判断脚本是否适合当前任务。允许值是封闭的：

| 字段 | 允许值 |
|---|---|
| `搜索范围:` | `无`、`选区`、`当前范围`、`全文`、`指定范围` |
| `搜索对象:` | `无`、`文本`、`公式`、`图表` |

只有脚本离开该范围或对象就无法有意义执行时，才写限制值。处理当前区域、全文或用户稍后指定的范围时，`搜索对象` 写 `无`。拿不准时用 `全文` 和 `无`。不要因为名称里有「公式」或「图表」就写限制对象，也不要写 `表格`、`幻灯片`、`当前工作簿` 这类未登记的值。

### 兼容别名和审核字段

旧脚本可以使用 `标题:`/`Title:`、`说明:`、`App:`、`Scope:` 等兼容别名；新脚本不要依赖这些别名。`风险等级`、`写入模式`、`所需 Host`、`License` 目前主要用于人工审核、生成提示词和贡献说明，不会替代 Host 权限检查。

## 4. `Function Main()` 是固定入口

公共脚本必须包含无参数的：

```vbscript
Function Main()
    Main = "{""ok"":true}"
End Function
```

插件的 VBS 执行器按 `Main()` 入口调用脚本。不要把入口改成 `Run`、`Execute` 或带参数的 `Main(arg)`；可以在 `Main()` 外定义辅助 `Function`/`Sub`，但最终必须由 `Main()` 组织输入检查、Host 调用和结果返回。

`函数名:` 只是命令显示名，尤其影响英文界面；它不会改变执行入口。若 `函数名:` 写成中文，英文界面也会显示中文，所以建议使用英文标识符，例如 `HostWordFormatProfilePreflight`。

## 5. COM 和 Host 兼容边界

脚本通过插件注入的 `Host` 对象访问能力，例如：

```vbscript
Dim appObj
Set appObj = Host.GetApplication()
```

公共脚本不要自行创建 Office 进程，也不要使用以下入口绕过 Host：

- `CreateObject`、`GetObject`；
- `WScript.Shell`、`Shell.Application`、注册表或原生进程调用；
- `FileSystemObject`、`ADODB.Stream` 或直接删除/覆盖文件；
- `InputBox`、`MsgBox` 等阻塞式交互。

需要 Office 对象、文件、剪贴板或写入事务时，先查 [`host/host-api-v1.json`](../host/host-api-v1.json)。写入脚本应遵循“预检 → 影响范围预览 → 用户确认 → 提交 → 失败恢复”的顺序。

## 6. 返回值

`Main()` 返回字符串。推荐返回可解析 JSON：

```json
{"ok":true,"code":"OK","message":"done","data":{}}
```

失败也返回稳定的 `code` 和中英文 `message`，不要把完整客户文档、密钥或内部路径写入结果、日志、Issue 或 PR。

## 7. 洞察推荐规则

洞察推荐的是已经安装的受信任脚本，不是一段临时生成后直接运行的代码。推荐成立需要两层信息：

1. 脚本头部说明它适合什么应用、范围和对象。卡片上的名称和说明来自 `函数名` 与 `描述`。
2. 用户确认后的头部声明 `' @insightBinding {"schemaVersion":1,"policy":"local_trusted","bindings":[...]}` 说明它对应哪个已发布场景。拷贝这个 `.vbs` 会把声明一起带走。

`bindings` 里每条包含 `triggerId`、`apps`、`confirmMode`、`minRepeat`、`cooldownSec`、`enabled`。索引文件可以出现在 `VBA Script` 同级目录 `Insight Bindings\<文件名>.vbs.binding.json`，用来避免脚本目录里堆绑定文件。本地已确认声明不依赖这份索引。旧的同目录 `<文件名>.vbs.binding.json` 会在读取时迁出脚本目录；没有头部声明时，过期哈希的旧索引不会被推荐。

本地可声明的 `triggerId` 只有：

- `insight.format.font.bold`、`insight.format.font.color`、`insight.format.font.size`：仅 Word；
- `insight.format.repeat`、`insight.edit.repeat`、`insight.selection.find_replace`：Word 或 Excel。

`insight.document.save.repeat` 不能写进本地绑定。未知编号一律拒绝。`confirmMode` 只能是 `preview_first` 或 `required`。`auto` 会被拒绝。`preview_first` 先预演再写入，`required` 要求明确确认。`minRepeat` 至少为 1，`cooldownSec` 至少为 20。`apps` 使用 `word`、`excel`、`powerpoint`，并且必须落在该场景允许的应用里。

`@insightProposal` 可以出现在头部连续注释中，表示模型或作者提出的待审建议。宿主只会提取这一行。它永远不是已启用绑定。用户确认之后，宿主才写入 `@insightBinding`。模型不要写 `@insightBinding`，也不要在运行时登记触发器。无效声明不会回退到旧索引。

准备进入公共或市场分发的脚本使用 `market_restricted`：只能有一条绑定，冷却至少 300 秒，并且需要审核。本地受信任脚本使用 `local_trusted`，最多 16 条绑定。

不要在 VBS 正文里调用不存在的 `Host.ExcelListSheetCatalog`、`Host.ExcelInsertNamedChart` 或 `Host.ExcelCreateQueryDraft`，也不要发明新的 `triggerId` 来换取推荐。

## 8. 提交前清单

- [ ] 文件是 UTF-8 BOM，扩展名为 `.vbs`；
- [ ] 头部连续注释位于 `Option Explicit` 和所有代码之前；
- [ ] 有 `函数名`、`描述`/`Description`、`适用应用`；
- [ ] `搜索范围` 和 `搜索对象` 使用允许值；拿不准时是 `全文` 和 `无`；
- [ ] 若声明洞察绑定，写的是已确认的 `@insightBinding`，不是 `@insightProposal`；`triggerId` 来自已发布场景，`confirmMode` 不是 `auto`；
- [ ] 有无参数 `Function Main()`；
- [ ] 所有 `Host.*` 方法存在于公共 Host API 契约；
- [ ] 不使用公共包禁止的旧式 COM/脚本入口；
- [ ] 写入操作有预览、确认、备份或回滚说明；
- [ ] 已运行单文件校验和整包校验；
- [ ] 已在 staging 的真实 Word/Excel/PowerPoint 中测试。
