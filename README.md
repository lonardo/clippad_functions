# OfficeAddin Community Pack

这是 OfficeAddin 的公开创作包，面向希望编写、校验和分享 Office 自动化函数的用户。它包含 Host-first VBScript、Host API 契约、工作流示例、提示词、Schema、校验工具和贡献文档。

## 先看这里

- 官网：[https://clippad.vip/](https://clippad.vip/)
- 下载：[https://clippad.vip/download](https://clippad.vip/download)
- 官网论坛：[https://clippad.vip/community](https://clippad.vip/community)
- 支持中心：[https://clippad.vip/support](https://clippad.vip/support)

论坛适合讨论脚本效果、Office 版本兼容性和贡献建议；不要在公开 Issue 或论坛贴出客户文档、账号、Token、内部地址或完整诊断日志。

## 这个包能做什么

OfficeAddin 是运行在 Windows 桌面版 Word、Excel 和 PowerPoint 中的插件。它把重复的 Office 操作整理成可复用命令：

- Word：正文和标题排版、空白清理、目录/页码、表格和审阅痕迹处理；
- Excel：区域清洗、筛选、去重、公式审计、汇总和副本导出；
- PowerPoint：标题/页脚统一、对象对齐、图片整理、备注和大纲导出；
- 通用能力：读取当前上下文、剪贴板、预检、预览、写入计划和回滚；
- AI 辅助：根据用户任务生成或修改 Host-first VBS，但生成结果必须经过人工复核和校验。

本仓库是公共函数和契约仓库，不包含产品 DLL、后端、签名材料或商业运营资料。

## 五分钟安装一个 VBS

### 1. 获取代码

下载本仓库 ZIP，或克隆仓库。先从 `vbs/common`、`vbs/word`、`vbs/excel`、`vbs/powerpoint` 选择脚本。

### 2. 放到插件真正会扫描的目录

仓库中的 `vbs/` 是便于分类和审查的源码目录；插件运行时扫描的是安装目录下的一级目录：

```text
<OfficeAddin 安装目录>\VBA Script\*.vbs
```

请把需要安装的 `.vbs` 文件直接复制到 `VBA Script` 根目录，不要只复制 `vbs\word` 这个子目录，也不要保留多层子目录。当前命令目录索引只枚举 `VBA Script\*.vbs`，不会递归扫描 `vbs\word\*.vbs`。

文件名（不含 `.vbs`）是命令的稳定 ID。不要在同一目录放两个同名脚本；更新脚本时覆盖原文件即可。

### 3. 先校验

在公共包根目录运行整包校验：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass `
  -File .\tools\validate_public_pack.ps1
```

只检查新写的一个脚本：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass `
  -File .\tools\validate_public_pack.ps1 `
  -ScriptPath .\path\to\MyFunction.vbs
```

校验会检查 UTF-8 BOM、头部注释、`Function Main()`、危险旧式 COM/脚本入口和可解析的结果约定。它不能代替真实 Office 测试。

### 4. 刷新命令索引

通常重启 Word/Excel/PowerPoint 并重新打开命令目录即可。如果新文件仍未出现，可以用随安装包提供的 `VbsProxy.exe` 强制重建索引：

```powershell
& "<OfficeAddin 安装目录>\VbsProxy.exe" `
  -index `
  -scriptdir "<OfficeAddin 安装目录>\VBA Script" `
  -force `
  -output "$env:TEMP\vbs-index-result.json"
```

成功后会在 `VBA Script` 目录生成 `vbs_command_index.json`（或在用户配置目录使用后备索引）。这个索引是运行时缓存，不要提交到公共仓库。

### 5. 在 Office 中运行

关闭并重新打开目标 Office 应用，在插件的命令编辑器/命令搜索中查找脚本。中文界面通常显示去掉 `Host预设_` 和应用前缀后的文件名；英文界面优先显示头部的 `函数名:` 值。

## VBS 是怎样被插件执行的

流程可以简化为：

1. 插件定位 `VBA Script` 目录并枚举一级 `.vbs` 文件；
2. 读取文件开头连续的单引号注释，建立标题、描述、适用应用和搜索元数据；
3. 在命令目录中展示脚本；
4. 执行脚本的无参数 `Function Main()`；
5. 脚本通过插件注入的 `Host` 对象访问 Office、剪贴板、预检和写入计划；
6. `Main()` 返回字符串，建议是包含 `ok`、`code`、`message` 或 `data` 的 JSON。

Host-first 的重点是：脚本不要自己 `CreateObject("Word.Application")`，也不要通过 `WScript.Shell`、`FileSystemObject` 等绕过插件的安全边界。Office 对象、文件写入、剪贴板和事务能力应使用 `host/host-api-v1.json` 中列出的 Host 方法。

## 最小可用脚本

下面的示例同时满足插件解析和英文显示要求：

```vbscript
' 函数名: HostGetSelectionText
' 描述: 读取当前选区文本 / Read the current selection text
' 适用应用: Word|Excel|PowerPoint
' 搜索范围: 选区
' 搜索对象: 文本
' 风险等级: readonly
' 写入模式: none
' 所需 Host: GetSelection,JsonEscape
' License: Apache-2.0

Option Explicit

Function Main()
    On Error Resume Next
    Err.Clear

    Dim selectedText
    selectedText = Host.GetSelection()
    If Err.Number <> 0 Then
        Main = "{""ok"":false,""code"":""E_GET_SELECTION"",""message"":""Unable to read selection""}"
        Exit Function
    End If

    Main = "{""ok"":true,""text"":""" & Host.JsonEscape(selectedText) & """}"
End Function
```

`函数名`是显示名，不是入口函数；入口函数必须仍然叫 `Main`。更多生成提示词见 [`prompts/vbs-system-prompt.md`](prompts/vbs-system-prompt.md)，生成示例见 [`examples/vbs/README.md`](examples/vbs/README.md)。

## 头部注释标准

头部必须放在任何 `Option Explicit`、变量声明或可执行代码之前。允许空行，但从第一个非空行开始，必须连续使用单引号注释；一旦出现代码，后面的注释就不再作为元数据读取。

程序会读取的标准字段：

| 字段 | 是否建议 | 用途 |
|---|---:|---|
| `函数名:` | 必须 | 英文界面的显示名；建议使用稳定的英文标识符 |
| `描述:` 或 `Description:` | 必须 | 命令描述；`描述:` 优先于英文别名 |
| `适用应用:` | 必须 | `Word`、`Excel`、`PowerPoint`，多个值用 `\|` 分隔 |
| `搜索范围:` | 强烈建议 | 例如 `选区`、`全文`、`当前工作簿` |
| `搜索对象:` | 强烈建议 | 例如 `文本`、`表格`、`幻灯片` |

`标题:`/`Title:`、`说明:`、`App:`、`Scope:` 等是兼容别名，但公共贡献应优先使用上表中的标准写法。`风险等级`、`写入模式`、`所需 Host` 和 `License` 主要用于人工审查、生成器和文档，不要误以为它们会自动授予权限。

具体格式和字段约束见 [`docs/vbs-contribution-format.md`](docs/vbs-contribution-format.md)。

## 生成、贡献和验证

推荐顺序：

1. 给生成器提供 [`host/host-api-v1.json`](host/host-api-v1.json)；
2. 明确目标 Office、输入、输出、是否写入、风险和回滚方式；
3. 要求生成 UTF-8 BOM、头部注释和无参数 `Function Main()`；
4. 先运行单文件校验，再运行整包校验；
5. 在 Word、Excel、PowerPoint 的 staging 客户端中做真实测试；
6. 通过 Issue/PR 提交脚本和测试说明。

完整步骤见 [`docs/testing.md`](docs/testing.md) 和 [`CONTRIBUTING.md`](CONTRIBUTING.md)。

## 常见问题

- **脚本不显示**：确认扩展名是 `.vbs`，文件直接位于 `VBA Script` 根目录，并重启 Office 或强制重建索引。
- **中文乱码**：确认文件是 UTF-8 BOM，不要保存成 ANSI/GBK，也不要删除 BOM。
- **英文界面显示中文**：把头部 `函数名:` 的值改成英文；中文界面仍主要显示文件名。
- **脚本无法执行**：确认存在无参数 `Function Main()`，并检查调用的每个 `Host.*` 方法是否存在于 Host API 契约。
- **校验失败**：先按报错修复结构，再在真实 Office 中测试；`cscript` 只能做有限的语法/隔离测试，不能模拟真实 Host 和 COM 对象。

## 许可与安全

本目录中的原创 VBS、示例、Host 契约、提示词和文档默认采用 Apache-2.0，详见 [`LICENSE`](LICENSE)。只运行你信任的脚本；写入文档或文件的脚本应先预检、预览、确认，并尽量提供备份或回滚路径。
