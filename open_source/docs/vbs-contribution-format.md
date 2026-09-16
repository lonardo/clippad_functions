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

`搜索范围:` 和 `搜索对象:` 会进入 VBS 命令索引，帮助命令搜索和推荐判断脚本是否适合当前任务。例如：

```text
' 搜索范围: 选区
' 搜索对象: 文本
```

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

## 7. 提交前清单

- [ ] 文件是 UTF-8 BOM，扩展名为 `.vbs`；
- [ ] 头部连续注释位于 `Option Explicit` 和所有代码之前；
- [ ] 有 `函数名`、`描述`/`Description`、`适用应用`；
- [ ] 有无参数 `Function Main()`；
- [ ] 所有 `Host.*` 方法存在于公共 Host API 契约；
- [ ] 不使用公共包禁止的旧式 COM/脚本入口；
- [ ] 写入操作有预览、确认、备份或回滚说明；
- [ ] 已运行单文件校验和整包校验；
- [ ] 已在 staging 的真实 Word/Excel/PowerPoint 中测试。
