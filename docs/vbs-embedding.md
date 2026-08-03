# VBS 如何嵌入 OfficeAddin

## 运行模型

OfficeAddin 启动命令目录时，会在运行时 `VBA Script` 目录中枚举一级 `*.vbs` 文件，读取头部元数据，然后把脚本名称、描述和适用应用交给命令搜索/推荐界面。它不会递归扫描仓库中的 `vbs\word`、`vbs\excel` 等分类目录。

被选中的脚本由 VBS 执行器运行，入口固定为无参数的 `Function Main()`。`Main()` 返回字符串，建议返回包含 `ok`、`code`、`message` 或 `data` 的 JSON。

```text
VBA Script\*.vbs
        │
        ├─ 读取连续的单引号头部注释
        ├─ 建立 vbs_command_index.json
        ├─ 中文界面：文件名去掉 Host预设_ 和应用前缀
        └─ 英文界面：优先显示“函数名:”的值
                │
                └─ 调用 Function Main() → Host.* → 返回 JSON 字符串
```

## Host 与 COM

脚本通过 `Host.GetApplication()`、`Host.GetDocument()`、`Host.GetSelection()`、`Host.GetSelectionObject()` 等 Host API 访问当前 Office 上下文；不要自行创建新的 Word、Excel 或 PowerPoint 进程。需要文件、剪贴板或写入事务时，也使用公共 Host API 契约，而不是 `CreateObject`、`FileSystemObject` 或 `WScript.Shell`。

`Function Main()` 是实际执行入口；头部的 `函数名:` 只负责显示名。不要把 `函数名:` 当成 VBScript 函数名，也不要把入口改成其他名称。

## 编码

- 公共 VBS 源文件使用 UTF-8 BOM；
- 头部注释必须出现在 `Option Explicit` 和所有代码之前；
- 生成器写出文件时要保留 BOM；
- 独立 `cscript` 对 UTF-8 VBS 的编码行为与插件不同，语法测试可使用 UTF-16LE 临时副本；
- 不要用 `ADODB.Stream`、`FileSystemObject` 或 `WScript.Shell` 绕过 Host 文件能力。

## 从生成到运行

1. 用 [`prompts/vbs-system-prompt.md`](../prompts/vbs-system-prompt.md) 和 [`host/host-api-v1.json`](../host/host-api-v1.json) 约束生成器；
2. 生成带标准头部和 `Function Main()` 的 UTF-8 BOM 文件；
3. 运行单文件校验：

   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File .\tools\validate_public_pack.ps1 -ScriptPath .\MyFunction.vbs
   ```

4. 将 `.vbs` 直接复制到 `<安装目录>\VBA Script\`；
5. 重启 Office；仍未出现时，用 `VbsProxy.exe -index -scriptdir "<安装目录>\VBA Script" -force` 强制更新索引；
6. 在真实 Word、Excel 或 PowerPoint 中验证上下文、预览、确认、提交和回滚。

## 失败时先看什么

- 文件不在 `VBA Script` 根目录：不会进入索引；
- 头部前出现 `Option Explicit` 或代码：元数据不会被完整读取；
- 缺少 BOM：中文头部可能乱码，公共校验会失败；
- 缺少 `Function Main()`：脚本不能按公共入口执行；
- `函数名:` 是中文：英文界面会显示中文；
- 文件修改后索引未更新：重启 Office 或强制刷新索引；
- `Host.*` 方法不存在：先对照 Host API v1，不要自行猜测 COM 方法。
