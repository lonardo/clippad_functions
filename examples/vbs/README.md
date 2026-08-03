# VBS 生成示例

这里的两个 `.vbs` 文件是最小 Host-first 示例：

- [`host_context_snapshot.vbs`](host_context_snapshot.vbs)：读取当前 Office 上下文；
- [`host_get_selection_text.vbs`](host_get_selection_text.vbs)：读取当前选区文本。

## 给生成器的最小提示词

把目标、输入、输出和风险说清楚，并同时提供 [`../../host/host-api-v1.json`](../../host/host-api-v1.json)：

```text
请生成一个 OfficeAddin Host-first VBScript。

目标应用：Excel
任务：读取当前选区的文本，返回 JSON，不修改工作簿
输出：{"ok":true,"text":"..."}
风险：readonly

硬性要求：
1. 文件使用 UTF-8 BOM；
2. 文件最前面先写连续的单引号元数据注释：函数名、描述、适用应用、搜索范围、搜索对象；
3. 必须提供无参数 Function Main()，不要改名；
4. 只能使用 host-api-v1.json 中存在的 Host 方法；
5. 不使用 CreateObject、GetObject、WScript.Shell、FileSystemObject、ADODB.Stream、InputBox 或 MsgBox；
6. Main() 返回可解析 JSON，并提供稳定的失败 code 和中英文 message。
```

## 生成后的检查顺序

```powershell
# 在公共包根目录执行
powershell -NoProfile -ExecutionPolicy Bypass `
  -File .\tools\validate_public_pack.ps1 `
  -ScriptPath .\examples\vbs\host_get_selection_text.vbs
```

通过后，把脚本复制到客户端的：

```text
<OfficeAddin 安装目录>\VBA Script\
```

注意：仓库中的 `vbs/` 分类目录只用于源码组织；插件扫描的是 `VBA Script` 根目录的一级 `.vbs` 文件。复制后重启 Office；如果仍未显示，按照 [`../testing.md`](../testing.md) 的步骤使用 `VbsProxy.exe -index -force` 刷新索引。
