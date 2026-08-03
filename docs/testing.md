# Public Pack Testing

## 1. 整包静态校验

在公共仓库根目录运行：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass `
  -File .\tools\validate_public_pack.ps1
```

成功时会报告 VBS、workflow、Host 方法和 JSON 数量。它会检查：

- VBS 数量、UTF-8 BOM 和头部元数据；
- 无参数 `Function Main()` 入口；
- Host API JSON 方法数量和 `Host.GetSelection()` 语义；
- 工作流、Schema 和示例 JSON 可解析；
- 公共 VBS 中不存在危险的旧式 COM/脚本入口；
- 排除目录没有被带入公共包。

## 2. 单文件校验

新写一个脚本时，不必先把它复制进公共包，可以先运行：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass `
  -File .\tools\validate_public_pack.ps1 `
  -ScriptPath C:\path\to\MyFunction.vbs
```

单文件校验至少确认：

- 文件是 `.vbs` 且为 UTF-8 BOM；
- 顶部连续单引号注释中存在 `函数名`、`描述`/`Description`、`适用应用`；
- 文件包含无参数 `Function Main()`；
- 没有公共包禁止的 `CreateObject`、`GetObject`、`WScript.Shell`、`FileSystemObject` 等入口。

## 3. 独立 cscript 语法烟测

`cscript` 不能自动提供插件注入的 `Host` 对象，所以它不能证明真实 Office 操作成功；它只适合检查转换后的脚本能被 Windows Script Host 解析。对 UTF-8 BOM 文件先生成 UTF-16LE 临时副本：

```powershell
$source = (Resolve-Path .\vbs\word\Host预设_Word更新目录.vbs).Path
$temp = Join-Path $env:TEMP "officeaddin-vbs-smoke.vbs"
$text = [IO.File]::ReadAllText($source, [Text.Encoding]::UTF8)
[IO.File]::WriteAllText($temp, $text, [Text.Encoding]::Unicode)
cscript //nologo $temp
Remove-Item -LiteralPath $temp -Force
```

如果脚本依赖 `Host.*`，即使语法正确，也必须继续做 staging Host 测试。

## 4. 安装到 VBS Scripts 并确认被插件读取

1. 把脚本直接复制到 `<OfficeAddin 安装目录>\VBA Script\`；不要只放在仓库的 `vbs\word` 子目录，也不要保留多层子目录。
2. 确认文件名唯一，并且没有把 `vbs_command_index.json` 当成源文件提交。
3. 重启 Word、Excel 或 PowerPoint，打开插件命令目录。
4. 若仍未出现，用随安装包提供的代理强制刷新：

   ```powershell
   & "<OfficeAddin 安装目录>\VbsProxy.exe" `
     -index `
     -scriptdir "<OfficeAddin 安装目录>\VBA Script" `
     -force `
     -output "$env:TEMP\vbs-index-result.json"
   ```

5. 检查输出 JSON 中的 `ok`、`scriptDir`、`indexPath` 和 `message`，再重新打开命令搜索。

## 5. 真实 Office 测试

1. 在 Word、Excel、PowerPoint 中分别打开真实但可恢复的测试文档；
2. 覆盖空文档、空选区、复杂选区和不满足前置条件的失败路径；
3. 对写入脚本检查预览、确认、备份、提交、取消和回滚；
4. 检查中文界面和英文界面的显示名；
5. 检查 Excel 多单元格数组、PowerPoint 形状和 Word 复杂表格等宿主差异。

公共仓库静态检查不能替代真实 Office 测试，尤其不能替代 COM 兼容性和写入回滚测试。
