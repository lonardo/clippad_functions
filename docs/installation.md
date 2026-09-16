# Installation and First Run

本目录不直接存放安装包。正式安装包应从项目 GitHub Release 或官方发布渠道下载，并以对应 Release 的 SHA256 为准。

## 安装前

1. 确认使用 Windows 桌面版 Office 2010 及以上。
2. 关闭 Word、Excel 和 PowerPoint。
3. 只使用带版本号和校验值的正式安装包。
4. 不要从 Issue、Discussion 或个人网盘下载未知 DLL、VBS 或安装程序。

## 首次运行

1. 安装后重新打开目标 Office 应用。
2. 确认 OfficeAddin 功能区已出现。
3. 先运行只读诊断或预览脚本。
4. 确认输入文档、输出路径和影响范围后，再运行写入脚本。
5. 重要文档先备份；批处理前优先使用写入计划和避免覆盖策略。

## 安装社区 VBS

公共仓库中的脚本按 `vbs/common`、`vbs/word`、`vbs/excel`、`vbs/powerpoint` 分类，便于阅读；运行时请把要安装的 `.vbs` 文件直接复制到：

```text
<OfficeAddin 安装目录>\VBA Script\
```

插件当前只扫描这个目录下的一级 `.vbs` 文件，不会递归扫描子目录。复制前运行单文件校验，复制后重启 Office；仍未出现时，用随安装包提供的 `VbsProxy.exe -index -scriptdir "<安装目录>\VBA Script" -force` 重建索引。完整说明见 [`vbs-embedding.md`](vbs-embedding.md) 和 [`testing.md`](testing.md)。

## 问题排查

请在 Issue 中提供版本、Office 应用/版本、Windows 版本、32/64 位、脚本名和脱敏错误信息。不要上传私有文档、完整日志、令牌、内部路径或申报资料。

也可以在[官网论坛](https://clippad.vip/community)讨论脚本安装、兼容性和使用方法。
