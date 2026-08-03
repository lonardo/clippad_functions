# FAQ

## 这是完整的 OfficeAddin 源码吗？

不是。本仓库只公开社区脚本、提示词、Host 契约、Schema、示例和文档。产品 DLL、后端和商业发布材料仍然独立管理。

## Office 版本支持到哪里？

当前公开口径是 Windows 桌面版 Office 2010 及以上，覆盖 Word、Excel 和 PowerPoint。

## Host.GetSelection() 会不会再次让用户选择？

不会。它只读取当前选区的可提取文本，不弹出选择框。需要用户重新选择范围时使用 Host.SelectRange()。

## 为什么直接 cscript 运行 UTF-8 VBS 可能失败？

Windows cscript 对 UTF-8 BOM 的直接读取存在环境差异。OfficeAddin 会按自己的脚本加载规则处理公共 VBS；独立语法烟测应使用 UTF-16LE 临时副本。

## 公共包是否包含安装程序？

公共源包不存放安装程序。正式安装包作为独立 Release 附件发布，并使用独立的产品许可和 SHA256 校验。

## 是否提供硬件最低配置？

暂不提供。没有经过统一基准测试的数字不应作为硬件承诺。

## 可以提交自己生成的 VBS 吗？

可以。请遵循 CONTRIBUTING.md、vbs-contribution-format.md，附带目标 Office、输入输出、风险、测试和许可证说明。

## 我把脚本放进仓库了，为什么插件看不到？

仓库里的 `vbs/word`、`vbs/excel` 等目录只是分类目录。插件只扫描客户端安装目录 `VBA Script` 下的一级 `.vbs` 文件。把脚本复制到正确目录后重启 Office；如果仍未出现，运行 `VbsProxy.exe -index -scriptdir "<安装目录>\VBA Script" -force` 刷新索引。

## 为什么英文界面的脚本名称不对？

英文界面优先读取脚本头部的 `函数名:`。它必须是稳定的英文显示名；`Function Main()` 才是执行入口，两者不是一回事。

## 去哪里反馈脚本问题？

可以在[官网论坛](https://clippad.vip/community)讨论使用和兼容性，也可以在 GitHub Issue 提交最小复现、版本和脱敏错误信息。
