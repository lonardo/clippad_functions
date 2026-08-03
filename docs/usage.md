# Usage Guide

## 使用公共 VBS

1. 按 Word、Excel、PowerPoint 或 common 目录选择脚本。
2. 阅读脚本头部的函数名、描述和适用应用。
3. 在正确的 Office 应用和活动文档中运行。
4. 只读脚本可以直接检查结果；写入脚本必须先看预览和确认提示。
5. 对输出文件使用避免覆盖路径，对原文档保留备份。

## 选择正确的 Host 方法

- 当前选区文本：Host.GetSelection()
- 当前选区结构化摘要：Host.GetSelectionInfo()
- 当前选区对象：Host.GetSelectionObject()
- 新的用户选择：Host.SelectRange()
- 当前 Office 应用：Host.GetApplication()
- 文档/工作簿/演示文稿摘要：Host.GetContextInfo() 或应用专属 ContextInfo 方法

## 失败处理

脚本应返回 ok、code、message 或 data 字段。遇到空文档、未保存文档、没有选区、输出路径冲突或批处理超限时，应停止并给出可理解的提示，不应静默修改源文件。
