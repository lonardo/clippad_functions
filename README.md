# OfficeAddin Community Pack

这是 OfficeAddin 的公共创作包，面向 GitHub 独立公开仓库使用。它包含 Host-first VBScript、示例工作流、Host API 契约、提示词和贡献/发布文档。

本目录是“公开白名单导出区”。发布 GitHub 时应将本目录复制到单独的公共仓库；不要把当前工程根目录直接作为公共仓库发布。

## 当前内容

- vbs/：94 个 Host-first 预设，按 common、word、excel、powerpoint 分类。
- examples/vbs/：可供大模型和贡献者参考的最小 Host 示例。
- examples/workflows/：4 个不含账号、密钥、部署地址和旧式文件 COM 操作的工作流 JSON 示例。
- schemas/：工作流、模板和用户模板格式的公共 JSON Schema。
- examples/scene-profiles/：不含真实文档数据的格式档案示例。
- host/host-api-v1.json：从 Host 方法单一事实源整理出的 116 个方法和选区语义。
- prompts/：生成、修复和批量变更 VBS 的系统提示词。
- tools/validate_public_pack.ps1：不依赖私有工程的公共包校验器。
- docs/：安装、嵌入、兼容性、批量生成、性能、资产盘点、Issue 和 Release 说明。
- .github/ISSUE_TEMPLATE/：公共脚本、提示词和文档的讨论入口模板。

## 兼容性口径

当前公开口径为：Windows 桌面版 Office 2010 及以上，覆盖 Word、Excel、PowerPoint。硬件最低参数暂不作公开承诺，性能说明见 docs/performance.md。

## Host.GetSelection()

Host.GetSelection() 是可用的只读 Host 方法。它返回当前 Office 选区的纯文本，不弹出选择对话框，也不返回 Range/Selection COM 对象。

- 只需要当前选区文本：使用 Host.GetSelection()。
- 需要结构化选区信息：使用 Host.GetSelectionInfo()。
- 需要操作 Office 选区对象：使用 Host.GetSelectionObject() 或 Host.GetApplication() 后取得对象。
- 需要用户重新选择范围：使用 Host.SelectRange()。

没有活动文档、活动 Host 或可提取的选区文本时，Host.GetSelection() 可能返回空字符串。Word 读取 Selection.Text；Excel 读取当前 Selection 的 Value2，复杂的多单元格数组不保证展开；PowerPoint 读取当前 ShapeRange 中可取得的文本。

## 明确不包含的内容

本公共包不包含以下内容：

- copyright_submission/、icp_filing/、patent/、patent_submission/；
- app_market_submission/、backend/、logs/、本地环境配置和部署资料；
- 签名私钥、内部证书材料、客户文件、诊断包和生产凭据；
- OfficeAddin 私有 C++ 实现、后端服务和商业运营资料；
- 未完成来源和许可证审查的旧版非 Host VBS 脚本。

## 许可证

本目录中的 VBS、示例代码、Host API 契约、提示词和原创文档默认采用 Apache-2.0，详见 LICENSE。第三方文件如日后加入，必须保留其原始许可证和 NOTICE。

## 贡献与发布

贡献要求见 CONTRIBUTING.md，安全问题见 SECURITY.md，Issue/Discussion 和 Release 规则见 docs/release-and-issues.md。公共版本发布前先在 staging 验证，生产发布需要单独确认。
