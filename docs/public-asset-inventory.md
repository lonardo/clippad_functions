# Public Asset Inventory

本文件用于决定哪些工程资产适合进入公共 GitHub 仓库。判断标准是：公共创作价值、来源和许可证清楚、不会暴露业务秘密、可以脱离私有后端理解和验证。

## 已进入公共包

| 资产 | 状态 | 说明 |
|---|---|---|
| VBA Script/Host预设_*.vbs | 已公开导出 | 94 个 Host-first 预设，已完成 Host 方法、安全 API 和 BOM 检查 |
| Plugin/WorkflowHostMethodList.inl | 只生成结果 | 只公开 host-api-v1.json，不公开 C++ 源文件 |
| workflows/examples 中的安全 JSON | 已公开导出 | 当前保留 4 个；移除了嵌入旧式 CreateObject/ADODB.Stream 的两个示例 |
| schemas/workflow-schema.json | 已公开导出 | 通用工作流结构契约 |
| schemas/template-schema.json | 已公开导出 | 通用模板定义结构契约，不授权任何危险实现 |
| schemas/format-profile-schema.json | 已公开导出 | 用户模板格式档案结构 |
| scene_presets/profiles/word/*.example.json | 已公开导出 | 无真实文档数据的结构示例，不代表法律或公文标准 |
| 提示词、Host 文档、贡献规范和校验工具 | 已公开导出 | 面向社区创作和 PR 审核 |

## 整理后适合公开

| 工程资产 | 当前处理 | 公开前工作 |
|---|---|---|
| templates/prompt_patterns/vbs_prompt_patterns.md | 作为提示词素材 | 去除私有 UI 同步说明，改成公共模式目录，并以 host-api-v1 为唯一方法依据 |
| templates/definitions/*.json | 作为结构素材 | 只公开定义 JSON；对应 VBS 模板必须全部改成 Host-first |
| templates/input、process、output、integration | 暂不复制 | 当前包含 CreateObject、GetObject、FSO、ADODB.Stream 或 WScript.Shell，需要逐个重写 |
| docs/Workflow_QuickStart.md、Workflow_Quick_Reference.md | 作为文档素材 | 剥离私有路径、内部实现和过期 API，合并到公共使用文档 |
| scripts/tests/validate_vbs_host_presets.ps1 | 不直接复制 | 它依赖私有仓库目录；已用 tools/validate_public_pack.ps1 提供公共版校验 |

## 暂不公开

- scene_presets/index.json、task_coverage_pilot_catalog.json 和 office intelligence 相关 Schema：包含仍在演进的产品路线、内部覆盖策略或预览契约。
- docs 下的架构计划、阶段总结、内部诊断报告、COM 内存/线程实现分析：它们依赖私有源码和内部研发上下文。
- Plugin、VbsProxy、VbsWorker、backend、setup、release、dist 和 deployment 配置：属于产品实现、部署或发布边界。
- test_*.vbs、diag_*.vbs 和内部 smoke 脚本：许多用于 COM 注册、私有调试或环境诊断，不适合作为公共示例。
- templates/settings*.ini、upgrade 配置、注册表文件、日志和任何本地环境文件。

## 永久排除

copyright_submission、icp_filing、patent、patent_submission、app_market_submission、backend、logs、签名材料、生产凭据、客户文件和私有 Git 历史。

公共仓库应从本目录重新建立，而不是把父仓库直接改成 public。
