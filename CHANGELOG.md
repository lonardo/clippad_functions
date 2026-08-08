# Changelog

## community-v0.1.3 - 2026-08-08

- 同步 `Host预设_Office导出计划预检.vbs` 为源仓库人话只读预检版（不改公开脚本数量，仍为 105 + 2 示例）。
- 明确 TablePipeline / 表管道样例暂不进入公开包：当前 bootstrap 仍依赖 `CreateObject`/`FileSystemObject`/`ADODB.Stream`，不符合 Host-first 公开安全边界。
- 校正公开文档中的分类计数与版本审计备注。

## community-v0.1.2 - 2026-08-07

- 同步主包中的 105 个 Host-first VBS 公共脚本，修正公共包与安装包脚本内容漂移。
- 新增 Excel 场景包公共镜像，并更新导出清单及校验数量（VBS=105，含示例共 107）。
- 保持公共包只包含脚本、Host API、示例、提示词、Schema 和贡献文档，不包含插件二进制及部署材料。

## community-v0.1.1 - 2026-08-04

- 补充 README、官网/论坛/支持入口和普通用户快速上手流程。
- 明确 `VBA Script` 一级目录扫描规则、运行时索引刷新和仓库分类目录的区别。
- 明确头部注释字段、`Function Main()` 入口、英文界面显示名和 Host-first COM 边界。
- 增加单文件 VBS 校验入口、生成提示词示例和从生成到 staging Office 验证的步骤。

## community-v0.1.0-draft - 2026-08-02

- 导出 94 个 Host-first VBS 预设，并按 common、Word、Excel、PowerPoint 分类。
- 导出 116 个 Host 方法的 v1 JSON 契约。
- 明确 Host.GetSelection() 可用、只读、返回当前选区纯文本且不触发交互。
- 增加 VBS 嵌入、Office 2010+ 兼容性、使用、安装、批量生成和性能披露文档。
- 增加 Apache-2.0 许可证、贡献规范、安全政策、Issue 模板和 Release 模板。
- 排除软著、ICP备案、专利、市场申报、后端部署、日志、签名和生产配置。
