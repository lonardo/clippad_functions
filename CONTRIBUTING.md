# Contributing

欢迎提交 VBS、Host 提示词、工作流示例和文档改进。

## 提交范围

- VBS 必须以 Host-first 方式访问 Office、文件、剪贴板和事务能力。
- 每个脚本必须有 Function Main()，并在头部注释中写明函数名、描述和适用应用。
- 需要写入文档或文件时，应先预检、预览、确认，并尽量提供备份或回滚路径。
- 生成器只能使用 host/host-api-v1.json 中列出的 Host 方法，不得自行臆造 API。
- Host.GetSelection() 可以使用，但只用于读取当前选区纯文本；不得用它模拟新的用户选择。

## 安全边界

公共 VBS 默认禁止直接使用 CreateObject、GetObject、WScript.Shell、Shell.Application、ADODB.Stream、FileSystemObject、InputBox、MsgBox、Shell、注册表和原生文件删除入口。需要相应能力时，请通过 Host API 表达。

提交内容不得包含客户文件、真实业务数据、账号、密钥、内部 URL、部署配置、软著/ICP备案/专利资料或市场申报材料。

## Pull Request 清单

- [ ] 文件来源和许可证明确，或为贡献者原创。
- [ ] VBS 使用 UTF-8 BOM，未改变现有脚本的编码。
- [ ] 已在测试环境验证，未直接发布生产环境。
- [ ] 变更说明包含目标 Office 应用、输入、输出、风险和回滚方式。
- [ ] 提示词变更附带至少一个输入、期望结果和失败案例。
