# Issues、Discussions 与 Release

## Issue 与 Discussion

官网入口：[论坛](https://clippad.vip/community) · [支持中心](https://clippad.vip/support) · [下载](https://clippad.vip/download)

- Issue：用于可复现的脚本缺陷、Host 兼容性问题、提示词回归和文档错误。
- Discussion：用于新脚本想法、Host API 设计、分类规则、提示词共同创作和性能经验。
- Security：涉及密钥、任意文件写入、文档破坏、权限绕过或供应链问题时，不要公开细节。

公共提交必须脱敏，不上传客户文档、真实业务数据、内部路径、部署配置、软著/ICP备案/专利资料。

## 每次公共版本

每个版本同步发布：

- Git tag，例如 community-v0.1.0；
- GitHub Release；
- VBS、Host API JSON、提示词和文档 ZIP；
- SHA256 校验值；
- CHANGELOG.md、兼容性说明和已知问题；
- 本次新增/修改脚本清单及风险等级。

先完成 staging 验证，再在获得明确确认后发布生产环境内容。安装包可以作为单独 Release asset 发布，但其产品使用许可与本公共包的 Apache-2.0 内容分开。
