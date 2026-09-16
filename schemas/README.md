# Public Schemas

当前公开三个通用 Schema：

- workflow-schema.json：工作流结构；
- template-schema.json：模板定义结构；
- format-profile-schema.json：用户模板格式档案结构。

Schema 只约束数据形状，不授予执行权限，也不允许绕过 Host 安全边界。任何 VBS 实现仍需通过公共校验和真实 staging Office 测试。

scene-preset、task-coverage 和 office-intelligence 相关 Schema 暂不公开，待产品契约稳定且完成路线信息审查后再单独评估。
