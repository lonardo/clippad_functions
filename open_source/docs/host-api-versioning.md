# Host API Versioning

## 版本来源

公共 Host 契约文件是 host/host-api-v1.json。它只公开方法名、分类、签名和行为说明，不公开私有 C++ 实现、DISPID 细节或部署配置。

## 版本规则

- 主版本：删除方法、改变参数/返回语义、改变安全边界时递增；
- 次版本：增加向后兼容的方法、字段或能力说明时递增；
- 补丁版本：修正文档、示例和非行为性错误时递增。

脚本可以声明 minimumHostApi。公共仓库中的脚本不得依赖比声明版本更新的 API。

## 方法生命周期

1. 先加入 Host 方法单一事实源和实现；
2. 生成或更新公共 JSON 契约；
3. 增加 fake Host 和 staging Office smoke；
4. 更新提示词、示例和兼容性说明；
5. 至少保留一个版本周期的弃用说明；
6. 在 Release 中记录新增、变更和弃用方法。

## 选区方法

Host.GetSelection() 是稳定的纯文本读取方法，不负责交互，也不返回 COM 对象。需要结构化数据或对象操作时，脚本应使用 GetSelectionInfo() 或 GetSelectionObject()。
