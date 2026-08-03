# Public VBS Prompt Patterns

这些范式用于帮助大模型理解任务边界。它们不是硬编码 API 清单；生成代码时仍必须加载 host/host-api-v1.json。

## 通用字段

每个范式建议包含：

- id；
- category；
- targetApps；
- risk；
- inputContext；
- outputContract；
- requiredHostMethods；
- previewAndRollback；
- acceptanceCases。

## 推荐范式

### excel-cleanup

目标：清理当前 Excel 区域的空白、格式和文本规范问题。

约束：默认只预览；不得修改源表；确认后写入新工作表或新文件；使用 GetRangeSummary、GetHeaders、GetTypeStats 和写入计划。

### excel-validation

目标：检查必填、日期、金额、手机号、错误值和重复键。

约束：输出异常行、字段、原因和样例；只读诊断优先；不要把原始敏感字段写入日志或剪贴板。

### excel-mapping-copy

目标：根据用户确认的字段映射生成新工作表或交付文件。

约束：先展示源字段、目标字段、行数和输出路径；源表保持不变；避免覆盖已有文件。

### word-format-preview

目标：检查或规范 Word 标题、正文、表格、页眉页脚和页边距。

约束：先生成格式问题清单和写入计划；不要改写业务事实、金额、当事人或附件内容；确认后再应用格式。

### word-review-audit

目标：导出批注、修订、跟踪修订和格式风险摘要。

约束：只读模式默认开启；接受或拒绝修订必须单独确认并提供恢复路径。

### powerpoint-layout-audit

目标：检查标题缺失、越界对象、空文本框、字体混用和页脚页码风险。

约束：默认只读；母版、主题和业务内容不得自动迁移；格式写入需要预览和确认。

### office-file-batch

目标：对用户选择的文件夹执行有限数量的 Office 文件处理。

约束：使用 EnumerateFiles、GetBatchPlan、GetRunBudgetPlan 和 GetBatchFailureJson；设置最大数量、失败清单和取消路径；不得死循环。

### host-diagnostics

目标：输出 HostInfo、ContextInfo、DocumentInfo、SelectionInfo 和应用专属摘要。

约束：只读，不上传文档；摘要应截断，避免把整份业务文档写入日志。

### selection-text

目标：读取当前选区纯文本。

约束：Host.GetSelection() 不触发交互；如果需要对象或结构化元数据，改用 GetSelectionObject() 或 GetSelectionInfo()。

## 生成后的检查

大模型必须说明目标应用、风险、所需 Host 方法、预览/确认/回滚方式，并返回一个可执行的最小测试案例。
