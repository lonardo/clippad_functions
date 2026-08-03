# VBS Contribution Format

## 文件头

新增脚本建议使用以下连续注释头。现有预设至少保留函数名、描述和适用应用；新贡献应补齐风险、搜索范围和 Host 依赖。

~~~vbscript
' 函数名: Excel 清洗示例
' Description: Normalize selected Excel data without changing the source sheet.
' 适用应用: Excel
' 搜索范围: 当前范围
' 搜索对象: 无
' 风险等级: preview
' 写入模式: new_output
' 所需 Host: GetApplication,GetRangeSummary,BeginWritePlan,PreviewWritePlan,CommitWritePlan
' License: Apache-2.0
~~~

## 风险等级

- readonly：只读诊断、摘要、检查和导出计划；
- preview：生成影响范围或写入计划，但不提交变更；
- write：用户确认后修改文档或写入本地文件；
- destructive：删除、清空、接受修订、覆盖原文件等高风险操作。

写入和 destructive 脚本必须说明预览、确认、备份和回滚策略。

## 目录和命名

- common：跨 Office 的 Host、文件、剪贴板、计划和诊断；
- word：Word 专属；
- excel：Excel 专属；
- powerpoint：PowerPoint 专属；
- 文件名使用 Host预设_应用_动作.vbs；
- 一个脚本只解决一个主要任务，批量操作通过参数或明确的输入清单控制。

## 结果约定

脚本入口为 Function Main()。成功建议返回 ok=true；失败返回 ok=false、稳定 code 和中英文 message。不要把原始业务文档内容写入日志、Issue 或错误消息。
