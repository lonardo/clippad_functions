# Public Pack Testing

## 静态检查

在公共仓库根目录运行：

~~~powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tools\validate_public_pack.ps1
~~~

检查内容包括：

- VBS 数量和 UTF-8 BOM；
- Function Main() 入口；
- Host API JSON 方法数量和 GetSelection 语义；
- 工作流、Schema 和示例 JSON 可解析；
- 公共 VBS 中不存在危险的旧式 COM/脚本入口；
- 排除目录没有被带入公共包。

## VBS 检查层级

1. cscript 语法：将 UTF-8 BOM 源文件转换为临时 UTF-16LE 副本后解析；
2. fake Host smoke：验证空文档、空选区、失败 JSON 和只读路径；
3. staging Office smoke：在 Word、Excel、PowerPoint 中使用真实文档和选区；
4. 写入回归：检查预览、确认、备份、提交、取消和回滚；
5. 发布检查：生成 ZIP、SHA256、变更清单和兼容性矩阵。

公共仓库的静态检查不能替代真实 Office 测试，尤其是 Excel 多单元格数组、PowerPoint 形状和 Word 复杂表格。
