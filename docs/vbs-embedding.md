# VBS 如何嵌入 OfficeAddin

## 运行模型

OfficeAddin 加载 VBS 后，在脚本运行时提供 Host 调度对象。脚本通过 Host.GetApplication()、Host.GetDocument()、Host.GetSelectionObject() 等接口访问当前 Office 上下文；不需要自行创建新的 Word、Excel 或 PowerPoint 进程。

脚本入口约定为：

~~~vbscript
Function Main()
    Main = "{""ok"":true}"
End Function
~~~

宿主读取 Main() 返回的字符串作为脚本结果。建议返回可解析的 JSON，并包含 ok、code、message 或 data 字段。

## 编码

- 公共 VBS 源文件使用 UTF-8 BOM；
- 不要将现有 BOM 文件转换为无 BOM 或本机 ANSI；
- 生成端写出并通过 cscript 执行的临时 VBS 建议使用 UTF-16LE；
- Windows cscript 直接读取 UTF-8 BOM 源文件时可能报编码错误；这不代表脚本逻辑错误，公共源文件应由 OfficeAddin 加载，独立 cscript 烟测应使用 UTF-16LE 临时副本；
- 不要在脚本中使用 ADODB.Stream、FileSystemObject 或 WScript.Shell 代替 Host 文件能力。

## 安全执行顺序

1. 检查当前应用、文档和选区是否满足前置条件。
2. 读取上下文摘要，估算批处理规模和输出路径。
3. 对写入操作构建写入计划。
4. 展示预览并获取用户确认。
5. 提交写入计划；失败时执行回滚或恢复备份。
6. 返回结果 JSON，同时将必要摘要写入 Host 日志或剪贴板。

## 独立 cscript 测试

直接用 cscript 运行公共脚本只能验证语法和 fake Host 路径，不能证明真实 Office 对象操作成功。真实验证必须在 staging 客户端中打开对应的 Word、Excel 或 PowerPoint 后执行。
