# VBS Script Library

本目录只收录经过 Host-first 验证的 105 个预设脚本，另有 `examples/vbs/` 中的 2 个最小示例。脚本保留 UTF-8 BOM，适合由 OfficeAddin 的 VBS Host 运行时加载。

## 分类

| 目录 | 数量 | 内容 |
|---|---:|---|
| common/ | 11 | 上下文、文件、剪贴板、计划和诊断 |
| word/ | 24 | Word 排版、审阅、导出和选区处理 |
| excel/ | 42 | Excel 清洗、校验、汇总、映射和格式化 |
| powerpoint/ | 28 | PowerPoint 文本、版式、导出和对象处理 |

脚本还可按头部描述区分只读、预览和写入型操作。写入型脚本应在运行前展示影响范围，并使用 Host.BeginWritePlan()、Host.PreviewWritePlan()、Host.ValidateWritePlan() 和 Host.CommitWritePlan() 等能力。

## 运行方式

这些脚本依赖 OfficeAddin 注入的 Host 对象，不能假设直接使用 cscript script.vbs 时存在真实 Office 上下文。没有 Host 时，脚本应返回可读的错误 JSON 或空结果，不应创建新的隐藏 Office 进程。

## 安装时的重要区别

这里的分类目录只服务于 GitHub 阅读和贡献审核。插件运行时不会递归扫描这里的子目录；请把选中的 `.vbs` 文件直接复制到：

```text
<OfficeAddin 安装目录>\VBA Script\
```

文件必须带标准头部注释和无参数 `Function Main()`。复制后重启 Office，或使用 `VbsProxy.exe -index -scriptdir "<安装目录>\VBA Script" -force` 刷新命令索引。

编码、注释、入口函数和安全检查要求见 [`../docs/vbs-contribution-format.md`](../docs/vbs-contribution-format.md)、[`../docs/vbs-embedding.md`](../docs/vbs-embedding.md) 和 [`../docs/testing.md`](../docs/testing.md)。
