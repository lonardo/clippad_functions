# Performance Disclosure

当前公开文案采用“低常驻资源占用、尽量不打扰 Office 用户”的定性描述，不发布未经统一测试的 CPU、内存或 IO 数值，也不列硬件最低参数。

## 后续测量口径

正式发布性能数据前，应在 Windows 10/11、Office 2010 及以上代表版本、32/64 位组合中分别测量：

- Office 启动后空闲 10 分钟的 CPU、Working Set、Private Bytes；
- 选区变化和普通 Ribbon 操作的 CPU 峰值与持续时间；
- 典型 VBS 预览、单文档写入和批量文件处理；
- AI 请求期间的本地 CPU/内存/IO 与等待时间；
- 失败、取消、回滚后的资源是否释放。

结果应报告测试环境、样本数量、p50/p95 和峰值，区分 Office 自身资源、插件资源、VBS/COM 操作资源和网络请求资源。
