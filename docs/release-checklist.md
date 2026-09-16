# Public Release Checklist

## 内容检查

- [ ] 所有新增 VBS、示例代码和原创文档采用 Apache-2.0。
- [ ] 没有软著、ICP备案、专利、市场申报、后端、部署、签名或生产资料。
- [ ] 没有客户文件、个人信息、内部路径、令牌、密钥或完整日志。
- [ ] 工作流 Action 没有 CreateObject、GetObject、ADODB.Stream、FSO 或 WScript.Shell。
- [ ] 第三方内容的许可证和 NOTICE 已保留。

## 技术检查

- [ ] 运行 tools/validate_public_pack.ps1。
- [ ] VBS 保持 UTF-8 BOM。
- [ ] Host API 方法数量和选区语义与当前 Release 一致。
- [ ] JSON Schema 和示例 JSON 可解析。
- [ ] Word、Excel、PowerPoint 各至少完成一次 staging smoke。
- [ ] 写入脚本验证预览、确认、备份和回滚。

## 发布物

- [ ] 创建 community-vX.Y.Z 标签。
- [ ] 发布公共 ZIP、SHA256、CHANGELOG 和兼容性说明。
- [ ] 记录本次脚本、提示词、Schema 和 Host 契约变化。
- [ ] 安装包作为独立 Release asset，使用产品许可。
- [ ] 先完成 staging，再获得明确确认后处理 prod。
