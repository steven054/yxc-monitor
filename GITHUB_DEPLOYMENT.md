# GitHub部署指南

## 🚀 部署步骤

### 1. 推送代码到GitHub
```bash
git add .
git commit -m "准备GitHub部署"
git push origin main
```

### 2. 设置GitHub Secrets
在GitHub仓库中设置以下Secrets（Settings → Secrets and variables → Actions）：

#### 邮件配置
- `EMAIL_ENABLED`: `true`
- `SMTP_SERVER`: `smtp.163.com`
- `SMTP_PORT`: `465`
- `EMAIL_USERNAME`: `shidewei054@163.com`
- `EMAIL_PASSWORD`: `CEpJp32m4rX6weNH`
- `TO_EMAIL`: `408838485@qq.com`

#### 微信配置（可选）
- `WECHAT_ENABLED`: `false`
- `WECHAT_WEBHOOK_URL`: `https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=YOUR_KEY`

#### 短信配置（可选）
- `SMS_ENABLED`: `false`
- `SMS_API_KEY`: `your_sms_api_key`
- `SMS_API_URL`: `https://api.sms-provider.com/send`
- `SMS_PHONE_NUMBER`: `13800138000`

### 3. 验证部署
1. 进入GitHub仓库的Actions页面
2. 点击"智能监控任务"
3. 点击"Run workflow"手动触发一次测试
4. 检查是否成功执行并发送邮件

## 📅 执行时间
- **自动执行**: 每天北京时间6:30（UTC 22:30）
- **手动触发**: 可在Actions页面手动运行

## 📧 邮件通知
- 当有项目到期时，会发送详细的通知邮件
- 包含到期项目列表和已重置项目信息

## 🔄 自动更新
- 每天自动将剩余天数减1
- 当剩余天数为0时，自动重置为总天数
- 更新开始时间为当天日期
- 自动提交更新到GitHub

## 📁 文件说明
- `github_monitor.py`: GitHub Actions专用脚本
- `smart_monitor.py`: 本地运行脚本（带定时循环）
- `yxc.xlsx`: 主数据文件
- `yxc_backup.xlsx`: 自动备份文件

## 🛠️ 故障排除
1. **检查Secrets设置**: 确保所有环境变量都正确设置
2. **查看Actions日志**: 在Actions页面查看详细执行日志
3. **测试邮件配置**: 手动触发一次检查邮件是否正常发送
4. **检查文件权限**: 确保Excel文件有读写权限

## 🎯 功能特性
✅ 自动备份Excel文件  
✅ 智能列名识别  
✅ 剩余天数自动减1  
✅ 到期项目自动重置  
✅ 邮件通知  
✅ 自动提交更新  
✅ 错误处理和日志记录  

部署完成后，您的监控系统将完全自动化运行！🎉 