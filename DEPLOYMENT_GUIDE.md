# 🚀 云平台部署指南

## 📋 概述

为了避免本地电脑关机导致监控脚本停止，我们提供了多种免费的云平台部署方案。

## 🎯 推荐方案：GitHub Actions（最推荐）

### 优点
- ✅ 完全免费
- ✅ 无需服务器
- ✅ 自动定时执行
- ✅ 自动备份和版本控制
- ✅ 可以手动触发

### 部署步骤

#### 1. 创建GitHub仓库
```bash
# 在你的GitHub账户下创建一个新仓库
# 例如：yxc-monitor
```

#### 2. 上传文件到GitHub
```bash
# 克隆仓库
git clone https://github.com/你的用户名/yxc-monitor.git
cd yxc-monitor

# 复制文件到仓库
cp cloud_monitor.py ./
cp yxc.xlsx ./
cp .github/workflows/monitor.yml ./.github/workflows/
cp requirements.txt ./

# 提交到GitHub
git add .
git commit -m "初始提交"
git push origin main
```

#### 3. 设置GitHub Secrets
在GitHub仓库页面：
1. 点击 `Settings` → `Secrets and variables` → `Actions`
2. 点击 `New repository secret`
3. 添加以下环境变量：

**邮件配置：**
```
EMAIL_ENABLED: true
SMTP_SERVER: smtp.163.com
SMTP_PORT: 465
EMAIL_USERNAME: 你的邮箱@163.com
EMAIL_PASSWORD: 你的授权码
TO_EMAIL: 接收通知的邮箱
```

**微信配置（可选）：**
```
WECHAT_ENABLED: true
WECHAT_WEBHOOK_URL: 你的企业微信webhook地址
```

**短信配置（可选）：**
```
SMS_ENABLED: true
SMS_API_URL: 短信API地址
SMS_API_KEY: 短信API密钥
SMS_PHONE_NUMBER: 接收短信的手机号
```

#### 4. 测试运行
1. 在GitHub仓库页面点击 `Actions`
2. 选择 `智能监控任务`
3. 点击 `Run workflow` 手动触发一次测试

## 🌐 其他云平台方案

### Railway.app

#### 优点
- ✅ 每月有免费额度
- ✅ 部署简单
- ✅ 支持Python

#### 部署步骤
1. 注册 [Railway.app](https://railway.app/)
2. 连接GitHub仓库
3. 创建新项目
4. 设置环境变量
5. 部署

### Render.com

#### 优点
- ✅ 有免费计划
- ✅ 支持Python应用
- ✅ 可以设置定时任务

#### 部署步骤
1. 注册 [Render.com](https://render.com/)
2. 连接GitHub仓库
3. 创建Web Service
4. 设置环境变量
5. 部署

### PythonAnywhere

#### 优点
- ✅ 专门为Python设计
- ✅ 有免费计划
- ✅ 内置定时任务功能

#### 部署步骤
1. 注册 [PythonAnywhere](https://www.pythonanywhere.com/)
2. 上传文件
3. 设置定时任务
4. 配置环境变量

## ⚙️ 环境变量配置

### 必需的环境变量
```bash
EXCEL_FILE=yxc.xlsx
```

### 邮件通知配置
```bash
EMAIL_ENABLED=true
SMTP_SERVER=smtp.163.com
SMTP_PORT=465
EMAIL_USERNAME=你的邮箱@163.com
EMAIL_PASSWORD=你的授权码
TO_EMAIL=接收通知的邮箱
```

### 微信通知配置（可选）
```bash
WECHAT_ENABLED=true
WECHAT_WEBHOOK_URL=你的企业微信webhook地址
```

### 短信通知配置（可选）
```bash
SMS_ENABLED=true
SMS_API_URL=短信API地址
SMS_API_KEY=短信API密钥
SMS_PHONE_NUMBER=接收短信的手机号
```

## 🔧 自定义配置

### 修改执行时间
在 `.github/workflows/monitor.yml` 中修改cron表达式：
```yaml
- cron: '1 13 * * *'  # UTC 13:01 = 北京时间 21:01
```

### 修改Excel文件名
在环境变量中设置：
```bash
EXCEL_FILE=你的文件名.xlsx
```

## 📊 监控和日志

### GitHub Actions
- 在GitHub仓库的 `Actions` 页面查看执行日志
- 可以看到每次执行的详细输出
- 支持手动触发测试

### 文件备份
- 每次执行前会自动备份Excel文件
- 备份文件保存在 `backups` 目录
- 文件名包含时间戳

## 🚨 故障排除

### 常见问题

1. **Excel文件不存在**
   - 确保文件名正确
   - 检查文件是否上传到GitHub

2. **邮件发送失败**
   - 检查邮箱配置
   - 确认授权码正确
   - 检查网络连接

3. **GitHub Actions执行失败**
   - 查看Actions日志
   - 检查环境变量配置
   - 确认依赖包安装成功

### 联系支持
如果遇到问题，可以：
1. 查看GitHub Actions的执行日志
2. 检查环境变量配置
3. 确认Excel文件格式正确

## 🎉 完成部署

部署完成后，你的监控系统将：
- ✅ 每天自动执行
- ✅ 自动更新剩余天数
- ✅ 自动重置到期项目
- ✅ 自动发送通知
- ✅ 自动备份文件
- ✅ 无需本地电脑运行

享受自动化的便利吧！🚀 