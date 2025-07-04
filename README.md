# Excel到期监控脚本

这个脚本可以自动监控Excel表格中的剩余天数，当发现剩余天数为0的项目时，自动发送通知到微信、邮箱或短信。

## 功能特点

- ✅ 每天早晨7点自动检查Excel表格
- ✅ 智能识别剩余天数列
- ✅ 支持多种通知方式：邮件、微信、短信
- ✅ 详细的日志输出
- ✅ 可配置的通知参数

## 安装步骤

### 1. 安装Python依赖

```bash
pip3 install -r requirements.txt
```

### 2. 配置通知参数

复制配置文件模板：
```bash
cp config.env.example .env
```

编辑 `.env` 文件，配置你需要的通知方式：

#### 邮件通知配置
```bash
EMAIL_ENABLED=true
SMTP_SERVER=smtp.gmail.com
SMTP_PORT=587
EMAIL_USERNAME=your_email@gmail.com
EMAIL_PASSWORD=your_app_password
TO_EMAIL=recipient@example.com
```

#### 微信通知配置（企业微信webhook）
```bash
WECHAT_ENABLED=true
WECHAT_WEBHOOK_URL=https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=YOUR_KEY
```

#### 短信通知配置
```bash
SMS_ENABLED=true
SMS_API_KEY=your_sms_api_key
SMS_API_URL=https://api.sms-provider.com/send
SMS_PHONE_NUMBER=13800138000
```

## 使用方法

### 方法1：使用启动脚本（推荐）
```bash
chmod +x start_monitor.sh
./start_monitor.sh
```

### 方法2：直接运行Python脚本
```bash
python3 check_expiry.py
```

## Excel表格要求

脚本会自动识别包含以下关键词的列作为剩余天数列：
- 剩余天数
- 剩余时间
- 到期天数
- 过期天数
- 天数
- days
- remaining_days
- 剩余
- 到期
- 过期

如果找不到匹配的列名，脚本会使用第一列作为剩余天数列。

## 通知方式详解

### 1. 邮件通知
支持Gmail、QQ邮箱、163邮箱等SMTP服务。

**Gmail配置示例：**
- 需要开启两步验证
- 生成应用专用密码
- SMTP服务器：smtp.gmail.com
- 端口：587

### 2. 微信通知
使用企业微信或钉钉的webhook功能。

**企业微信配置：**
1. 在企业微信群中添加机器人
2. 获取webhook URL
3. 配置到 `WECHAT_WEBHOOK_URL`

### 3. 短信通知
需要根据具体的短信服务商API进行调整。

## 日志输出

脚本运行时会输出详细的日志信息：
- Excel文件读取状态
- 找到的剩余天数列
- 发现的到期项目
- 通知发送状态

## 定时任务

脚本默认每天早晨7点执行检查。如果需要修改时间，可以编辑 `check_expiry.py` 文件中的这一行：

```python
schedule.every().day.at("07:00").do(checker.run_check)
```

## 故障排除

### 1. Excel文件读取失败
- 确保Excel文件存在且可读
- 检查文件是否被其他程序占用
- 确保文件格式为.xlsx

### 2. 通知发送失败
- 检查网络连接
- 验证配置参数是否正确
- 查看具体的错误信息

### 3. 找不到剩余天数列
- 检查Excel表格中是否有包含"天数"、"剩余"等关键词的列
- 可以手动修改 `find_expiry_column` 方法

## 注意事项

1. 脚本需要持续运行才能执行定时任务
2. 建议在服务器或云主机上运行
3. 确保Excel文件路径正确
4. 定期检查日志确保脚本正常运行

## 技术支持

如果遇到问题，请检查：
1. Python版本（建议3.7+）
2. 依赖包是否正确安装
3. 配置文件格式是否正确
4. 网络连接是否正常 