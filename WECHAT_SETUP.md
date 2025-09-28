# 微信群消息发送配置指南

## 方案1：企业微信机器人（推荐）

### 步骤1：添加企业微信机器人到微信群

1. **在微信群中添加机器人**：
   - 打开微信群聊
   - 点击右上角的"..."菜单
   - 选择"群机器人"或"添加机器人"
   - 选择"企业微信机器人"

2. **获取Webhook地址**：
   - 机器人创建成功后，会提供一个webhook地址
   - 地址格式类似：`https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=xxxxxxxx`

### 步骤2：配置环境变量

在`.env`文件中添加：
```bash
WECHAT_WEBHOOK_URL=https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=YOUR_KEY_HERE
```

### 步骤3：运行脚本

```bash
python3 wechat_group_sender.py
```

## 方案2：通过微信开放平台API

如果您有微信开放平台的开发者账号，可以使用微信开放平台的API。

## 方案3：手动发送（临时方案）

如果上述方案都不可行，可以：

1. 运行脚本生成图片：
```bash
python3 test_table_image.py
```

2. 手动将生成的`table_with_notes.png`图片发送到微信群

## 注意事项

- 企业微信机器人有发送频率限制
- 图片大小不能超过2MB
- 需要确保机器人有发送消息的权限

## 测试配置

运行以下命令测试配置是否正确：
```bash
python3 -c "
import os
from dotenv import load_dotenv
load_dotenv()
webhook = os.getenv('WECHAT_WEBHOOK_URL', '')
if webhook:
    print('✅ Webhook配置正确')
    print(f'地址: {webhook[:50]}...')
else:
    print('❌ 未配置Webhook地址')
"
```
