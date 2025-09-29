# GitHub Actions 定时执行微信脚本设置指南

## 概述
本指南将帮助你在GitHub上设置定时器，自动执行 `wechat_with_image_fix.py` 脚本。

## 已创建的文件
- `.github/workflows/wechat-scheduler.yml` - GitHub Actions工作流配置文件

## 设置步骤

### 1. 推送配置文件到GitHub
```bash
git add .github/
git commit -m "feat: 添加GitHub Actions定时执行配置"
git push origin main
```

### 2. 在GitHub上设置环境变量
1. 进入你的GitHub仓库：https://github.com/steven054/yxc-monitor
2. 点击 **Settings** 标签页
3. 在左侧菜单中找到 **Secrets and variables** → **Actions**
4. 点击 **New repository secret**
5. 添加以下密钥：
   - **Name**: `WECHAT_WEBHOOK_URL`
   - **Value**: 你的企业微信机器人webhook地址

### 3. 定时执行配置
当前配置为每天北京时间早上7点执行：
- **Cron表达式**: `0 23 * * *` (UTC时间23:00 = 北京时间07:00)
- **时区**: UTC (GitHub Actions默认)

### 4. 手动触发测试
1. 进入GitHub仓库的 **Actions** 标签页
2. 选择 "微信定时发送监控报告" 工作流
3. 点击 **Run workflow** 按钮
4. 可以选择测试模式（不发送微信）

## 工作流功能

### 自动执行
- 每天北京时间早上7点自动执行
- 读取 `yxc.xlsx` 文件
- 生成监控报告图片
- 发送到企业微信群

### 手动触发
- 支持手动触发执行
- 可选择测试模式
- 查看执行日志和结果

### 日志和监控
- 自动上传执行日志
- 保留7天的执行记录
- 显示执行成功/失败状态

## 时间调整

如果需要修改执行时间，编辑 `.github/workflows/wechat-scheduler.yml` 文件中的cron表达式：

```yaml
schedule:
  - cron: '0 23 * * *'  # 每天北京时间07:00
```

### 常用时间示例
- `0 22 * * *` - 每天北京时间06:00
- `0 0 * * *` - 每天北京时间08:00  
- `30 22 * * *` - 每天北京时间06:30
- `0 22 * * 1-5` - 工作日北京时间06:00

## 故障排除

### 1. 检查环境变量
确保在GitHub仓库设置中正确配置了 `WECHAT_WEBHOOK_URL`

### 2. 查看执行日志
在GitHub Actions页面查看详细的执行日志

### 3. 测试脚本
可以手动触发工作流进行测试

### 4. 依赖问题
如果遇到依赖安装问题，检查 `requirements.txt` 文件

## 注意事项

1. **时区**: GitHub Actions使用UTC时间，需要转换为北京时间
2. **权限**: 确保仓库有Actions权限
3. **配额**: GitHub免费账户每月有2000分钟的执行时间限制
4. **网络**: 确保脚本能正常访问外部API

## 下一步
1. 推送配置文件到GitHub
2. 设置环境变量
3. 测试工作流
4. 监控执行结果