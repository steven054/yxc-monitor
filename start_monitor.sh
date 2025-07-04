#!/bin/bash

# 启动Excel监控脚本
echo "启动Excel监控脚本..."

# 检查Python是否安装
if ! command -v python3 &> /dev/null; then
    echo "错误: 未找到Python3，请先安装Python3"
    exit 1
fi

# 检查依赖是否安装
if [ ! -f "requirements.txt" ]; then
    echo "错误: 未找到requirements.txt文件"
    exit 1
fi

# 安装依赖
echo "安装Python依赖..."
pip3 install -r requirements.txt

# 检查配置文件
if [ ! -f ".env" ]; then
    echo "警告: 未找到.env配置文件，将使用默认配置"
    echo "请复制config.env.example为.env并配置通知参数"
fi

# 启动监控脚本
echo "启动智能监控脚本..."
python3 smart_monitor.py 