#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
短信内容配置工具
"""

import json
import os
from datetime import datetime

class SMSContentConfig:
    def __init__(self):
        self.config_file = "sms_templates.json"
        self.default_templates = {
            "single_item": {
                "template": "【到期提醒】{store_name}({address})已到期，请及时处理！时间:{date}",
                "description": "单个项目到期的短信模板"
            },
            "multiple_items": {
                "template": "【到期提醒】{count}个项目已到期，请及时处理！时间:{date}",
                "description": "多个项目到期的短信模板"
            },
            "urgent": {
                "template": "🚨紧急：{store_name}已到期！地址：{address}，请立即处理！",
                "description": "紧急提醒模板"
            },
            "simple": {
                "template": "有{count}个项目到期了，请查看Excel表格处理。",
                "description": "简单提醒模板"
            }
        }
        self.load_config()
    
    def load_config(self):
        """加载配置文件"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    self.templates = json.load(f)
                print(f"✅ 已加载短信模板配置: {self.config_file}")
            except Exception as e:
                print(f"❌ 加载配置文件失败: {e}")
                self.templates = self.default_templates.copy()
        else:
            self.templates = self.default_templates.copy()
            self.save_config()
    
    def save_config(self):
        """保存配置文件"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.templates, f, ensure_ascii=False, indent=2)
            print(f"✅ 短信模板配置已保存: {self.config_file}")
        except Exception as e:
            print(f"❌ 保存配置文件失败: {e}")
    
    def get_template(self, template_name):
        """获取指定模板"""
        return self.templates.get(template_name, self.default_templates.get(template_name))
    
    def add_template(self, name, template, description=""):
        """添加新模板"""
        self.templates[name] = {
            "template": template,
            "description": description
        }
        self.save_config()
        print(f"✅ 已添加模板: {name}")
    
    def list_templates(self):
        """列出所有模板"""
        print("\n📋 当前可用的短信模板:")
        print("=" * 50)
        for name, config in self.templates.items():
            print(f"📝 {name}:")
            print(f"   描述: {config['description']}")
            print(f"   模板: {config['template']}")
            print()
    
    def test_template(self, template_name, test_data=None):
        """测试模板"""
        template_config = self.get_template(template_name)
        if not template_config:
            print(f"❌ 模板 '{template_name}' 不存在")
            return
        
        if test_data is None:
            # 使用默认测试数据
            test_data = {
                'store_name': '测试店铺',
                'address': '测试地址',
                'count': 3,
                'date': datetime.now().strftime("%Y-%m-%d")
            }
        
        try:
            content = template_config['template'].format(**test_data)
            print(f"\n🧪 模板测试结果:")
            print(f"模板名称: {template_name}")
            print(f"模板内容: {template_config['template']}")
            print(f"测试数据: {test_data}")
            print(f"生成内容: {content}")
            print(f"内容长度: {len(content)} 字符")
            
            if len(content) > 70:
                print("⚠️  警告: 短信内容较长，可能超出单条短信长度限制")
            else:
                print("✅ 短信长度合适")
                
        except KeyError as e:
            print(f"❌ 模板测试失败: 缺少参数 {e}")
        except Exception as e:
            print(f"❌ 模板测试失败: {e}")
    
    def interactive_config(self):
        """交互式配置"""
        print("🎯 短信模板配置工具")
        print("=" * 50)
        
        while True:
            print("\n请选择操作:")
            print("1. 查看所有模板")
            print("2. 测试模板")
            print("3. 添加新模板")
            print("4. 修改现有模板")
            print("5. 删除模板")
            print("6. 退出")
            
            choice = input("\n请输入选择 (1-6): ").strip()
            
            if choice == '1':
                self.list_templates()
            
            elif choice == '2':
                self.list_templates()
                template_name = input("请输入要测试的模板名称: ").strip()
                self.test_template(template_name)
            
            elif choice == '3':
                name = input("请输入新模板名称: ").strip()
                template = input("请输入模板内容 (使用{变量名}作为占位符): ").strip()
                description = input("请输入模板描述: ").strip()
                self.add_template(name, template, description)
            
            elif choice == '4':
                self.list_templates()
                name = input("请输入要修改的模板名称: ").strip()
                if name in self.templates:
                    template = input("请输入新的模板内容: ").strip()
                    description = input("请输入新的模板描述: ").strip()
                    self.templates[name] = {
                        "template": template,
                        "description": description
                    }
                    self.save_config()
                    print(f"✅ 已修改模板: {name}")
                else:
                    print(f"❌ 模板 '{name}' 不存在")
            
            elif choice == '5':
                self.list_templates()
                name = input("请输入要删除的模板名称: ").strip()
                if name in self.templates:
                    confirm = input(f"确认删除模板 '{name}'? (y/N): ").strip().lower()
                    if confirm == 'y':
                        del self.templates[name]
                        self.save_config()
                        print(f"✅ 已删除模板: {name}")
                else:
                    print(f"❌ 模板 '{name}' 不存在")
            
            elif choice == '6':
                print("👋 退出配置工具")
                break
            
            else:
                print("❌ 无效选择，请重新输入")

def main():
    """主函数"""
    config = SMSContentConfig()
    
    if len(os.sys.argv) > 1:
        command = os.sys.argv[1]
        
        if command == "list":
            config.list_templates()
        elif command == "test":
            template_name = os.sys.argv[2] if len(os.sys.argv) > 2 else "single_item"
            config.test_template(template_name)
        elif command == "add":
            if len(os.sys.argv) >= 4:
                name = os.sys.argv[2]
                template = os.sys.argv[3]
                description = os.sys.argv[4] if len(os.sys.argv) > 4 else ""
                config.add_template(name, template, description)
            else:
                print("用法: python3 sms_config.py add <模板名称> <模板内容> [描述]")
        else:
            print("用法:")
            print("  python3 sms_config.py list                    # 列出所有模板")
            print("  python3 sms_config.py test [模板名称]         # 测试模板")
            print("  python3 sms_config.py add <名称> <内容> [描述] # 添加模板")
            print("  python3 sms_config.py                         # 交互式配置")
    else:
        config.interactive_config()

if __name__ == "__main__":
    main() 