#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
深化设计数据查询工具 - API模块

此模块负责与简道云API进行交互，封装了API访问方法和数据获取功能。
主要组件:
1. JianDaoYunAPI类 - 封装了简道云API的所有访问方法
   - 项目数据查询
   - 深化清单详情获取
   - 批量数据获取

该模块隔离了API访问逻辑，使系统更易于维护和扩展。
"""

import requests
import json
from typing import List, Dict, Any

# 尝试导入PySide6，如果不可用则使用纯控制台输出
try:
    from PySide6.QtWidgets import QMessageBox
    USE_GUI = True
except ImportError:
    USE_GUI = False

# 导入FormFields类用于字段访问
from io_generator import FormFields

# 简道云API基础URL
BASE_URL = "https://api.jiandaoyun.com/api"
API_ENDPOINT = f"{BASE_URL}/v5/app/entry/data/list"

def show_message(title, message, message_type="error", parent=None):
    """
    显示消息，支持GUI和控制台两种方式
    
    Args:
        title: 消息标题
        message: 消息内容
        message_type: 消息类型 (info, warning, error)
        parent: 父窗口
    """
    if USE_GUI:
        if message_type == "info":
            QMessageBox.information(parent, title, message)
        elif message_type == "warning":
            QMessageBox.warning(parent, title, message)
        else:  # error
            QMessageBox.critical(parent, title, message)
    else:
        print(f"{title}: {message}")

class JianDaoYunAPI:
    """
    简道云API客户端类
    
    负责与简道云API进行交互，包括数据查询、获取详情等功能
    封装了API访问的细节，提供简洁的接口供上层调用
    """
    
    def __init__(self, api_key: str, app_id: str, entry_id: str, 
                 field_mapping: Dict[str, str], subform_field_id: str):
        """
        初始化简道云API客户端
        
        Args:
            api_key: API密钥
            app_id: 应用ID
            entry_id: 表单ID
            field_mapping: 字段映射字典
            subform_field_id: 子表单字段ID
        """
        self.api_key = api_key
        self.app_id = app_id
        self.entry_id = entry_id
        self.field_mapping = field_mapping
        self.subform_field_id = subform_field_id
    
    def get_auth_header(self) -> Dict[str, str]:
        """
        生成简道云API认证头 - 使用Bearer认证
        
        Returns:
            dict: 包含认证信息的请求头字典
        """
        # 构建认证头
        headers = {
            'Authorization': f'Bearer {self.api_key}',  # Bearer认证方式
            'Content-Type': 'application/json'          # 请求内容类型为JSON
        }
        
        return headers
    
    def search_project_data(self, project_number: str) -> List[Dict[str, Any]]:
        """
        根据项目编号搜索项目数据
        
        Args:
            project_number: 项目编号，如"OPP.25011100829"
            
        Returns:
            list: 匹配的项目数据记录列表
        """
        # 检查API凭证是否已设置
        if not self.api_key:
            show_message("错误", "API凭证未配置，请在配置文件中设置有效的API_KEY")
            return []
            
        # 指定要返回的字段，确保包含_id字段
        fields = list(self.field_mapping.values())
        # 添加_id字段，如果列表中不存在
        if "_id" not in fields:
            fields.append("_id")
        
        # 构建API请求载荷
        payload = {
            "app_id": self.app_id,         # 应用ID
            "entry_id": self.entry_id,     # 表单ID
            "limit": 100,                  # 最大返回数量
            "fields": fields,              # 需要返回的字段列表
            "filter": {                    # 查询条件
                "rel": "and",              # 条件关系：且
                "cond": [                  # 条件列表
                    {
                        "field": self.field_mapping[FormFields.MainForm.PROJECT_NUMBER],  # 按项目编号字段过滤
                        "operator": "eq",                                                 # 操作符：等于
                        "value": project_number                                           # 查询值：项目编号
                    }
                ]
            }
        }
        
        try:
            # 获取认证头
            headers = self.get_auth_header()
            
            # 发送API请求
            response = requests.post(API_ENDPOINT, headers=headers, json=payload, timeout=10)
            
            # 检查HTTP响应状态
            if response.status_code != 200:
                error_msg = f"API请求失败，HTTP错误: {response.status_code}\n错误详情: {response.text}"
                show_message("API错误", error_msg)
                return []
            
            # 解析JSON响应
            result = response.json()
            
            # 检查响应中是否包含数据
            if 'data' in result and isinstance(result['data'], list):
                # 筛选出与指定项目编号匹配的记录
                matched_data = []
                for item in result['data']:
                    # 确保记录中包含_id字段
                    if '_id' not in item:
                        print("警告: 记录中没有_id字段，可能无法查询详情")
                        
                    field_id = self.field_mapping[FormFields.MainForm.PROJECT_NUMBER]
                    if item.get(field_id) == project_number:
                        matched_data.append(item)
                
                if matched_data:
                    return matched_data
                else:
                    print(f"提示: API返回了数据，但没有匹配项目编号 '{project_number}' 的记录。")
                    return []
            else:
                print("警告: API响应中没有找到数据字段。")
                return []
                
        except requests.exceptions.RequestException as e:
            show_message("网络错误", f"网络请求错误: {str(e)}")
            return []
        except json.JSONDecodeError as e:
            show_message("解析错误", f"JSON解析错误: {str(e)}")
            return []
        except Exception as e:
            show_message("未知错误", f"查询数据时发生未知错误: {str(e)}")
            return []
    
    def get_shenhua_detail(self, data_id: str) -> List[Dict[str, Any]]:
        """
        获取深化清单详情数据
        
        根据表单数据ID获取深化清单子表单的详细数据。
        
        Args:
            data_id: 表单数据ID，如"67d7f1b07e1277592a4bd513"
            
        Returns:
            list: 深化清单子表单数据列表，如果发生错误则返回空列表
        """
        # 检查API凭证是否已设置
        if not self.api_key:
            show_message("错误", "API凭证未配置")
            return []
        
        try:
            # 获取认证头
            headers = self.get_auth_header()
            
            # 构建API请求载荷
            payload = {
                "app_id": self.app_id,     # 应用ID
                "entry_id": self.entry_id, # 表单ID
                "limit": 1,                # 只需要1条记录
                "filter": {                # 查询条件
                    "rel": "and",          # 条件关系：且
                    "cond": [              # 条件列表
                        {
                            "field": "_id",      # 按_id字段过滤
                            "operator": "eq",    # 操作符：等于
                            "value": data_id     # 查询值：数据ID
                        }
                    ]
                }
            }
            
            # 发送API请求获取数据
            response = requests.post(API_ENDPOINT, headers=headers, json=payload, timeout=10)
            
            # 检查HTTP响应状态
            if response.status_code != 200:
                error_msg = f"API请求失败，HTTP错误: {response.status_code}\n错误详情: {response.text}"
                show_message("API错误", error_msg)
                return []
            
            # 解析JSON响应
            result = response.json()
            
            # 判断是否有数据返回
            if 'data' not in result or not isinstance(result['data'], list) or not result['data']:
                show_message("数据错误", "API响应中没有找到有效数据")
                return []
            
            # 从数据中提取第一条记录
            data = result['data'][0]
            
            # 检查返回的数据ID是否与请求的ID一致
            returned_id = data.get('_id', 'unknown')
            if returned_id != data_id:
                show_message("数据警告", f"返回的数据ID ({returned_id}) 与请求的ID ({data_id}) 不匹配")
            
            # 检查是否存在子表单字段
            if self.subform_field_id not in data:
                error_msg = f"响应中没有找到深化清单子表单字段 {self.subform_field_id}"
                show_message("字段错误", error_msg)
                return []
                
            shenhua_list = data.get(self.subform_field_id, [])
            
            if not shenhua_list:
                show_message("提示", "未找到深化清单数据")
                return []
            
            return shenhua_list
            
        except requests.exceptions.RequestException as e:
            show_message("网络错误", f"网络请求错误: {str(e)}")
            return []
        except json.JSONDecodeError as e:
            show_message("解析错误", f"JSON解析错误: {str(e)}")
            return []
        except Exception as e:
            error_msg = f"获取深化清单详情时发生未知错误: {str(e)}"
            show_message("未知错误", error_msg)
            return []

    def get_all_shenhua_data(self) -> List[Dict[str, Any]]:
        """
        获取所有深化清单数据
        
        Returns:
            list: 所有深化清单数据列表
        """
        # 检查API凭证是否已设置
        if not self.api_key:
            show_message("错误", "API凭证未配置")
            return []
        
        try:
            # 获取认证头
            headers = self.get_auth_header()
            
            # 构建API请求载荷 - 不设置筛选条件，获取所有数据
            payload = {
                "app_id": self.app_id,     # 应用ID
                "entry_id": self.entry_id, # 表单ID
                "limit": 100,              # 最大返回数量
                "fields": ["_id", self.subform_field_id]  # 只获取必要字段
            }
            
            # 添加场站和项目编号字段
            for field_name, field_id in self.field_mapping.items():
                if field_name in ["场站", "项目编号", "项目名称"]:
                    payload["fields"].append(field_id)
            
            # 发送API请求获取数据
            response = requests.post(API_ENDPOINT, headers=headers, json=payload, timeout=15)
            
            # 检查HTTP响应状态
            if response.status_code != 200:
                error_msg = f"API请求失败，HTTP错误: {response.status_code}\n错误详情: {response.text}"
                show_message("API错误", error_msg)
                return []
            
            # 解析JSON响应
            result = response.json()
            
            # 判断是否有数据返回
            if 'data' not in result or not isinstance(result['data'], list) or not result['data']:
                show_message("数据错误", "API响应中没有找到有效数据")
                return []
            
            # 获取所有数据
            all_data = result['data']
            return all_data
            
        except requests.exceptions.RequestException as e:
            show_message("网络错误", f"网络请求错误: {str(e)}")
            return []
        except json.JSONDecodeError as e:
            show_message("解析错误", f"JSON解析错误: {str(e)}")
            return []
        except Exception as e:
            error_msg = f"获取数据时发生未知错误: {str(e)}"
            show_message("未知错误", error_msg)
            return [] 