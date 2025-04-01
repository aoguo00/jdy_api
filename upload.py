#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
深化设计数据查询工具 - 文件上传模块

此模块负责将生成的HMI点表和PLC点表上传到简道云平台
实现了文件上传的核心功能，支持进度显示和错误处理
"""

import requests
import json
import uuid
import os
from tkinter import messagebox
import traceback

# 导入API配置
from config.api_config import (
    API_KEY,
    API_V5_BASE_URL,
    get_app_id,
    get_entry_id,
    get_field_id
)

# 获取文件上传应用和表单信息
APP_ID = get_app_id("文件上传")
ENTRY_ID = get_entry_id("文件上传", "文件目录")
BASE_URL = API_V5_BASE_URL
# 表单字段
upload_widget = get_field_id("文件上传", "文件目录", "上传字段")
file_type_widget = get_field_id("文件上传", "文件目录", "文件类型")


def upload_file(file_path, file_type="其他"):
    """
    将文件上传到简道云并关联到表单
    
    参数:
        file_path (str): 要上传文件的本地路径
        file_type (str): 文件类型描述，例如"HMI点表"或"PLC点表"
        
    返回:
        str/bool: 成功返回True，失败返回False
    """
    try:
        # 获取文件名和生成唯一事务ID
        file_name = os.path.basename(file_path)
        transaction_id = str(uuid.uuid4())
        
        # ========== 步骤1: 获取上传凭证 ==========
        token_url = f"{BASE_URL}/app/entry/file/get_upload_token"
        headers = {
            "Authorization": f"Bearer {API_KEY}",
            "Content-Type": "application/json"
        }
        token_data = {
            "app_id": APP_ID,
            "entry_id": ENTRY_ID,
            "transaction_id": transaction_id,
            "data": {
                upload_widget: {
                    "value": []
                }
            }
        }
        
        # 发送请求获取上传凭证
        token_response = requests.post(token_url, json=token_data, headers=headers)
        token_response.raise_for_status()
        
        # 解析上传凭证和URL
        token_info = token_response.json()["token_and_url_list"][0]
        upload_url = token_info["url"]
        upload_token = token_info["token"]
        
        # ========== 步骤2: 上传文件 ==========
        # 打开文件并准备上传
        file = {'file': (file_name, open(file_path, 'rb'), 'application/form-data')}
        params = {'token': upload_token}
        
        # 发送上传请求
        upload_response = requests.post(upload_url, files=file, data=params)
        upload_response.raise_for_status()
        
        # 关闭文件句柄
        file['file'][1].close()
        
        # 解析上传结果
        result = upload_response.json()
        if 'key' not in result:
            return False
        
        # 获取文件标识符key
        file_key = result["key"]
        
        # ========== 步骤3: 创建表单数据关联文件 ==========
        # 等待确保文件上传处理完成
        # time.sleep(1)
        
        # 准备创建表单数据请求
        create_form_url = f"{BASE_URL}/app/entry/data/create"
        create_form_data = {
            "app_id": APP_ID,
            "entry_id": ENTRY_ID,
            "transaction_id": transaction_id,  # 使用同一个transaction_id确保关联
            "data": {
                upload_widget: {
                    "value": [
                        file_key
                    ]
                },
                file_type_widget: {  # 文件类型字段
                    "value": file_type
                }
            }
        }
        
        # 发送创建表单数据请求
        create_response = requests.post(
            create_form_url, 
            json=create_form_data, 
            headers=headers
        )
        create_response.raise_for_status()
        
        return True
    except FileNotFoundError:
        print(f"文件未找到: {file_path}")
        return False
    except requests.exceptions.RequestException as e:
        print(f"API请求错误: {str(e)}")
        return False
    except json.JSONDecodeError as e:
        print(f"JSON解析错误: {str(e)}")
        return False
    except Exception as e:
        print(f"上传过程中发生错误: {str(e)}")
        return False


def upload_hmi_table(io_data, root_window=None):
    """
    生成并上传HMI点表到简道云
    
    参数:
        io_data (pandas.DataFrame): IO数据
        root_window (tk.Tk): 主窗口对象，用于创建进度窗口
        
    返回:
        bool: 上传成功返回True，失败返回False
    """
    try:
        # 导入HMI生成器
        from hmi_generator import HMIGenerator
        import tempfile
        
        # 创建临时文件
        temp_dir = tempfile.gettempdir()
        temp_hmi_file_path = os.path.join(temp_dir, "HMI点表.xls")
        temp_dict_file_path = os.path.join(temp_dir, "数据词典点表.xls")
        
        # 调用HMI生成器生成IO_Server点表
        hmi_success = HMIGenerator.generate_hmi_table(
            io_data=io_data,
            output_path=temp_hmi_file_path,
            root_window=root_window
        )
        
        if not hmi_success:
            return False
            
        # 调用HMI生成器生成数据词典点表
        dict_success = HMIGenerator.generate_data_dictionary_table(
            io_data=io_data,
            output_path=temp_dict_file_path,
            root_window=root_window
        )
        
        if not dict_success:
            messagebox.showwarning("警告", "HMI点表生成成功，但数据词典点表生成失败。")
        
        # 上传HMI点表到简道云
        hmi_upload_success = upload_file(temp_hmi_file_path, "HMI点表")
        
        # 如果数据词典点表生成成功，则上传
        dict_upload_success = False
        if dict_success:
            dict_upload_success = upload_file(temp_dict_file_path, "数据词典点表")
        
        # 根据上传结果显示不同的提示信息
        if hmi_upload_success and dict_upload_success:
            messagebox.showinfo("成功", "HMI点表和数据词典点表已成功生成并上传到简道云!")
            return True
        elif hmi_upload_success and not dict_upload_success and dict_success:
            messagebox.showinfo("部分成功", "HMI点表已成功上传，但数据词典点表上传失败。")
            return True
        elif hmi_upload_success and not dict_success:
            messagebox.showinfo("部分成功", "HMI点表已成功上传，但数据词典点表生成失败。")
            return True
        else:
            messagebox.showerror("错误", "上传HMI点表到简道云失败!")
            return False
            
    except Exception as e:
        error_details = traceback.format_exc()
        messagebox.showerror("错误", f"生成或上传HMI点表时发生错误:\n{str(e)}\n\n详细错误信息:\n{error_details}")
        return False


def upload_plc_table(io_data, root_window=None):
    """
    生成并上传PLC点表到简道云
    
    参数:
        io_data (pandas.DataFrame): IO数据
        root_window (tk.Tk): 主窗口对象，用于创建进度窗口
        
    返回:
        bool: 上传成功返回True，失败返回False
    """
    try:
        # 导入PLC生成器
        from plc_generator import PLCGenerator
        import tempfile
        
        # 创建临时文件
        temp_dir = tempfile.gettempdir()
        temp_file_path = os.path.join(temp_dir, "PLC点表.xls")
        
        # 调用PLC生成器生成点表
        success = PLCGenerator.generate_plc_table(
            io_data=io_data,
            output_path=temp_file_path,
            root_window=root_window
        )
        
        if not success:
            return False
        
        # 上传到简道云
        upload_success = upload_file(temp_file_path, "PLC点表")
        
        if upload_success:
            messagebox.showinfo("成功", "PLC点表已成功生成并上传到简道云!")
            return True
        else:
            messagebox.showerror("错误", "上传PLC点表到简道云失败!")
            return False
            
    except Exception as e:
        error_details = traceback.format_exc()
        messagebox.showerror("错误", f"生成或上传PLC点表时发生错误:\n{str(e)}\n\n详细错误信息:\n{error_details}")
        return False


def upload_fat_table(io_data, root_window=None):
    """
    生成并上传FAT点表到简道云
    
    参数:
        io_data (pandas.DataFrame): IO数据
        root_window (tk.Tk): 主窗口对象，用于创建进度窗口
        
    返回:
        bool: 上传成功返回True，失败返回False
    """
    try:
        # 导入FAT生成器
        from FAT_generator import FATGenerator
        import tempfile
        
        # 创建临时文件
        temp_dir = tempfile.gettempdir()
        temp_file_path = os.path.join(temp_dir, "FAT点表.xls")
        
        # 调用FAT生成器生成点表
        success = FATGenerator.generate_fat_table(
            io_data=io_data,
            output_path=temp_file_path,
            root_window=root_window
        )
        
        if not success:
            return False
        
        # 上传到简道云
        upload_success = upload_file(temp_file_path, "FAT点表")
        
        if upload_success:
            messagebox.showinfo("成功", "FAT点表已成功生成并上传到简道云!")
            return True
        else:
            messagebox.showerror("错误", "上传FAT点表到简道云失败!")
            return False
            
    except Exception as e:
        error_details = traceback.format_exc()
        messagebox.showerror("错误", f"生成或上传FAT点表时发生错误:\n{str(e)}\n\n详细错误信息:\n{error_details}")
        return False 