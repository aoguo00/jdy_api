#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
HMI点表生成模块 - 用于生成亚控HMI点表

此模块负责将IO数据转换为亚控HMI点表格式
实现了HMI点表生成的核心逻辑，支持BOOL类型和REAL类型数据
提供进度显示和错误处理功能
"""

import os
import traceback
from tkinter import messagebox, Toplevel, ttk
from pathlib import Path
from config.settings import TEMPLATE_DIR, HMI_TEMPLATE
import pandas as pd

# 添加数据词典模板配置常量
DATA_DICTIONARY_TEMPLATE = "数据词典点表模板.xls"

class HMIGenerator:
    """
    HMI点表生成器类
    
    负责将IO数据转换为亚控HMI点表格式
    提供静态方法用于生成不同类型(BOOL/REAL)的HMI点表
    支持UI进度显示和异常处理
    """
    
    # 定义需要处理的扩展点位列表
    EXTENDED_POINTS = [
        {"name": "SLL设定点位", "comm_addr": "SLL设定点位_通讯地址", "suffix": "_LoLoLimit", "is_bool": False},
        {"name": "SL设定点位", "comm_addr": "SL设定点位_通讯地址", "suffix": "_LoLimit", "is_bool": False},
        {"name": "SH设定点位", "comm_addr": "SH设定点位_通讯地址", "suffix": "_HiLimit", "is_bool": False},
        {"name": "SHH设定点位", "comm_addr": "SHH设定点位_通讯地址", "suffix": "_HiHiLimit", "is_bool": False},
        {"name": "LL报警", "comm_addr": "LL报警_通讯地址", "suffix": "_LL", "is_bool": True},
        {"name": "L报警", "comm_addr": "L报警_通讯地址", "suffix": "_L", "is_bool": True},
        {"name": "H报警", "comm_addr": "H报警_通讯地址", "suffix": "_H", "is_bool": True},
        {"name": "HH报警", "comm_addr": "HH报警_通讯地址", "suffix": "_HH", "is_bool": True},
        {"name": "维护值设定点位", "comm_addr": "维护值设定点位_通讯地址", "suffix": "_whz", "is_bool": False},
        {"name": "维护使能开关点位", "comm_addr": "维护使能开关点位_通讯地址", "suffix": "_MAIN_EN", "is_bool": True}
    ]
    
    # 亚控HMI点表默认字段值
    @classmethod
    def get_default_bool_field_values(cls):
        """获取布尔类型点位的默认字段值"""
        return {
            "TagType": "用户变量",
            "TagDataType": "IODisc",
            "ChannelName": "Network1",
            "ChannelDriver": "ModbusMaster",
            "DeviceSeries": "ModbusTCP",
            "CollectInterval": 1000
        }
        
    @classmethod
    def get_default_real_field_values(cls):
        """获取实数类型点位的默认字段值"""
        return {
            "TagType": "用户变量",
            "TagDataType": "IOFloat",
            "ChannelName": "Network1",
            "ChannelDriver": "ModbusMaster",
            "DeviceSeries": "ModbusTCP",
            "CollectInterval": 1000
        }
    
    @staticmethod
    def generate_hmi_table(io_data, output_path, root_window=None):
        """
        生成HMI点表(BOOL类型数据)
        
        处理IO数据并生成符合亚控HMI要求的布尔类型点表Excel文件
        支持进度显示窗口，提供用户友好的导出体验
        
        Args:
            io_data: 上传的IO点表数据(DataFrame)
            output_path: 输出文件路径
            root_window: 父窗口，用于显示进度窗口
            
        Returns:
            bool: 操作是否成功
        """
        # 导入xlrd和xlwt库，确保它们在这个方法中可用
        import xlrd 
        import xlwt
        
        try:
            # 显示导出进度窗口
            export_window = None
            if root_window:
                export_window = Toplevel(root_window)
                export_window.title("正在导出HMI点表")
                export_window.geometry("300x100")
                export_window.transient(root_window)
                export_window.grab_set()
                
                # 设置窗口在主窗口中央显示
                export_window.withdraw()  # 先隐藏窗口
                export_window.update()    # 更新窗口信息
                
                # 计算窗口位置
                x = root_window.winfo_x() + (root_window.winfo_width() - 300) // 2
                y = root_window.winfo_y() + (root_window.winfo_height() - 100) // 2
                export_window.geometry(f"300x100+{x}+{y}")
                
                export_window.deiconify()  # 显示窗口
                
                export_label = ttk.Label(export_window, text="正在生成HMI点表，请稍候...", font=("Microsoft YaHei", 10))
                export_label.pack(pady=20)
                export_window.update()
            
            # 从上传的IO点表中筛选出BOOL类型的点位
            bool_df = io_data[io_data["数据类型"] == "BOOL"].copy()
            
            # 如果没有BOOL类型数据，显示警告
            if len(bool_df) == 0:
                if export_window:
                    export_window.destroy()
                messagebox.showwarning("警告", "上传的IO点表中没有BOOL类型的数据点！")
                return False
            
            # 检查本地模板文件是否存在
            template_file = os.path.join(TEMPLATE_DIR, HMI_TEMPLATE)
            if not os.path.exists(template_file):
                if export_window:
                    export_window.destroy()
                messagebox.showwarning("警告", f"找不到模板文件: {template_file}，请确保该文件在 {TEMPLATE_DIR} 目录下！")
                return False
            
            # 定义输出路径，只使用.xls格式
            xls_output_path = str(Path(output_path).with_suffix('.xls'))
            
            try:
                # 确保目标文件不存在
                if os.path.exists(xls_output_path):
                    os.remove(xls_output_path)
                
                # 读取模板获取结构
                template_workbook = xlrd.open_workbook(template_file)
                
                # 创建新的工作簿
                workbook = xlwt.Workbook(encoding='utf-8')
                
                # 创建宋体字体样式，大小为10
                font = xlwt.Font()
                font.name = '宋体'
                font.height = 20 * 10  # 10号字体对应的高度是200
                
                # 设置文本格式
                text_style = xlwt.XFStyle()
                text_style.num_format_str = '@'
                text_style.font = font
                
                # 设置标准单元格样式（非文本格式）
                standard_style = xlwt.XFStyle()
                standard_style.font = font
                
                # 设置数字格式样式
                number_style = xlwt.XFStyle()
                number_style.font = font
                number_style.num_format_str = '0'  # 整数格式
                
                # 设置浮点数格式样式
                float_style = xlwt.XFStyle()
                float_style.font = font
                float_style.num_format_str = '0.000000'  # 浮点数格式
                
                # 首先复制模板中的所有工作表
                disc_sheet_idx = -1
                float_sheet_idx = -1
                
                # 第一遍循环，复制所有工作表并获取IO_DISC和IO_FLOAT的索引
                for idx in range(template_workbook.nsheets):
                    template_sheet = template_workbook.sheet_by_index(idx)
                    sheet_name = template_sheet.name
                    
                    # 复制工作表
                    new_sheet = workbook.add_sheet(sheet_name)
                    
                    # 复制表头和所有内容
                    for row in range(template_sheet.nrows):
                        for col in range(template_sheet.ncols):
                            value = template_sheet.cell_value(row, col)
                            # 使用宋体字体样式写入单元格
                            new_sheet.write(row, col, value, standard_style)
                    
                    # 记录特殊工作表的索引
                    if sheet_name == "IO_DISC":
                        disc_sheet_idx = idx
                    elif sheet_name == "IO_FLOAT":
                        float_sheet_idx = idx
                
                # 检查是否找到了IO_DISC和IO_FLOAT工作表
                if disc_sheet_idx == -1:
                    raise ValueError("模板文件中没有找到IO_DISC工作表")
                
                # 重新获取工作表引用
                disc_sheet = workbook.get_sheet(disc_sheet_idx)
                
                # 获取模板中IO_DISC工作表
                template_disc_sheet = template_workbook.sheet_by_index(disc_sheet_idx)
                
                # 查找表中的所有列索引和列名
                column_indices = {}
                for col in range(template_disc_sheet.ncols):
                    header = template_disc_sheet.cell_value(0, col)
                    if header:
                        column_indices[header] = col
                
                # 设置IO_DISC工作簿的固定值
                disc_fixed_values = {
                    "TagType": "用户变量",
                    "TagDataType": "IODisc",
                    "ChannelName": "Network1",
                    "ChannelDriver": "ModbusMaster",
                    "DeviceSeries": "ModbusTCP",
                    "DeviceSeriesType": "0",
                    "CollectControl": "否",
                    "CollectInterval": 1000,  # 使用数字而非字符串
                    "CollectOffset": 0,  # 使用数字而非字符串
                    "TimeZoneBias": 0,  # 使用数字而非字符串
                    "TimeAdjustment": 0,  # 使用数字而非字符串
                    "Enable": "是",
                    "ForceWrite": "否",
                    "RegName": 0,  # 使用数字而非字符串
                    "RegType": 0,  # 使用数字而非字符串
                    "ItemDataType": "BIT",
                    "ItemAccessMode": "读写",
                    "HisRecordMode": "不记录",
                    "HisDeadBand": 0.000000,  # 使用数字而非字符串
                    "HisInterval": 60  # 使用数字而非字符串
                }
                
                # 数字字段列表
                number_fields = ["TagID", "CollectInterval", "CollectOffset", "TimeZoneBias", "TimeAdjustment", 
                                "RegName", "RegType", "HisDeadBand", "HisInterval"]
                
                # 设置从第二行开始填充数据（表头是第一行）
                disc_row_start = 1
                
                # 填充BOOL数据 - 从表头后的第二行开始添加
                for i, (_, row) in enumerate(bool_df.iterrows()):
                    # 获取变量信息
                    hmi_name = row.get("变量名称（HMI）", "")
                    description = row.get("变量描述", "")
                    station_name = row.get("场站名", "未知站点")
                    comm_address = row.get("上位机通讯地址", "")
                    channel_code = row.get("通道位号", "")
                    
                    # 如果变量名为空，则自动补全
                    if pd.isna(hmi_name) or str(hmi_name).strip() == "":
                        hmi_name = f"YLDW{channel_code}"
                        description = f"预留点位{channel_code}" if pd.isna(description) or str(description).strip() == "" else description
                    
                    # 填充数据到Excel - 行索引从表头后的第二行开始添加
                    excel_row = disc_row_start + i
                    
                    # 先填充必要的字段
                    if "TagID" in column_indices:
                        disc_sheet.write(excel_row, column_indices["TagID"], excel_row, number_style)
                    if "TagName" in column_indices:
                        disc_sheet.write(excel_row, column_indices["TagName"], hmi_name, standard_style)
                    if "Description" in column_indices:
                        disc_sheet.write(excel_row, column_indices["Description"], description, standard_style)
                    if "DeviceName" in column_indices:
                        disc_sheet.write(excel_row, column_indices["DeviceName"], station_name, standard_style)
                    if "TagGroup" in column_indices:
                        disc_sheet.write(excel_row, column_indices["TagGroup"], station_name, standard_style)
                    if "ItemName" in column_indices:
                        # ItemName = 上位机通讯地址前面部分，使用文本格式
                        try:
                            item_name_value = str(comm_address).split('.')[0]
                            # 将通讯地址转换为整数（去掉可能的前导0）并再次转为字符串
                            item_name_int = int(item_name_value)
                            disc_sheet.write(excel_row, column_indices["ItemName"], str(item_name_int), text_style)
                        except (ValueError, TypeError):
                            # 如果无法转换为整数，则保持原样使用文本格式
                            disc_sheet.write(excel_row, column_indices["ItemName"], item_name_value, text_style)
                    
                    # 填充固定值字段
                    for field, value in disc_fixed_values.items():
                        if field in column_indices:
                            # 根据字段类型选择样式
                            if field in number_fields:
                                disc_sheet.write(excel_row, column_indices[field], value, number_style)
                            else:
                                disc_sheet.write(excel_row, column_indices[field], value, standard_style)
                
                # 处理布尔类型的扩展点位
                bool_ext_row_counter = disc_row_start + len(bool_df)  # 从基本点位后开始添加
                bool_ext_id_counter = disc_row_start + len(bool_df) - 1  # 从最后一个ID开始递增
                
                # 遍历所有数据行
                for _, row in io_data.iterrows():
                    # 获取基础信息
                    base_hmi_name = row.get("变量名称（HMI）", "")
                    base_description = row.get("变量描述", "")
                    station_name = row.get("场站名", "未知站点")
                    channel_code = row.get("通道位号", "")
                    
                    # 如果变量名为空，则自动补全
                    if pd.isna(base_hmi_name) or str(base_hmi_name).strip() == "":
                        base_hmi_name = f"YLDW{channel_code}"
                        base_description = f"预留点位{channel_code}" if pd.isna(base_description) or str(base_description).strip() == "" else base_description
                    
                    # 处理该行的BOOL类型扩展点位
                    for ext_point in HMIGenerator.EXTENDED_POINTS:
                        # 只处理BOOL类型点位
                        if not ext_point["is_bool"]:
                            continue
                            
                        point_name = ext_point["name"]
                        comm_addr_field = ext_point["comm_addr"]
                        point_suffix = ext_point["suffix"]
                        
                        # 获取扩展点位的值和通讯地址
                        point_value = row.get(point_name, "")
                        point_comm_addr = row.get(comm_addr_field, "")
                        
                        # 如果扩展点位值为空或"/"，则跳过
                        if not point_value or point_value == "/":
                            continue
                        
                        # 如果通讯地址为空或"/"，则跳过
                        if not point_comm_addr or point_comm_addr == "/":
                            continue
                        
                        # 为扩展点位创建变量名和描述
                        ext_hmi_name = str(base_hmi_name) + point_suffix
                        ext_description = str(base_description) + "_" + point_name
                        
                        # 填充扩展点位数据
                        bool_ext_id_counter += 1
                        
                        if "TagID" in column_indices:
                            disc_sheet.write(bool_ext_row_counter, column_indices["TagID"], bool_ext_id_counter, number_style)
                        if "TagName" in column_indices:
                            disc_sheet.write(bool_ext_row_counter, column_indices["TagName"], ext_hmi_name, standard_style)
                        if "Description" in column_indices:
                            disc_sheet.write(bool_ext_row_counter, column_indices["Description"], ext_description, standard_style)
                        if "DeviceName" in column_indices:
                            disc_sheet.write(bool_ext_row_counter, column_indices["DeviceName"], station_name, standard_style)
                        if "TagGroup" in column_indices:
                            disc_sheet.write(bool_ext_row_counter, column_indices["TagGroup"], station_name, standard_style)
                        if "ItemName" in column_indices:
                            # ItemName = 上位机通讯地址，使用文本格式
                            try:
                                item_name_value = str(point_comm_addr).split('.')[0]
                                item_name_int = int(item_name_value)
                                disc_sheet.write(bool_ext_row_counter, column_indices["ItemName"], str(item_name_int), text_style)
                            except (ValueError, TypeError):
                                disc_sheet.write(bool_ext_row_counter, column_indices["ItemName"], item_name_value, text_style)
                        
                        # 填充固定值字段
                        for field, value in disc_fixed_values.items():
                            if field in column_indices:
                                if field in number_fields:
                                    disc_sheet.write(bool_ext_row_counter, column_indices[field], value, number_style)
                                else:
                                    disc_sheet.write(bool_ext_row_counter, column_indices[field], value, standard_style)
                        
                        # 递增行计数器
                        bool_ext_row_counter += 1
                
                # 处理REAL类型数据
                real_df = io_data[io_data["数据类型"] == "REAL"].copy()
                if len(real_df) > 0 and float_sheet_idx != -1:
                    # 获取FLOAT工作表的引用
                    float_sheet = workbook.get_sheet(float_sheet_idx)
                    
                    # 获取模板中的IO_FLOAT工作表
                    template_float_sheet = template_workbook.sheet_by_index(float_sheet_idx)
                    
                    # 查找FLOAT表中的所有列索引和列名
                    float_column_indices = {}
                    for col in range(template_float_sheet.ncols):
                        header = template_float_sheet.cell_value(0, col)
                        if header:
                            float_column_indices[header] = col
                    
                    # 设置IO_FLOAT工作簿的固定值
                    float_fixed_values = {
                        "TagType": "用户变量",
                        "TagDataType": "IOFloat",
                        "MaxRawValue": 1000000000.000000,  # 使用数字而非字符串
                        "MinRawValue": -1000000000.000000,  # 使用数字而非字符串
                        "MaxValue": 1000000000.000000,  # 使用数字而非字符串
                        "MinValue": -1000000000.000000,  # 使用数字而非字符串
                        "ConvertType": "无",
                        "IsFilter": "否",
                        "DeadBand": 0,  # 使用数字而非字符串
                        "ChannelName": "Network1",
                        "ChannelDriver": "ModbusMaster",
                        "DeviceSeries": "ModbusTCP",
                        "DeviceSeriesType": 0,  # 使用数字而非字符串
                        "CollectControl": "否",
                        "CollectInterval": 1000,  # 使用数字而非字符串
                        "CollectOffset": 0,  # 使用数字而非字符串
                        "TimeZoneBias": 0,  # 使用数字而非字符串
                        "TimeAdjustment": 0,  # 使用数字而非字符串
                        "Enable": "是",
                        "ForceWrite": "否",
                        "RegName": 4,  # 使用数字而非字符串
                        "RegType": 3,  # 使用数字而非字符串
                        "ItemDataType": "FLOAT",
                        "ItemAccessMode": "读写",
                        "HisRecordMode": "不记录",
                        "HisDeadBand": 0.000000,  # 使用数字而非字符串
                        "HisInterval": 60  # 使用数字而非字符串
                    }
                    
                    # 浮点数字段列表
                    float_fields = ["HiHiLimit", "HiLimit", "LoLimit", "LoLoLimit"]
                    
                    # 设置从第二行开始填充数据（表头是第一行）
                    float_row_start = 1
                    
                    # BOOL ID计数器值
                    last_disc_id = bool_ext_id_counter
                    
                    # 填充REAL数据
                    start_id = last_disc_id + 1  # 从DISC表的最后ID+1开始
                    
                    # 行计数器，从表头后开始
                    excel_row_counter = float_row_start  # 从第二行开始添加
                    current_id_counter = start_id  # ID计数器从last_disc_id+1开始
                    
                    for i, (_, row) in enumerate(real_df.iterrows()):
                        # 获取变量信息
                        hmi_name = row.get("变量名称（HMI）", "")
                        description = row.get("变量描述", "")
                        station_name = row.get("场站名", "未知站点")
                        comm_address = row.get("上位机通讯地址", "")
                        channel_code = row.get("通道位号", "")
                        
                        # 如果变量名为空，则自动补全
                        if pd.isna(hmi_name) or str(hmi_name).strip() == "":
                            hmi_name = f"YLDW{channel_code}"
                            description = f"预留点位{channel_code}" if pd.isna(description) or str(description).strip() == "" else description
                        
                        # 当前ID
                        current_id = current_id_counter
                        current_id_counter += 1
                        
                        # 填充数据到Excel - 使用行计数器
                        if "TagID" in float_column_indices:
                            float_sheet.write(excel_row_counter, float_column_indices["TagID"], current_id, number_style)
                        if "TagName" in float_column_indices:
                            float_sheet.write(excel_row_counter, float_column_indices["TagName"], hmi_name, standard_style)
                        if "Description" in float_column_indices:
                            float_sheet.write(excel_row_counter, float_column_indices["Description"], description, standard_style)
                        if "DeviceName" in float_column_indices:
                            float_sheet.write(excel_row_counter, float_column_indices["DeviceName"], station_name, standard_style)
                        if "TagGroup" in float_column_indices:
                            float_sheet.write(excel_row_counter, float_column_indices["TagGroup"], station_name, standard_style)
                        if "ItemName" in float_column_indices:
                            # 确保值为文本格式
                            try:
                                item_name_value = str(comm_address).split('.')[0]
                                item_name_int = int(item_name_value)
                                float_sheet.write(excel_row_counter, float_column_indices["ItemName"], str(item_name_int), text_style)
                            except (ValueError, TypeError):
                                float_sheet.write(excel_row_counter, float_column_indices["ItemName"], item_name_value, text_style)
                        
                        # 填充固定值字段
                        for field, value in float_fixed_values.items():
                            if field in float_column_indices:
                                if field in float_fields:
                                    float_sheet.write(excel_row_counter, float_column_indices[field], value, float_style)
                                elif field in number_fields:
                                    float_sheet.write(excel_row_counter, float_column_indices[field], value, number_style)
                                else:
                                    float_sheet.write(excel_row_counter, float_column_indices[field], value, standard_style)
                        
                        # 递增行计数器
                        excel_row_counter += 1
                
                # 处理REAL类型的扩展点位
                real_ext_row_counter = float_row_start + len(real_df)  # 从基本点位后开始添加
                real_ext_id_counter = last_disc_id + len(real_df)  # 从最后一个ID计数
                
                # 遍历所有数据行
                for _, row in io_data.iterrows():
                    # 获取基础信息
                    base_hmi_name = row.get("变量名称（HMI）", "")
                    base_description = row.get("变量描述", "")
                    station_name = row.get("场站名", "未知站点")
                    channel_code = row.get("通道位号", "")
                    
                    # 如果变量名为空，则自动补全
                    if pd.isna(base_hmi_name) or str(base_hmi_name).strip() == "":
                        base_hmi_name = f"YLDW{channel_code}"
                        base_description = f"预留点位{channel_code}" if pd.isna(base_description) or str(base_description).strip() == "" else base_description
                    
                    # 处理该行的REAL类型扩展点位
                    for ext_point in HMIGenerator.EXTENDED_POINTS:
                        # 只处理REAL类型点位
                        if ext_point["is_bool"]:
                            continue
                            
                        point_name = ext_point["name"]
                        comm_addr_field = ext_point["comm_addr"]
                        point_suffix = ext_point["suffix"]
                        
                        # 获取扩展点位的值和通讯地址
                        point_value = row.get(point_name, "")
                        point_comm_addr = row.get(comm_addr_field, "")
                        
                        # 如果扩展点位值为空或"/"或None，则跳过
                        if pd.isna(point_value) or not point_value or point_value == "/":
                            continue
                        
                        # 如果通讯地址为空或"/"或None，则跳过
                        if pd.isna(point_comm_addr) or not point_comm_addr or point_comm_addr == "/":
                            continue
                        
                        # 为扩展点位创建变量名和描述
                        ext_hmi_name = str(base_hmi_name) + point_suffix
                        ext_description = str(base_description) + "_" + point_name
                        
                        # 填充扩展点位数据
                        real_ext_id_counter += 1
                        
                        if "TagID" in float_column_indices:
                            float_sheet.write(real_ext_row_counter, float_column_indices["TagID"], real_ext_id_counter, number_style)
                        if "TagName" in float_column_indices:
                            float_sheet.write(real_ext_row_counter, float_column_indices["TagName"], ext_hmi_name, standard_style)
                        if "Description" in float_column_indices:
                            float_sheet.write(real_ext_row_counter, float_column_indices["Description"], ext_description, standard_style)
                        if "DeviceName" in float_column_indices:
                            float_sheet.write(real_ext_row_counter, float_column_indices["DeviceName"], station_name, standard_style)
                        if "TagGroup" in float_column_indices:
                            float_sheet.write(real_ext_row_counter, float_column_indices["TagGroup"], station_name, standard_style)
                        if "ItemName" in float_column_indices:
                            # 确保值为文本格式
                            try:
                                item_name_value = str(point_comm_addr).split('.')[0]
                                item_name_int = int(item_name_value)
                                float_sheet.write(real_ext_row_counter, float_column_indices["ItemName"], str(item_name_int), text_style)
                            except (ValueError, TypeError):
                                float_sheet.write(real_ext_row_counter, float_column_indices["ItemName"], item_name_value, text_style)
                        
                        # 填充固定值字段
                        for field, value in float_fixed_values.items():
                            if field in float_column_indices:
                                if field in float_fields:
                                    float_sheet.write(real_ext_row_counter, float_column_indices[field], value, float_style)
                                elif field in number_fields:
                                    float_sheet.write(real_ext_row_counter, float_column_indices[field], value, number_style)
                                else:
                                    float_sheet.write(real_ext_row_counter, float_column_indices[field], value, standard_style)
                        
                        # 递增行计数器
                        real_ext_row_counter += 1
                
                # 保存工作簿
                workbook.save(xls_output_path)
                
                # 检查文件是否生成成功
                if not (os.path.exists(xls_output_path) and os.path.getsize(xls_output_path) > 0):
                    raise ValueError(f"生成的文件不存在或为空: {xls_output_path}")
                
                # 不再自动打开文件，由UI层负责显示消息和打开文件
                
                # 关闭导出进度窗口
                if export_window:
                    export_window.destroy()
                
                return True
                
            except Exception as e:
                error_msg = f"生成HMI REAL点表失败: {str(e)}\n{traceback.format_exc()}"
                messagebox.showerror("错误", error_msg)
                
                if export_window:
                    export_window.destroy()
                return False
                
        except Exception as e:
            if export_window and export_window.winfo_exists():
                export_window.destroy()
            error_details = traceback.format_exc()
            messagebox.showerror("错误", f"生成HMI REAL点表时发生错误:\n{str(e)}\n\n详细错误信息:\n{error_details}")
            return False
    
    @staticmethod
    def generate_io_real(io_data, output_path, root_window=None):
        """
        生成IO_REAL工作簿（用于模拟量点位）
        
        Args:
            io_data: 上传的IO点表数据(DataFrame)
            output_path: 输出文件路径
            root_window: 父窗口，用于显示进度窗口
            
        Returns:
            bool: 操作是否成功
        """
        # 导入xlrd和xlwt库，确保它们在这个方法中可用
        import xlrd
        import xlwt
        
        try:
            # 显示导出进度窗口
            export_window = None
            if root_window:
                export_window = Toplevel(root_window)
                export_window.title("正在导出HMI点表")
                export_window.geometry("300x100")
                export_window.transient(root_window)
                export_window.grab_set()
                
                # 设置窗口在主窗口中央显示
                export_window.withdraw()  # 先隐藏窗口
                export_window.update()    # 更新窗口信息
                
                # 计算窗口位置
                x = root_window.winfo_x() + (root_window.winfo_width() - 300) // 2
                y = root_window.winfo_y() + (root_window.winfo_height() - 100) // 2
                export_window.geometry(f"300x100+{x}+{y}")
                
                export_window.deiconify()  # 显示窗口
                
                export_label = ttk.Label(export_window, text="正在生成HMI REAL点表，请稍候...", font=("Microsoft YaHei", 10))
                export_label.pack(pady=20)
                export_window.update()
            
            # 从上传的IO点表中筛选出REAL类型的点位
            real_df = io_data[io_data["数据类型"] == "REAL"].copy()
            
            # 如果没有REAL类型数据，显示警告
            if len(real_df) == 0:
                if export_window:
                    export_window.destroy()
                messagebox.showwarning("警告", "上传的IO点表中没有REAL类型的数据点！")
                return False
            
            # 检查本地模板文件是否存在
            template_file = os.path.join(TEMPLATE_DIR, HMI_TEMPLATE)
            if not os.path.exists(template_file):
                if export_window:
                    export_window.destroy()
                messagebox.showwarning("警告", f"找不到模板文件: {template_file}，请确保该文件在 {TEMPLATE_DIR} 目录下！")
                return False
            
            # 定义输出路径，只使用.xls格式
            xls_output_path = str(Path(output_path).with_suffix('.xls'))
            
            try:
                # 检查输出文件是否存在，不存在则创建新文件
                if os.path.exists(xls_output_path):
                    os.remove(xls_output_path)
                
                # 读取模板获取结构
                template_workbook = xlrd.open_workbook(template_file)
                
                # 创建新的工作簿
                workbook = xlwt.Workbook(encoding='utf-8')
                
                # 创建宋体字体样式，大小为10
                font = xlwt.Font()
                font.name = '宋体'
                font.height = 20 * 10  # 10号字体对应的高度是200
                
                # 设置文本格式
                text_style = xlwt.XFStyle()
                text_style.num_format_str = '@'
                text_style.font = font
                
                # 设置标准单元格样式（非文本格式）
                standard_style = xlwt.XFStyle()
                standard_style.font = font
                
                # 设置数字格式样式
                number_style = xlwt.XFStyle()
                number_style.font = font
                number_style.num_format_str = '0'  # 整数格式
                
                # 设置浮点数格式样式
                float_style = xlwt.XFStyle()
                float_style.font = font
                float_style.num_format_str = '0.000000'  # 浮点数格式
                
                # 首先复制模板中的所有工作表
                float_sheet_idx = -1
                
                # 第一遍循环，复制所有工作表并获取IO_FLOAT的索引
                for idx in range(template_workbook.nsheets):
                    template_sheet = template_workbook.sheet_by_index(idx)
                    sheet_name = template_sheet.name
                    
                    # 复制工作表
                    new_sheet = workbook.add_sheet(sheet_name)
                    
                    # 复制表头和所有内容
                    for row in range(template_sheet.nrows):
                        for col in range(template_sheet.ncols):
                            value = template_sheet.cell_value(row, col)
                            # 使用宋体字体样式写入单元格
                            new_sheet.write(row, col, value, standard_style)
                    
                    # 记录IO_FLOAT工作表的索引
                    if sheet_name == "IO_FLOAT":
                        float_sheet_idx = idx
                
                # 检查是否找到了IO_FLOAT工作表
                if float_sheet_idx == -1:
                    raise ValueError("模板文件中没有找到IO_FLOAT工作表")
                
                # 查找IO_DISC工作表以获取TagID计数
                last_disc_id = 0
                disc_sheet_idx = -1
                
                for idx in range(template_workbook.nsheets):
                    template_sheet = template_workbook.sheet_by_index(idx)
                    if template_sheet.name == "IO_DISC":
                        disc_sheet_idx = idx
                        # 有数据行则计算最后的ID
                        if template_sheet.nrows > 1:
                            # 查找TagID列
                            for col in range(template_sheet.ncols):
                                if template_sheet.cell_value(0, col) == "TagID":
                                    if template_sheet.nrows > 1:
                                        try:
                                            last_id_value = template_sheet.cell_value(template_sheet.nrows-1, col)
                                            last_disc_id = int(last_id_value)
                                        except (ValueError, TypeError):
                                            pass
                                    break
                        break
                
                # 获取工作表引用和模板信息
                float_sheet = workbook.get_sheet(float_sheet_idx)
                template_float_sheet = template_workbook.sheet_by_index(float_sheet_idx)
                
                # 查找FLOAT表中的所有列索引和列名
                column_indices = {}
                for col in range(template_float_sheet.ncols):
                    header = template_float_sheet.cell_value(0, col)
                    if header:
                        column_indices[header] = col
                
                # 设置IO_FLOAT工作簿的固定值
                float_fixed_values = {
                    "TagType": "用户变量",
                    "TagDataType": "IOFloat",
                    "MaxRawValue": 1000000000.000000,  # 使用数字而非字符串
                    "MinRawValue": -1000000000.000000,  # 使用数字而非字符串
                    "MaxValue": 1000000000.000000,  # 使用数字而非字符串
                    "MinValue": -1000000000.000000,  # 使用数字而非字符串
                    "ConvertType": "无",
                    "IsFilter": "否",
                    "DeadBand": 0,  # 使用数字而非字符串
                    "ChannelName": "Network1",
                    "ChannelDriver": "ModbusMaster",
                    "DeviceSeries": "ModbusTCP",
                    "DeviceSeriesType": 0,  # 使用数字而非字符串
                    "CollectControl": "否",
                    "CollectInterval": 1000,  # 使用数字而非字符串
                    "CollectOffset": 0,  # 使用数字而非字符串
                    "TimeZoneBias": 0,  # 使用数字而非字符串
                    "TimeAdjustment": 0,  # 使用数字而非字符串
                    "Enable": "是",
                    "ForceWrite": "否",
                    "RegName": 4,  # 使用数字而非字符串
                    "RegType": 3,  # 使用数字而非字符串
                    "ItemDataType": "FLOAT",
                    "ItemAccessMode": "读写",
                    "HisRecordMode": "不记录",
                    "HisDeadBand": 0.000000,  # 使用数字而非字符串
                    "HisInterval": 60  # 使用数字而非字符串
                }
                
                # 数字字段列表
                number_fields = ["TagID", "CollectInterval", "CollectOffset", "TimeZoneBias", "TimeAdjustment", 
                                "RegName", "RegType", "DeviceSeriesType", "HisInterval"]
                
                # 浮点数字段列表
                float_fields = ["HiHiLimit", "HiLimit", "LoLimit", "LoLoLimit"]
                
                # 设置从第二行开始填充数据（表头是第一行）
                float_row_start = 1
                
                # 填充REAL数据
                start_id = last_disc_id + 1  # 从DISC表的最后ID+1开始
                
                # 行计数器，从表头后开始
                excel_row_counter = float_row_start  # 从第二行开始添加
                current_id_counter = start_id  # ID计数器从last_disc_id+1开始
                
                for i, (_, row) in enumerate(real_df.iterrows()):
                    # 获取变量信息
                    hmi_name = row.get("变量名称（HMI）", "")
                    description = row.get("变量描述", "")
                    station_name = row.get("场站名", "未知站点")
                    comm_address = row.get("上位机通讯地址", "")
                    channel_code = row.get("通道位号", "")
                    
                    # 如果变量名为空，则自动补全
                    if pd.isna(hmi_name) or str(hmi_name).strip() == "":
                        hmi_name = f"YLDW{channel_code}"
                        description = f"预留点位{channel_code}" if pd.isna(description) or str(description).strip() == "" else description
                    
                    # 当前ID
                    current_id = current_id_counter
                    current_id_counter += 1
                    
                    # 填充数据到Excel - 使用行计数器
                    if "TagID" in column_indices:
                        float_sheet.write(excel_row_counter, column_indices["TagID"], current_id, number_style)
                    if "TagName" in column_indices:
                        float_sheet.write(excel_row_counter, column_indices["TagName"], hmi_name, standard_style)
                    if "Description" in column_indices:
                        float_sheet.write(excel_row_counter, column_indices["Description"], description, standard_style)
                    if "DeviceName" in column_indices:
                        float_sheet.write(excel_row_counter, column_indices["DeviceName"], station_name, standard_style)
                    if "TagGroup" in column_indices:
                        float_sheet.write(excel_row_counter, column_indices["TagGroup"], station_name, standard_style)
                    if "ItemName" in column_indices:
                        # 确保值为文本格式
                        try:
                            item_name_value = str(comm_address).split('.')[0]
                            item_name_int = int(item_name_value)
                            float_sheet.write(excel_row_counter, column_indices["ItemName"], str(item_name_int), text_style)
                        except (ValueError, TypeError):
                            float_sheet.write(excel_row_counter, column_indices["ItemName"], item_name_value, text_style)
                    
                    # 填充固定值字段
                    for field, value in float_fixed_values.items():
                        if field in column_indices:
                            if field in float_fields:
                                float_sheet.write(excel_row_counter, column_indices[field], value, float_style)
                            elif field in number_fields:
                                float_sheet.write(excel_row_counter, column_indices[field], value, number_style)
                            else:
                                float_sheet.write(excel_row_counter, column_indices[field], value, standard_style)
                    
                    # 递增行计数器
                    excel_row_counter += 1
                
                # 处理REAL类型的扩展点位
                for _, row in io_data.iterrows():
                    # 获取基础信息
                    base_hmi_name = row.get("变量名称（HMI）", "")
                    base_description = row.get("变量描述", "")
                    station_name = row.get("场站名", "未知站点")
                    channel_code = row.get("通道位号", "")
                    
                    # 如果变量名为空，则自动补全
                    if pd.isna(base_hmi_name) or str(base_hmi_name).strip() == "":
                        base_hmi_name = f"YLDW{channel_code}"
                        base_description = f"预留点位{channel_code}" if pd.isna(base_description) or str(base_description).strip() == "" else base_description
                    
                    # 处理该行的REAL类型扩展点位
                    for ext_point in HMIGenerator.EXTENDED_POINTS:
                        # 只处理REAL类型点位
                        if ext_point["is_bool"]:
                            continue
                            
                        point_name = ext_point["name"]
                        comm_addr_field = ext_point["comm_addr"]
                        point_suffix = ext_point["suffix"]
                        
                        # 获取扩展点位的值和通讯地址
                        point_value = row.get(point_name, "")
                        point_comm_addr = row.get(comm_addr_field, "")
                        
                        # 如果扩展点位值为空或"/"或None，则跳过
                        if pd.isna(point_value) or not point_value or point_value == "/":
                            continue
                        
                        # 如果通讯地址为空或"/"或None，则跳过
                        if pd.isna(point_comm_addr) or not point_comm_addr or point_comm_addr == "/":
                            continue
                        
                        # 为扩展点位创建变量名和描述
                        ext_hmi_name = str(base_hmi_name) + point_suffix
                        ext_description = str(base_description) + "_" + point_name
                        
                        # 填充扩展点位数据
                        current_id_counter += 1
                        
                        if "TagID" in column_indices:
                            float_sheet.write(excel_row_counter, column_indices["TagID"], current_id, number_style)
                        if "TagName" in column_indices:
                            float_sheet.write(excel_row_counter, column_indices["TagName"], ext_hmi_name, standard_style)
                        if "Description" in column_indices:
                            float_sheet.write(excel_row_counter, column_indices["Description"], ext_description, standard_style)
                        if "DeviceName" in column_indices:
                            float_sheet.write(excel_row_counter, column_indices["DeviceName"], station_name, standard_style)
                        if "TagGroup" in column_indices:
                            float_sheet.write(excel_row_counter, column_indices["TagGroup"], station_name, standard_style)
                        if "ItemName" in column_indices:
                            # 确保值为文本格式
                            try:
                                item_name_value = str(point_comm_addr).split('.')[0]
                                item_name_int = int(item_name_value)
                                float_sheet.write(excel_row_counter, column_indices["ItemName"], str(item_name_int), text_style)
                            except (ValueError, TypeError):
                                float_sheet.write(excel_row_counter, column_indices["ItemName"], item_name_value, text_style)
                        
                        # 填充固定值字段
                        for field, value in float_fixed_values.items():
                            if field in column_indices:
                                if field in float_fields:
                                    float_sheet.write(excel_row_counter, column_indices[field], value, float_style)
                                elif field in number_fields:
                                    float_sheet.write(excel_row_counter, column_indices[field], value, number_style)
                                else:
                                    float_sheet.write(excel_row_counter, column_indices[field], value, standard_style)
                        
                        # 递增行计数器
                        excel_row_counter += 1
                
                # 保存工作簿
                workbook.save(xls_output_path)
                
                # 检查文件是否生成成功
                if not (os.path.exists(xls_output_path) and os.path.getsize(xls_output_path) > 0):
                    raise ValueError(f"生成的文件不存在或为空: {xls_output_path}")
                
                # 不再自动打开文件，由UI层负责显示消息和打开文件
                
                # 关闭导出进度窗口
                if export_window:
                    export_window.destroy()
                
                return True
                
            except Exception as e:
                error_msg = f"生成HMI REAL点表失败: {str(e)}\n{traceback.format_exc()}"
                messagebox.showerror("错误", error_msg)
                
                if export_window:
                    export_window.destroy()
                return False
                
        except Exception as e:
            if export_window and export_window.winfo_exists():
                export_window.destroy()
            error_details = traceback.format_exc()
            messagebox.showerror("错误", f"生成HMI REAL点表时发生错误:\n{str(e)}\n\n详细错误信息:\n{error_details}")
            return False
    
    @staticmethod
    def generate_data_dictionary_table(io_data, output_path, root_window=None):
        """
        生成数据词典点表
        
        Args:
            io_data: 上传的IO点表数据(DataFrame)
            output_path: 输出文件路径
            root_window: 父窗口，用于显示进度窗口
            
        Returns:
            bool: 操作是否成功
        """
        # 导入xlrd和xlwt库，确保它们在这个方法中可用
        import xlrd 
        import xlwt
        
        try:
            # 显示导出进度窗口
            export_window = None
            if root_window:
                export_window = Toplevel(root_window)
                export_window.title("正在导出数据词典点表")
                export_window.geometry("300x100")
                export_window.transient(root_window)
                export_window.grab_set()
                
                # 设置窗口在主窗口中央显示
                export_window.withdraw()  # 先隐藏窗口
                export_window.update()    # 更新窗口信息
                
                # 计算窗口位置
                x = root_window.winfo_x() + (root_window.winfo_width() - 300) // 2
                y = root_window.winfo_y() + (root_window.winfo_height() - 100) // 2
                export_window.geometry(f"300x100+{x}+{y}")
                
                export_window.deiconify()  # 显示窗口
                
                export_label = ttk.Label(export_window, text="正在生成数据词典点表，请稍候...", font=("Microsoft YaHei", 10))
                export_label.pack(pady=20)
                export_window.update()
            
            # 筛选出BOOL和REAL类型数据
            bool_df = io_data[io_data["数据类型"] == "BOOL"].copy()
            real_df = io_data[io_data["数据类型"] == "REAL"].copy()
            
            # 检查本地模板文件是否存在
            template_file = os.path.join(TEMPLATE_DIR, DATA_DICTIONARY_TEMPLATE)
            if not os.path.exists(template_file):
                if export_window:
                    export_window.destroy()
                messagebox.showwarning("警告", f"找不到模板文件: {template_file}，请确保该文件在 {TEMPLATE_DIR} 目录下！")
                return False
            
            # 定义输出路径，只使用.xls格式
            xls_output_path = str(Path(output_path).with_suffix('.xls'))
            
            try:
                # 确保目标文件不存在
                if os.path.exists(xls_output_path):
                    os.remove(xls_output_path)
                
                # 读取模板获取结构
                template_workbook = xlrd.open_workbook(template_file)
                
                # 创建新的工作簿
                workbook = xlwt.Workbook(encoding='utf-8')
                
                # 创建宋体字体样式，大小为10
                font = xlwt.Font()
                font.name = '宋体'
                font.height = 20 * 10  # 10号字体对应的高度是200
                
                # 设置文本格式
                text_style = xlwt.XFStyle()
                text_style.num_format_str = '@'
                text_style.font = font
                
                # 设置标准单元格样式（非文本格式）
                standard_style = xlwt.XFStyle()
                standard_style.font = font
                
                # 设置数字格式样式
                number_style = xlwt.XFStyle()
                number_style.font = font
                number_style.num_format_str = '0'  # 整数格式
                
                # 设置浮点数格式样式
                float_style = xlwt.XFStyle()
                float_style.font = font
                float_style.num_format_str = '0.000000'  # 浮点数格式
                
                # 首先复制模板中的所有工作表
                disc_sheet_idx = -1
                float_sheet_idx = -1
                
                # 第一遍循环，复制所有工作表并获取IO_DISC和IO_FLOAT的索引
                for idx in range(template_workbook.nsheets):
                    template_sheet = template_workbook.sheet_by_index(idx)
                    sheet_name = template_sheet.name
                    
                    # 复制工作表
                    new_sheet = workbook.add_sheet(sheet_name)
                    
                    # 复制表头和所有内容
                    for row in range(template_sheet.nrows):
                        for col in range(template_sheet.ncols):
                            value = template_sheet.cell_value(row, col)
                            # 使用宋体字体样式写入单元格
                            new_sheet.write(row, col, value, standard_style)
                    
                    # 记录特殊工作表的索引
                    if sheet_name == "IO_DISC":
                        disc_sheet_idx = idx
                    elif sheet_name == "IO_FLOAT":
                        float_sheet_idx = idx
                
                # 检查是否找到了IO_DISC和IO_FLOAT工作表
                if disc_sheet_idx == -1 or float_sheet_idx == -1:
                    raise ValueError("模板文件中没有找到IO_DISC或IO_FLOAT工作表")
                
                # 获取工作表引用
                disc_sheet = workbook.get_sheet(disc_sheet_idx)
                float_sheet = workbook.get_sheet(float_sheet_idx)
                
                # 获取模板中IO_DISC工作表
                template_disc_sheet = template_workbook.sheet_by_index(disc_sheet_idx)
                
                # 查找表中的所有列索引和列名
                disc_column_indices = {}
                for col in range(template_disc_sheet.ncols):
                    header = template_disc_sheet.cell_value(0, col)
                    if header:
                        disc_column_indices[header] = col
                
                # 获取模板中IO_FLOAT工作表
                template_float_sheet = template_workbook.sheet_by_index(float_sheet_idx)
                
                # 查找表中的所有列索引和列名
                float_column_indices = {}
                for col in range(template_float_sheet.ncols):
                    header = template_float_sheet.cell_value(0, col)
                    if header:
                        float_column_indices[header] = col
                
                # 设置IO_DISC工作簿的固定值 - 改为使用数字而非字符串
                disc_fixed_values = {
                    "ContainerType": 1,  # 改为数字
                    "InitialValueBool": "false",
                    "SecurityZoneID": "None",
                    "RecordEvent": "false",
                    "SaveValue": "true",
                    "SaveParameter": "true",
                    "AccessByOtherApplication": "false",
                    "ExtentField1": "",
                    "ExtentField2": "",
                    "HisRecMode": 2,  # 改为数字
                    "HisRecInterval": 60,  # 改为数字
                    "AlarmType": 256,  # 改为数字
                    "CloseString": "关闭",
                    "OpenString": "打开",
                    "AlarmDelay": 0,  # 改为数字
                    "DiscInhibitor": "",
                    "ExtentField3": "",
                    "ExtentField4": "",
                    "ExtentField5": "",
                    "ExtentField6": "",
                    "ExtentField7": "",
                    "ExtentField8": "",
                    "CloseToOpen": "关到开",
                    "OpenToClose": "开到关",
                    "StateEnumTable": "",
                    "IOConfigControl": "true",
                    "IOEnable": "true",
                    "ForceRead": "false",
                    "ForceWrite": "false",
                    "DataConvertMode": 1  # 改为数字
                }
                
                # 设置IO_FLOAT工作簿的固定值 - 改为使用数字而非字符串
                float_fixed_values = {
                    "ContainerType": 1,  # 改为数字
                    "MaxValue": 1000000000,  # 改为数字
                    "MinValue": -1000000000,  # 改为数字
                    "InitialValue": 0,  # 改为数字
                    "Sensitivity": 0,  # 改为数字
                    "EngineerUnits": "",
                    "SecurityZoneID": "None",
                    "RecordEvent": "false",
                    "SaveValue": "true",
                    "SaveParameter": "true",
                    "AccessByOtherApplication": "false",
                    "ExtentField1": "",
                    "ExtentField2": "",
                    "HisRecMode": 2,  # 改为数字
                    "HisRecChangeDeadband": 0,  # 改为数字
                    "HisRecInterval": 60,  # 改为数字
                    "HiHiText": "高高",
                    "HiHiPriority": 1,  # 改为数字
                    "HiHiInhibitor": "",
                    "HiText": "高",
                    "HiPriority": 1,  # 改为数字
                    "HiInhibitor": "",
                    "LoText": "低",
                    "LoPriority": 1,  # 改为数字
                    "LoInhibitor": "",
                    "LoLoText": "低低",
                    "LoLoPriority": 1,  # 改为数字
                    "LoLoInhibitor": "",
                    "LimitDeadband": 0,  # 改为数字
                    "LimitDelay": 0,  # 改为数字
                    "DevMajorEnabled": "false",
                    "DevMajorLimit": 80,  # 改为数字
                    "DevMajorText": "主要",
                    "DevMajorPriority": 1,  # 改为数字
                    "MajorInhibitor": "",
                    "DevMinorEnabled": "false",
                    "DevMinorLimit": 20,  # 改为数字
                    "DevMinorText": "次要",
                    "DevMinorPriority": 1,  # 改为数字
                    "MinorInhibitor": "",
                    "DevDeadband": 0,  # 改为数字
                    "DevTargetValue": 100,  # 改为数字
                    "DevDelay": 0,  # 改为数字
                    "RocEnabled": "false",
                    "RocPercent": 20,  # 改为数字
                    "RocTimeUnit": 0,  # 改为数字
                    "RocText": "变化率",
                    "RocDelay": 0,  # 改为数字
                    "RocPriority": 1,  # 改为数字
                    "RocInhibitor": "",
                    "StatusAlarmTableID": 0,  # 改为数字
                    "StatusAlarmEnabled": "false",
                    "StatusAlarmTableName": "",
                    "StatusInhibitor": "",
                    "AlarmGroup": "",
                    "ExtentField3": "",
                    "ExtentField4": "",
                    "ExtentField5": "",
                    "ExtentField6": "",
                    "ExtentField7": "",
                    "ExtentField8": "",
                    "StateEnumTable": "",
                    "IOConfigControl": "true",
                    "MaxRaw": 1000000000,  # 改为数字
                    "MinRaw": -1000000000,  # 改为数字
                    "IOEnable": "true",
                    "ForceRead": "false",
                    "ForceWrite": "false",
                    "DataConvertMode": 1,  # 改为数字
                    "NlnTableID": 0,  # 改为数字
                    "AddupMaxVal": 0,  # 改为数字
                    "AddupMinVal": 0  # 改为数字
                }
                
                # 数值型字段列表 - 用于判断应使用哪种样式
                number_fields = ["TagID", "ContainerType", "HisRecMode", "HisRecInterval", "AlarmType", 
                                "AlarmDelay", "DataConvertMode", "MaxValue", "MinValue", "InitialValue", 
                                "Sensitivity", "HisRecChangeDeadband", "HiHiPriority", "HiPriority", 
                                "LoPriority", "LoLoPriority", "LimitDeadband", "LimitDelay", "DevMajorLimit",
                                "DevMajorPriority", "DevMinorLimit", "DevMinorPriority", "DevDeadband", 
                                "DevTargetValue", "DevDelay", "RocPercent", "RocTimeUnit", "RocDelay", 
                                "RocPriority", "StatusAlarmTableID", "MaxRaw", "MinRaw", "NlnTableID", 
                                "AddupMaxVal", "AddupMinVal"]
                
                # 浮点数字段
                float_fields = ["HiHiLimit", "HiLimit", "LoLimit", "LoLoLimit"]
                
                # 设置从第二行开始填充数据（表头是第一行）
                disc_row_start = 1
                
                # 填充BOOL数据 - 从表头后的第二行开始添加
                for i, (_, row) in enumerate(bool_df.iterrows()):
                    # 获取变量信息
                    hmi_name = row.get("变量名称（HMI）", "")
                    description = row.get("变量描述", "")
                    station_name = row.get("场站名", "未知站点")
                    alarm_priority = row.get("报警等级", 1)  # 默认为1，确保为数字
                    
                    # 如果变量名为空，则自动补全
                    if pd.isna(hmi_name) or str(hmi_name).strip() == "":
                        channel_code = row.get("通道位号", "")
                        hmi_name = f"YLDW{channel_code}"
                        description = f"预留点位{channel_code}" if pd.isna(description) or str(description).strip() == "" else description
                    
                    # 如果报警等级为空，则设为默认值1
                    if pd.isna(alarm_priority) or str(alarm_priority).strip() == "":
                        alarm_priority = 1
                    else:
                        # 尝试将报警等级转换为数字
                        try:
                            alarm_priority = int(alarm_priority)
                        except (ValueError, TypeError):
                            alarm_priority = 1
                    
                    # 当前行索引
                    excel_row = disc_row_start + i
                    
                    # 填充必要的字段
                    if "TagID" in disc_column_indices:
                        disc_sheet.write(excel_row, disc_column_indices["TagID"], excel_row, number_style)
                    if "TagName" in disc_column_indices:
                        disc_sheet.write(excel_row, disc_column_indices["TagName"], hmi_name, standard_style)
                    if "Description" in disc_column_indices:
                        disc_sheet.write(excel_row, disc_column_indices["Description"], description, standard_style)
                    if "AlarmPriority" in disc_column_indices:
                        disc_sheet.write(excel_row, disc_column_indices["AlarmPriority"], alarm_priority, number_style)
                    if "AlarmGroup" in disc_column_indices:
                        disc_sheet.write(excel_row, disc_column_indices["AlarmGroup"], station_name, standard_style)
                    if "IOAccess" in disc_column_indices:
                        io_access = f"Sever1.{hmi_name}.Value"
                        disc_sheet.write(excel_row, disc_column_indices["IOAccess"], io_access, standard_style)
                    
                    # 填充固定值字段
                    for field, value in disc_fixed_values.items():
                        if field in disc_column_indices:
                            if field in number_fields:
                                disc_sheet.write(excel_row, disc_column_indices[field], value, number_style)
                            else:
                                disc_sheet.write(excel_row, disc_column_indices[field], value, standard_style)
                
                # 确定REAL数据的起始ID
                float_start_id = disc_row_start + len(bool_df)
                
                # 填充REAL数据
                for i, (_, row) in enumerate(real_df.iterrows()):
                    # 获取变量信息
                    hmi_name = row.get("变量名称（HMI）", "")
                    description = row.get("变量描述", "")
                    
                    # 如果变量名为空，则自动补全
                    if pd.isna(hmi_name) or str(hmi_name).strip() == "":
                        channel_code = row.get("通道位号", "")
                        hmi_name = f"YLDW{channel_code}"
                        description = f"预留点位{channel_code}" if pd.isna(description) or str(description).strip() == "" else description
                    
                    # 获取各类限值和报警信息
                    shh_value = row.get("SHH设定值", "")
                    sh_value = row.get("SH设定值", "")
                    sl_value = row.get("SL设定值", "")
                    sll_value = row.get("SLL设定值", "")
                    
                    # 判断是否启用各类报警
                    hihi_enabled = "true" if not pd.isna(shh_value) and shh_value and shh_value != "/" else "false"
                    hi_enabled = "true" if not pd.isna(sh_value) and sh_value and sh_value != "/" else "false"
                    lo_enabled = "true" if not pd.isna(sl_value) and sl_value and sl_value != "/" else "false"
                    lolo_enabled = "true" if not pd.isna(sll_value) and sll_value and sll_value != "/" else "false"
                    
                    # 当前行索引和ID
                    excel_row = disc_row_start + i
                    current_id = float_start_id + i
                    
                    # 填充必要的字段
                    if "TagID" in float_column_indices:
                        float_sheet.write(excel_row, float_column_indices["TagID"], current_id, number_style)
                    if "TagName" in float_column_indices:
                        float_sheet.write(excel_row, float_column_indices["TagName"], hmi_name, standard_style)
                    if "Description" in float_column_indices:
                        float_sheet.write(excel_row, float_column_indices["Description"], description, standard_style)
                    if "IOAccess" in float_column_indices:
                        io_access = f"Sever1.{hmi_name}.Value"
                        float_sheet.write(excel_row, float_column_indices["IOAccess"], io_access, standard_style)
                    
                    # 填充报警启用和限值信息
                    if "HiHiEnabled" in float_column_indices:
                        float_sheet.write(excel_row, float_column_indices["HiHiEnabled"], hihi_enabled, standard_style)
                    if "HiHiLimit" in float_column_indices and not pd.isna(shh_value) and shh_value and shh_value != "/":
                        # 尝试将限值转换为浮点数
                        try:
                            shh_float = float(shh_value)
                            float_sheet.write(excel_row, float_column_indices["HiHiLimit"], shh_float, float_style)
                        except (ValueError, TypeError):
                            float_sheet.write(excel_row, float_column_indices["HiHiLimit"], shh_value, standard_style)
                    
                    if "HiEnabled" in float_column_indices:
                        float_sheet.write(excel_row, float_column_indices["HiEnabled"], hi_enabled, standard_style)
                    if "HiLimit" in float_column_indices and not pd.isna(sh_value) and sh_value and sh_value != "/":
                        # 尝试将限值转换为浮点数
                        try:
                            sh_float = float(sh_value)
                            float_sheet.write(excel_row, float_column_indices["HiLimit"], sh_float, float_style)
                        except (ValueError, TypeError):
                            float_sheet.write(excel_row, float_column_indices["HiLimit"], sh_value, standard_style)
                    
                    if "LoEnabled" in float_column_indices:
                        float_sheet.write(excel_row, float_column_indices["LoEnabled"], lo_enabled, standard_style)
                    if "LoLimit" in float_column_indices and not pd.isna(sl_value) and sl_value and sl_value != "/":
                        # 尝试将限值转换为浮点数
                        try:
                            sl_float = float(sl_value)
                            float_sheet.write(excel_row, float_column_indices["LoLimit"], sl_float, float_style)
                        except (ValueError, TypeError):
                            float_sheet.write(excel_row, float_column_indices["LoLimit"], sl_value, standard_style)
                    
                    if "LoLoEnabled" in float_column_indices:
                        float_sheet.write(excel_row, float_column_indices["LoLoEnabled"], lolo_enabled, standard_style)
                    if "LoLoLimit" in float_column_indices and not pd.isna(sll_value) and sll_value and sll_value != "/":
                        # 尝试将限值转换为浮点数
                        try:
                            sll_float = float(sll_value)
                            float_sheet.write(excel_row, float_column_indices["LoLoLimit"], sll_float, float_style)
                        except (ValueError, TypeError):
                            float_sheet.write(excel_row, float_column_indices["LoLoLimit"], sll_value, standard_style)
                    
                    # 填充固定值字段
                    for field, value in float_fixed_values.items():
                        if field in float_column_indices:
                            if field in float_fields:
                                float_sheet.write(excel_row, float_column_indices[field], value, float_style)
                            elif field in number_fields:
                                float_sheet.write(excel_row, float_column_indices[field], value, number_style)
                            else:
                                float_sheet.write(excel_row, float_column_indices[field], value, standard_style)
                
                # 保存工作簿
                workbook.save(xls_output_path)
                
                # 检查文件是否生成成功
                if not (os.path.exists(xls_output_path) and os.path.getsize(xls_output_path) > 0):
                    raise ValueError(f"生成的文件不存在或为空: {xls_output_path}")
                
                # 关闭导出进度窗口
                if export_window:
                    export_window.destroy()
                
                return True
                
            except Exception as e:
                error_msg = f"生成数据词典点表失败: {str(e)}\n{traceback.format_exc()}"
                messagebox.showerror("错误", error_msg)
                
                if export_window:
                    export_window.destroy()
                return False
                
        except Exception as e:
            if export_window and export_window.winfo_exists():
                export_window.destroy()
            error_details = traceback.format_exc()
            messagebox.showerror("错误", f"生成数据词典点表时发生错误:\n{str(e)}\n\n详细错误信息:\n{error_details}")
            return False 