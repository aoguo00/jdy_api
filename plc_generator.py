#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
PLC点表生成模块 - 用于生成PLC点表

此模块负责将IO数据转换为PLC点表格式
实现了PLC点表生成的核心逻辑，包括数据处理和Excel生成
支持进度显示和错误处理
"""

import os
import pandas as pd
import traceback
import xlrd
import xlwt
from datetime import datetime
from pathlib import Path
from io_generator import IOChannelModels
from PySide6.QtWidgets import QMessageBox, QProgressDialog
from PySide6.QtCore import Qt

# 添加导入配置
from config.settings import TEMPLATE_DIR

class PLCGenerator:
    """
    PLC点表生成器类
    负责将IO数据转换为标准PLC点表格式
    提供实例方法用于生成PLC点表Excel文件
    支持UI进度显示和异常处理
    """
    
    # 定义需要处理的扩展点位列表
    EXTENDED_POINTS = [
        {"name": "SLL设定点位", "plc_addr": "SLL设定点位_PLC地址", "suffix": "_LoLoLimit"},
        {"name": "SL设定点位", "plc_addr": "SL设定点位_PLC地址", "suffix": "_LoLimit"},
        {"name": "SH设定点位", "plc_addr": "SH设定点位_PLC地址", "suffix": "_HiLimit"},
        {"name": "SHH设定点位", "plc_addr": "SHH设定点位_PLC地址", "suffix": "_HiHiLimit"},
        {"name": "LL报警", "plc_addr": "LL报警_PLC地址", "suffix": "_LL"},
        {"name": "L报警", "plc_addr": "L报警_PLC地址", "suffix": "_L"},
        {"name": "H报警", "plc_addr": "H报警_PLC地址", "suffix": "_H"},
        {"name": "HH报警", "plc_addr": "HH报警_PLC地址", "suffix": "_HH"},
        {"name": "维护使能开关点位", "plc_addr": "维护使能开关点位_PLC地址", "suffix": "_MaintenanceEnable"}
    ]
    
    def __init__(self, jdy_api_client=None):
        """
        初始化PLC点表生成器
        
        Args:
            jdy_api_client: 简道云API客户端
        """
        self.api_client = jdy_api_client
        self.module_dir = os.path.dirname(os.path.abspath(__file__))
        self.template_file = os.path.join(self.module_dir, "templates", "PLC点表.xls")
        self.uploaded_io_data = None
        
    def set_uploaded_io_data(self, io_data):
        """
        设置已上传的IO数据
        
        Args:
            io_data: 已上传的IO数据
        """
        self.uploaded_io_data = io_data
    
    @staticmethod
    def get_default_field_values(data_type):
        """
        根据数据类型获取默认字段值
        
        Args:
            data_type: 数据类型（如"REAL", "BOOL"等）
            
        Returns:
            dict: 默认字段值字典
        """
        if data_type == "BOOL":
            return {
                "初始值": "FALSE",
                "掉电保护": "FALSE",
                "可强制": "TRUE",
                "SOE使能": "TRUE"
            }
        else:  # REAL或其他类型
            return {
                "初始值": "0",
                "掉电保护": "TRUE",  # REAL类型数据的掉电保护设置为TRUE
                "可强制": "TRUE",
                "SOE使能": "FALSE"
            }
    
    def generate_plc_table(self, equipment_data, station_name, project_number, parent_window=None):
        """
        生成PLC点表
        
        Args:
            equipment_data: 设备数据列表
            station_name: 场站名称
            project_number: 项目编号
            parent_window: 父窗口对象，用于显示进度
            
        Returns:
            tuple: (success, message, file_path)
        """
        # 检查上传的IO数据是否存在
        if self.uploaded_io_data is None or self.uploaded_io_data.empty:
            QMessageBox.warning(parent_window, "警告", "请先上传IO点表数据")
            return False, "请先上传IO点表数据", ""
            
        # 检查设备数据是否存在
        if not equipment_data or len(equipment_data) == 0:
            QMessageBox.warning(parent_window, "警告", "没有可用的设备数据")
            return False, "没有可用的设备数据", ""
            
        # 检查模板文件是否存在
        if not os.path.exists(self.template_file):
            QMessageBox.warning(parent_window, "警告", f"模板文件不存在: {self.template_file}")
            return False, f"模板文件不存在: {self.template_file}", ""
        
        # 创建进度窗口
        progress = QProgressDialog("生成PLC点表中...", "取消", 0, 100, parent_window)
        progress.setWindowTitle("生成PLC点表")
        progress.setWindowModality(Qt.WindowModal)
        progress.setMinimumDuration(0)
        progress.setValue(0)
        progress.show()
        
        try:
            # 创建输出目录 - 使用临时目录
            import tempfile
            # 使用临时目录替代本地output目录
            temp_dir = tempfile.gettempdir()
                
            # 生成输出文件名 - 使用临时文件
            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
            output_file = os.path.join(temp_dir, f"PLC点表_{station_name}_{timestamp}.xls")
            
            # 更新进度
            progress.setValue(10)
            
            # 读取Excel模板
            progress.setLabelText("正在读取模板文件...")
            template_wb = xlrd.open_workbook(self.template_file)
            template_sheet = template_wb.sheet_by_index(0)
            
            # 确定列映射关系
            # 注：使用模板头行确定列索引
            column_mapping = {
                "变量名": 0,         # 第1列: 变量名
                "变量地址": 1,       # 第2列: 变量地址
                "注释": 2,           # 第3列: 注释
                "变量类型": 3,       # 第4列: 变量类型
                "初始值": 4,         # 第5列: 初始值
                "掉电保护": 5,       # 第6列: 掉电保护
                "可强制": 6,         # 第7列: 可强制
                "SOE使能": 7         # 第8列: SOE使能
            }
            
            # 获取实际的列头
            column_headers = {}
            for col in range(template_sheet.ncols):
                value = template_sheet.cell_value(1, col)  # 第2行(索引1)是表头
                for field, idx in column_mapping.items():
                    if field in value:
                        column_headers[field] = col
                        break
            
            # 设置列索引
            var_name_col = column_headers.get("变量名", column_mapping["变量名"])
            address_col = column_headers.get("变量地址", column_mapping["变量地址"])
            comment_col = column_headers.get("注释", column_mapping["注释"])
            var_type_col = column_headers.get("变量类型", column_mapping["变量类型"])
            # 其他默认字段
            init_value_col = column_headers.get("初始值", column_mapping["初始值"])
            power_protect_col = column_headers.get("掉电保护", column_mapping["掉电保护"])
            forcible_col = column_headers.get("可强制", column_mapping["可强制"])
            soe_enable_col = column_headers.get("SOE使能", column_mapping["SOE使能"])
            
            # 创建新的Excel工作簿和工作表
            workbook = xlwt.Workbook(encoding='utf-8')
            worksheet = workbook.add_sheet('PLC点表')
            
            # 创建宋体字体样式，大小为11
            font = xlwt.Font()
            font.name = '宋体'
            font.height = 220  # 字体大小 11 (220 = 11 * 20)
            
            # 设置单元格通用样式
            common_style = xlwt.XFStyle()
            common_style.font = font
            
            # 设置单元格格式：文本格式用于地址列，同时使用宋体
            text_style = xlwt.XFStyle()
            text_style.num_format_str = '@'
            text_style.font = font
            
            # 从模板复制表头和格式（前两行），并应用宋体格式
            for row in range(2):
                for col in range(template_sheet.ncols):
                    value = template_sheet.cell_value(row, col)
                    worksheet.write(row, col, value, common_style)
            
            # 行计数器，用于记录当前Excel中的行数
            excel_row_counter = 2  # 从第3行开始，前2行是标题
            
            # 处理所有基础变量
            progress.setLabelText("正在生成PLC表数据...")
            for i, (_, row) in enumerate(self.uploaded_io_data.iterrows(), 1):
                # 获取基础变量信息
                var_name = row.get("变量名称（HMI）", "")
                plc_address = row.get("PLC绝对地址", "")
                description = row.get("变量描述", "")
                data_type = row.get("数据类型", "")
                channel_code = row.get("通道位号", "")
                
                # 如果变量名为空，则自动补全
                if pd.isna(var_name) or str(var_name).strip() == "":
                    var_name = f"YLDW{channel_code}"
                    description = f"预留点位{channel_code}" if pd.isna(description) or str(description).strip() == "" else description
                
                # 填充基础变量数据
                worksheet.write(excel_row_counter, var_name_col, var_name, common_style)
                worksheet.write(excel_row_counter, address_col, str(plc_address), text_style) 
                worksheet.write(excel_row_counter, comment_col, description, common_style)
                worksheet.write(excel_row_counter, var_type_col, data_type, common_style)
                
                # 填充默认字段
                default_values = PLCGenerator.get_default_field_values(data_type)
                if init_value_col is not None:
                    worksheet.write(excel_row_counter, init_value_col, default_values["初始值"], common_style)
                if power_protect_col is not None:
                    worksheet.write(excel_row_counter, power_protect_col, default_values["掉电保护"], common_style)
                if forcible_col is not None:
                    worksheet.write(excel_row_counter, forcible_col, default_values["可强制"], common_style)
                if soe_enable_col is not None:
                    worksheet.write(excel_row_counter, soe_enable_col, default_values["SOE使能"], common_style)
                
                # 增加行计数器
                excel_row_counter += 1
                
                # 处理该行的所有扩展点位
                for ext_point in PLCGenerator.EXTENDED_POINTS:
                    point_name = ext_point["name"]
                    plc_addr_field = ext_point["plc_addr"]
                    point_suffix = ext_point["suffix"]
                    
                    # 获取扩展点位的值和PLC地址
                    point_value = row.get(point_name, "")
                    point_plc_addr = row.get(plc_addr_field, "")
                    
                    # 如果扩展点位值为空或"/"或None，则跳过
                    if pd.isna(point_value) or not point_value or point_value == "/":
                        continue
                    
                    # 如果PLC地址为空或"/"或None，则跳过
                    if pd.isna(point_plc_addr) or not point_plc_addr or point_plc_addr == "/":
                        continue
                    
                    # 为扩展点位创建变量名和描述
                    ext_var_name = str(var_name) + point_suffix
                    ext_description = str(description) + " " + point_name.replace("_PLC地址", "")
                    
                    # 确定数据类型(根据地址判断)
                    ext_data_type = "REAL" if point_plc_addr.startswith("%MD") else "BOOL"
                    
                    # 填充扩展点位数据
                    worksheet.write(excel_row_counter, var_name_col, ext_var_name, common_style)
                    worksheet.write(excel_row_counter, address_col, str(point_plc_addr), text_style)
                    worksheet.write(excel_row_counter, comment_col, ext_description, common_style)
                    worksheet.write(excel_row_counter, var_type_col, ext_data_type, common_style)
                    
                    # 填充默认字段
                    ext_default_values = PLCGenerator.get_default_field_values(ext_data_type)
                    if init_value_col is not None:
                        worksheet.write(excel_row_counter, init_value_col, ext_default_values["初始值"], common_style)
                    if power_protect_col is not None:
                        worksheet.write(excel_row_counter, power_protect_col, ext_default_values["掉电保护"], common_style)
                    if forcible_col is not None:
                        worksheet.write(excel_row_counter, forcible_col, ext_default_values["可强制"], common_style)
                    if soe_enable_col is not None:
                        worksheet.write(excel_row_counter, soe_enable_col, ext_default_values["SOE使能"], common_style)
                    
                    # 增加行计数器
                    excel_row_counter += 1
                
                # 更新进度
                progress.setValue(10 + int(80 * (i / len(self.uploaded_io_data))))
                if progress.wasCanceled():
                    break
            
            # 保存工作簿
            progress.setLabelText("正在保存文件...")
            workbook.save(output_file)
            progress.setValue(95)
            
            # 检查文件是否生成成功
            if os.path.exists(output_file) and os.path.getsize(output_file) > 0:
                # 不显示本地文件生成成功的弹窗，只返回结果
                progress.setValue(100)
                return True, "PLC点表生成成功", output_file
            else:
                raise ValueError(f"生成的文件不存在或为空: {output_file}")
                
        except Exception as e:
            traceback.print_exc()
            QMessageBox.critical(parent_window, "错误", f"生成PLC点表时出错: {str(e)}")
            return False, f"生成PLC点表时出错: {str(e)}", ""
            
        finally:
            progress.close()
            
    def _generate_plc_data(self, equipment_data):
        """
        生成PLC数据列表
        
        Args:
            equipment_data: 设备数据列表
            
        Returns:
            list: PLC数据列表
        """
        plc_data = []
        
        for item in equipment_data:
            # 从设备数据中提取设备名称和子系统
            device_name = item.get("设备名称", "")  
            subsystem = item.get("子系统", "")
            
            if not device_name:
                continue
                
            # 为每个设备创建一个PLC表条目
            plc_entry = {
                "device_name": device_name,
                "subsystem": subsystem,
                "control_name": "",          # 控制量默认为空
                "signal_type": "AI",         # 信号类型默认为AI
                "function": "",              # 功能默认为空
                "default_value": "0",        # 默认值默认为0
                "data_address": "",          # 数据地址默认为空
                "hmi_address": "",           # HMI点地址默认为空
                "hmi_note": "",              # HMI点注释默认为空
                "program_point": "",         # 程序点名默认为空
                "program_note": "",          # 程序点注释默认为空
                "io_position": "",           # I/O位置默认为空
                "remark": ""                 # 备注默认为空
            }
            
            plc_data.append(plc_entry)
            
        return plc_data