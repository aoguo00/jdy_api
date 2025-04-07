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
from tkinter import messagebox, Toplevel, ttk

from pathlib import Path
from io_generator import IOChannelModels

# 添加导入配置
from config.settings import TEMPLATE_DIR, PLC_TEMPLATE

class PLCGenerator:
    """
    PLC点表生成器类
    
    负责将IO数据转换为标准PLC点表格式
    提供静态方法用于生成PLC点表Excel文件
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
        {"name": "维护值设定点位", "plc_addr": "维护值设定点位_PLC地址", "suffix": "_whz"},
        {"name": "维护使能开关点位", "plc_addr": "维护使能开关点位_PLC地址", "suffix": "_MAIN_EN"}
    ]
    
    @classmethod
    def get_plc_column_mapping(cls):
        """
        获取PLC点表的列映射配置
        
        Returns:
            dict: 列名到默认列索引的映射
        """
        return {
            "变量名": 2,        # 第2列：变量名
            "直接地址": 3,      # 第3列：直接地址
            "变量说明": 4,      # 第4列：变量说明
            "变量类型": 5,      # 第5列：变量类型
            "初始值": 6,        # 第6列：初始值
            "掉电保护": 7,      # 第7列：掉电保护
            "可强制": 8,        # 第8列：可强制
            "SOE使能": 9        # 第9列：SOE使能
        }
    
    @classmethod
    def get_default_field_values(cls, data_type):
        """
        获取PLC点表默认字段值
        
        Args:
            data_type: 数据类型 ("BOOL" 或 "REAL")
            
        Returns:
            dict: 默认字段值
        """
        # 通用默认值
        defaults = {
            "掉电保护": "FALSE",
            "可强制": "TRUE",
            "SOE使能": "FALSE"
        }
        
        # 根据数据类型设置初始值
        if data_type == "BOOL":
            defaults["初始值"] = "FALSE"
        else:  # REAL或其他类型
            defaults["初始值"] = "0"
            defaults["掉电保护"] = "TRUE"  # REAL类型数据的掉电保护设置为TRUE
            
        return defaults
    
    @staticmethod
    def generate_plc_table(io_data, output_path, root_window=None):
        """
        生成PLC点表
        
        处理IO数据并生成符合PLC要求的点表Excel文件
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
                export_window.title("正在导出PLC点表")
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
                
                export_label = ttk.Label(export_window, text="正在生成PLC点表，请稍候...", font=("Microsoft YaHei", 10))
                export_label.pack(pady=20)
                export_window.update()
            
            # 检查是否有上传数据
            if len(io_data) == 0:
                if export_window:
                    export_window.destroy()
                messagebox.showwarning("警告", "上传的IO点表中没有数据！")
                return False
            
            # 检查本地模板文件是否存在
            template_file = os.path.join(TEMPLATE_DIR, PLC_TEMPLATE)
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
                
                # 读取模板文件结构
                template_workbook = xlrd.open_workbook(template_file)
                template_sheet = template_workbook.sheet_by_index(0)  # 获取第一个工作表
                
                # 获取表头信息(第二行)
                column_headers = {}
                for col in range(template_sheet.ncols):
                    header = template_sheet.cell_value(1, col)  # 使用第二行作为标题行
                    if header:
                        column_headers[header] = col
                
                # 根据实际模板设置列映射
                column_mapping = PLCGenerator.get_plc_column_mapping()
                
                # 获取每个列的索引，如果列名不存在，则使用默认值
                var_name_col = column_headers.get("变量名", column_mapping["变量名"])
                address_col = column_headers.get("直接地址", column_mapping["直接地址"])
                comment_col = column_headers.get("变量说明", column_mapping["变量说明"])
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
                for i, (_, row) in enumerate(io_data.iterrows(), 1):
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
                
                # 保存工作簿
                workbook.save(xls_output_path)
                
                # 检查文件是否生成成功
                if os.path.exists(xls_output_path) and os.path.getsize(xls_output_path) > 0:
                    # 不再在这里显示弹窗和打开文件，改为由调用者处理
                    # 直接返回成功
                    result = True
                else:
                    raise ValueError(f"生成的文件不存在或为空: {xls_output_path}")
                
                # 关闭导出进度窗口
                if export_window:
                    export_window.destroy()
                
                return result
                
            except Exception as e:
                error_msg = f"生成PLC点表文件失败: {str(e)}\n{traceback.format_exc()}"
                messagebox.showerror("错误", error_msg)
                
                if export_window:
                    export_window.destroy()
                return False
                
        except Exception as e:
            if export_window and export_window.winfo_exists():
                export_window.destroy()
            error_details = traceback.format_exc()
            messagebox.showerror("错误", f"生成PLC点表时发生错误:\n{str(e)}\n\n详细错误信息:\n{error_details}")
            return False 