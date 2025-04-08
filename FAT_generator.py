#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
FAT点表生成模块 - 用于生成FAT点表

此模块负责将IO数据转换为FAT点表格式
实现了FAT点表生成的核心逻辑，包括数据处理和Excel生成
支持进度显示和错误处理
"""

import os
import pandas as pd
import traceback
# 将tkinter导入替换为PySide6导入
from PySide6.QtWidgets import QProgressDialog, QMessageBox
from PySide6.QtCore import Qt
from pathlib import Path
import tempfile
from io_generator import IOChannelModels

class FATGenerator:
    """
    FAT点表生成器类
    
    负责将IO数据转换为标准FAT点表格式
    提供静态方法用于生成FAT点表Excel文件
    支持UI进度显示和异常处理
    """
    
    @staticmethod
    def generate_fat_table(io_data, output_path, root_window=None):
        """
        生成FAT点表
        
        处理IO数据并生成FAT点表Excel文件
        支持进度显示窗口，提供用户友好的导出体验
        
        Args:
            io_data: 上传的IO点表数据(DataFrame)
            output_path: 输出文件路径
            root_window: 父窗口，用于显示进度窗口
            
        Returns:
            bool: 操作是否成功
        """
        # 导入xlwt库，确保它在这个方法中可用
        import xlwt
        
        try:
            # 显示导出进度窗口
            export_window = None
            if root_window:
                # 使用PySide6的QProgressDialog替代Toplevel
                export_window = QProgressDialog("正在生成FAT点表，请稍候...", "取消", 0, 100, root_window)
                export_window.setWindowTitle("正在导出FAT点表")
                export_window.setWindowModality(Qt.WindowModal)
                export_window.setMinimumDuration(0)
                export_window.setValue(0)
                export_window.show()
            
            # 检查是否有上传数据
            if len(io_data) == 0:
                if export_window:
                    export_window.close()
                QMessageBox.warning(root_window, "警告", "上传的IO点表中没有数据！")
                return False
            
            try:
                # 更新进度
                if export_window:
                    export_window.setLabelText("正在处理数据...")
                    export_window.setValue(10)
                
                # 确保输出路径是xls格式
                xls_output_path = str(Path(output_path).with_suffix('.xls'))
                
                # 确保目标文件不存在
                if os.path.exists(xls_output_path):
                    os.remove(xls_output_path)
                
                # 创建FAT点表数据的副本，以便进行处理
                fat_data = io_data.copy()
                
                # 获取需要高亮的字段列表（需要填写"/"的字段）
                highlight_fields = IOChannelModels.get_highlight_fields()
                
                # 先将所有将要填入字符串值的列转换为object/string类型
                if "变量名称（HMI）" in fat_data.columns:
                    fat_data["变量名称（HMI）"] = fat_data["变量名称（HMI）"].astype(object)
                
                if "变量描述" in fat_data.columns:
                    fat_data["变量描述"] = fat_data["变量描述"].astype(object)
                
                # 将所有需要填"/"的字段也转换为字符串类型
                for field in highlight_fields:
                    if field in fat_data.columns:
                        fat_data[field] = fat_data[field].astype(object)
                
                # 更新进度
                if export_window:
                    export_window.setLabelText("正在补全数据...")
                    export_window.setValue(30)
                
                # 第一步：处理变量名称为空的情况，补全变量名称和变量描述
                for idx, row in fat_data.iterrows():
                    # 获取变量名称和描述
                    hmi_name = row.get("变量名称（HMI）", "")
                    channel_code = row.get("通道位号", "")
                    
                    # 检查变量名称是否为空
                    if pd.isna(hmi_name) or str(hmi_name).strip() == "":
                        # 自动补全变量名称
                        fat_data.at[idx, "变量名称（HMI）"] = f"YLDW{channel_code}"
                        
                        # 自动补全变量描述（无论原来是否为空）
                        fat_data.at[idx, "变量描述"] = f"预留点位{channel_code}"
                
                # 更新进度
                if export_window:
                    export_window.setLabelText("正在填充默认值...")
                    export_window.setValue(50)
                
                # 第二步：处理所有标黄的空单元格，填写为"/"
                for idx, row in fat_data.iterrows():
                    for field in highlight_fields:
                        if field in fat_data.columns:
                            current_value = row.get(field, "")
                            if pd.isna(current_value) or str(current_value).strip() == "":
                                # 直接赋值，前面已经将列转换为对象类型
                                fat_data.at[idx, field] = "/"
                
                # 更新进度
                if export_window:
                    export_window.setLabelText("正在创建Excel文件...")
                    export_window.setValue(70)
                
                # 创建新的Excel工作簿
                workbook = xlwt.Workbook(encoding='utf-8')
                worksheet = workbook.add_sheet('FAT点表')
                
                # 创建宋体字体样式
                font = xlwt.Font()
                font.name = '宋体'
                font.height = 220  # 字体大小 11 (220 = 11 * 20)
                
                # 设置单元格样式
                common_style = xlwt.XFStyle()
                common_style.font = font
                
                # 写入表头
                for col_idx, col_name in enumerate(fat_data.columns):
                    worksheet.write(0, col_idx, col_name, common_style)
                
                # 更新进度
                if export_window:
                    export_window.setLabelText("正在写入数据...")
                    export_window.setValue(80)
                
                # 写入数据
                total_rows = len(fat_data)
                for row_idx, (_, row) in enumerate(fat_data.iterrows(), 1):
                    # 周期性更新进度
                    if export_window and row_idx % 10 == 0:  # 每10行更新一次进度
                        progress = 80 + int((row_idx / total_rows) * 15)  # 从80%到95%的进度
                        export_window.setValue(progress)
                        export_window.setLabelText(f"正在写入数据... ({row_idx}/{total_rows})")
                    
                    for col_idx, col_name in enumerate(fat_data.columns):
                        value = row.get(col_name, "")
                        # 处理空值
                        if pd.isna(value):
                            value = ""
                        worksheet.write(row_idx, col_idx, value, common_style)
                
                # 保存工作簿
                workbook.save(xls_output_path)
                
                # 更新进度
                if export_window:
                    export_window.setLabelText("正在完成操作...")
                    export_window.setValue(95)
                
                # 检查文件是否生成成功
                if os.path.exists(xls_output_path) and os.path.getsize(xls_output_path) > 0:
                    result = True
                else:
                    raise ValueError(f"生成的文件不存在或为空: {xls_output_path}")
                
                # 关闭导出进度窗口
                if export_window:
                    export_window.close()
                
                return result
                
            except Exception as e:
                error_msg = f"生成FAT点表文件失败: {str(e)}\n{traceback.format_exc()}"
                QMessageBox.critical(root_window, "错误", error_msg)
                
                if export_window:
                    export_window.close()
                return False
                
        except Exception as e:
            if export_window and export_window.isVisible():
                export_window.close()
            error_details = traceback.format_exc()
            QMessageBox.critical(root_window, "错误", f"生成FAT点表时发生错误:\n{str(e)}\n\n详细错误信息:\n{error_details}")
            return False 