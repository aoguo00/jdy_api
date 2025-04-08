#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
深化设计数据查询工具 - PySide6用户界面模块

此模块负责实现用户界面，包括主窗口和各种交互功能。
支持查询项目、生成和上传IO点表、查看项目详情等操作。
"""

import os
import traceback
import sys
from PySide6.QtWidgets import (
    QMainWindow, QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
    QGroupBox, QLabel, QLineEdit, QPushButton, QTableWidget, 
    QTableWidgetItem, QHeaderView, QScrollArea, QMessageBox,
    QDialog, QFileDialog, QProgressDialog, QGridLayout, QFrame, QStyle
)
from PySide6.QtCore import Qt, QCoreApplication
from PySide6.QtGui import QFont, QCursor

# 导入业务逻辑控制器
from controller import ProjectController

# 导入字段定义
from io_generator import FormFields

class DetailWindow(QDialog):
    """
    深化清单详情窗口
    
    用于显示单个项目的深化清单详细信息
    提供表格化展示和基本场站信息显示
    """
    def __init__(self, parent, title, detail_data, station_info, field_mapping):
        """
        初始化详情窗口
        
        Args:
            parent: 父窗口
            title: 窗口标题
            detail_data: 详情数据
            station_info: 场站信息
            field_mapping: 字段映射
        """
        super().__init__(parent)
        self.setWindowTitle(f"深化清单详情 - {title}")
        self.resize(1000, 700)
        self.setMinimumSize(800, 600)
        self.field_mapping = field_mapping
        self.setup_ui(detail_data, station_info)
        
    def setup_ui(self, detail_data, station_info):
        """
        设置详情窗口UI
        
        创建并配置所有UI组件，包括场站信息面板和深化清单数据表格
        
        Args:
            detail_data: 详情数据
            station_info: 场站信息
        """
        # 主布局
        main_layout = QVBoxLayout(self)
        
        # 标题字体
        title_font = QFont()  # 使用系统默认字体
        title_font.setPointSize(16)
        title_font.setBold(True)
        header_font = QFont()  # 使用系统默认字体
        header_font.setPointSize(11)
        header_font.setBold(True)
        
        # 信息面板 - 显示场站基本信息
        info_group = QGroupBox("场站信息")
        info_layout = QVBoxLayout(info_group)
        
        # 场站基本信息网格
        info_grid_widget = QWidget()
        info_grid = QHBoxLayout(info_grid_widget)
        info_grid.setContentsMargins(5, 5, 5, 5)
        
        # 显示场站基本信息 (每行展示3个字段)
        info_rows = []
        current_row = QHBoxLayout()
        col_count = 0
        
        for field, value in station_info.items():
            if col_count >= 3:  # 每行最多3个字段
                info_rows.append(current_row)
                current_row = QHBoxLayout()
                col_count = 0
                
            field_layout = QHBoxLayout()
            
            label = QLabel(f"{field}:")
            label.setFont(header_font)
            field_layout.addWidget(label)
            
            value_label = QLabel(str(value))
            field_layout.addWidget(value_label)
            
            # 确保字段占据足够空间
            field_layout.setStretch(1, 1)
            current_row.addLayout(field_layout)
            
            col_count += 1
        
        # 添加最后一行
        if col_count > 0:
            info_rows.append(current_row)
        
        # 将所有行添加到布局
        info_grid_widget = QWidget()
        info_grid = QVBoxLayout(info_grid_widget)
        for row in info_rows:
            info_grid.addLayout(row)
        
        info_layout.addWidget(info_grid_widget)
        
        # 返回按钮
        button_layout = QHBoxLayout()
        button_layout.addStretch(1)
        back_button = QPushButton("返回上一页")
        back_button.clicked.connect(self.close)
        button_layout.addWidget(back_button)
        info_layout.addLayout(button_layout)
        
        main_layout.addWidget(info_group)
        
        # 深化清单数据表格
        detail_group = QGroupBox("深化清单详情")
        detail_layout = QVBoxLayout(detail_group)
        
        # 获取列
        columns = FormFields.ShenhuaForm.get_all_fields()
        
        # 创建表格
        self.detail_table = QTableWidget()
        self.detail_table.setColumnCount(len(columns))
        self.detail_table.setHorizontalHeaderLabels(columns)
        
        # 设置列宽
        for i, col in enumerate(columns):
            width = 100
            if col == FormFields.ShenhuaForm.EQUIPMENT_NAME:
                width = 150
            elif col == FormFields.ShenhuaForm.SPEC_MODEL:
                width = 150
            elif col == FormFields.ShenhuaForm.TECH_PARAMS:
                width = 200
            self.detail_table.setColumnWidth(i, width)
        
        # 水平表头伸展模式
        self.detail_table.horizontalHeader().setStretchLastSection(True)
        
        # 填充数据
        self._display_detail_data(detail_data)
        
        detail_layout.addWidget(self.detail_table)
        
        # 显示记录数
        status_label = QLabel(f"共 {len(detail_data)} 条记录")
        detail_layout.addWidget(status_label)
        
        main_layout.addWidget(detail_group)
        main_layout.setStretch(1, 3)  # 表格部分占据更多空间
        
    def _display_detail_data(self, detail_data):
        """在表格中显示深化清单数据"""
        if not detail_data:
            return
            
        # 获取所有字段
        fields = FormFields.ShenhuaForm.get_all_fields()
        
        # 设置行数
        self.detail_table.setRowCount(len(detail_data))
        
        # 填充表格数据
        for row_idx, item in enumerate(detail_data):
            for col_idx, field_name in enumerate(fields):
                field_id = self.field_mapping.get(field_name)
                if not field_id:
                    value = "-"
                else:
                    value = item.get(field_id, "")
                    if value is None or value == "":
                        value = "-"
                
                table_item = QTableWidgetItem(str(value))
                table_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)  # 只读
                self.detail_table.setItem(row_idx, col_idx, table_item)

class ProjectQueryApp(QMainWindow):
    """项目查询应用主界面类"""
    def __init__(self, api_client, field_mapping, shenhua_field_mapping):
        """
        初始化项目查询应用
        
        Args:
            api_client: 简道云API客户端实例
            field_mapping: 字段映射字典
            shenhua_field_mapping: 深化清单字段映射字典
        """
        super().__init__()
        self.api_client = api_client
        self.field_mapping = field_mapping
        self.shenhua_field_mapping = shenhua_field_mapping
        
        # 创建控制器
        self.controller = ProjectController(api_client, field_mapping, shenhua_field_mapping)
        
        # 设置UI
        self.setup_ui()
        
    def setup_ui(self):
        """设置应用界面"""
        # 设置窗口
        self.setWindowTitle("深化设计数据查询工具")
        self.resize(1000, 700)
        self.setMinimumSize(900, 600)
        
        # 设置标题字体
        title_font = QFont()  # 使用系统默认字体
        title_font.setPointSize(16)
        title_font.setBold(True)
        header_font = QFont()  # 使用系统默认字体
        header_font.setPointSize(11)
        header_font.setBold(True)
        
        # 创建中央部件和主布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        self.main_layout = QVBoxLayout(central_widget)
        self.main_layout.setContentsMargins(10, 10, 10, 10)
        
        # 创建查询框架
        query_group = QGroupBox("查询条件")
        query_layout = QVBoxLayout(query_group)
        
        # 创建查询输入区域
        input_layout = QHBoxLayout()
        
        # 项目编号标签和输入框
        proj_label = QLabel("项目编号:")
        proj_label.setFont(header_font)
        input_layout.addWidget(proj_label)
        
        self.project_input = QLineEdit()
        default_font = QFont()
        default_font.setPointSize(11)
        self.project_input.setFont(default_font)
        self.project_input.setText("OPP.25011100829")  # 默认值
        input_layout.addWidget(self.project_input)
        input_layout.setStretch(1, 1)  # 输入框占据更多空间
        
        query_layout.addLayout(input_layout)
        
        # 创建按钮布局
        button_layout = QHBoxLayout()
        
        # 查询按钮
        self.search_button = QPushButton("查询")
        self.search_button.clicked.connect(self.search_project)
        button_layout.addWidget(self.search_button)
        
        # 清空按钮
        self.clear_button = QPushButton("清空")
        self.clear_button.clicked.connect(self.clear_results)
        button_layout.addWidget(self.clear_button)
        
        # 生成点表按钮
        self.generate_button = QPushButton("生成点表")
        self.generate_button.clicked.connect(self.generate_io_table)
        self.generate_button.setEnabled(False)
        button_layout.addWidget(self.generate_button)
        
        # 上传点表按钮
        self.upload_button = QPushButton("上传点表")
        self.upload_button.clicked.connect(self.upload_io_table)
        self.upload_button.setEnabled(False)
        button_layout.addWidget(self.upload_button)
        
        # 上传FAT点表按钮
        self.fat_button = QPushButton("上传FAT点表")
        self.fat_button.clicked.connect(self.generate_fat_io_table)
        self.fat_button.setEnabled(False)
        button_layout.addWidget(self.fat_button)
        
        # 上传HMI点表按钮
        self.hmi_button = QPushButton("上传HMI点表")
        self.hmi_button.clicked.connect(self.generate_hmi_io_table)
        self.hmi_button.setEnabled(False)
        button_layout.addWidget(self.hmi_button)
        
        # 上传PLC点表按钮
        self.plc_button = QPushButton("上传PLC点表")
        self.plc_button.clicked.connect(self.generate_plc_io_table)
        self.plc_button.setEnabled(False)
        button_layout.addWidget(self.plc_button)
        
        query_layout.addLayout(button_layout)
        self.main_layout.addWidget(query_group)
        
        # 创建项目列表表格
        project_group = QGroupBox("项目列表")
        project_layout = QVBoxLayout(project_group)
        
        # 项目表格
        self.project_table = QTableWidget()
        self.project_table.setColumnCount(5)  # 5列: 项目名称, 场站, 项目编号, 深化设计编号, 客户名称
        self.project_table.setHorizontalHeaderLabels(["项目名称", "场站", "项目编号", "深化设计编号", "客户名称"])
        
        # 设置列宽
        self.project_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        
        # 添加选择项目的点击处理
        self.project_table.cellClicked.connect(self.on_project_selected)
        
        project_layout.addWidget(self.project_table)
        self.main_layout.addWidget(project_group)
        
        # 创建设备清单表格
        self.equipment_group = QGroupBox("设备清单")
        equipment_layout = QVBoxLayout(self.equipment_group)
        
        # 设备表格
        self.equipment_table = QTableWidget()
        self.equipment_table.setColumnCount(7)  # 7列: 设备名称, 品牌, 规格型号, 技术参数, 数量, 单位, 子系统
        self.equipment_table.setHorizontalHeaderLabels(["设备名称", "品牌", "规格型号", "技术参数", "数量", "单位", "子系统"])
        
        # 设置列宽
        self.equipment_table.setColumnWidth(0, 150)  # 设备名称
        self.equipment_table.setColumnWidth(2, 150)  # 规格型号
        self.equipment_table.setColumnWidth(3, 200)  # 技术参数
        self.equipment_table.horizontalHeader().setStretchLastSection(True)
        
        equipment_layout.addWidget(self.equipment_table)
        self.main_layout.addWidget(self.equipment_group)
        
        # 设置比例
        self.main_layout.setStretch(2, 1)  # 项目列表
        self.main_layout.setStretch(3, 2)  # 设备清单
    
    def on_project_selected(self, row, column):
        """
        处理项目选择事件
        
        Args:
            row: 选中的行
            column: 选中的列
        """
        try:
            # 获取项目数据
            project_data = self.controller.project_data
            if not project_data or row >= len(project_data):
                return
                
            # 获取选中的项目
            selected_project = project_data[row]
            
            # 加载设备清单数据
            equipment_data, station_name, project_number = self.controller.load_equipment_data(selected_project)
            
            # 显示设备清单数据
            if equipment_data:
                self.display_equipment_data(equipment_data)
                
                # 启用生成点表和上传点表按钮
                self.generate_button.setEnabled(True)
                self.upload_button.setEnabled(True)
                
                # 检查是否已上传IO数据，决定是否启用其他按钮
                self.update_button_states()
                
                QMessageBox.information(
                    self,
                    "数据加载成功",
                    f"已加载 {station_name} 的设备清单数据，共 {len(equipment_data)} 条记录"
                )
            else:
                QMessageBox.warning(
                    self,
                    "数据加载失败",
                    f"未找到 {station_name} 的设备清单数据"
                )
        except Exception as e:
            QMessageBox.critical(
                self,
                "错误",
                f"加载设备清单时发生错误: {str(e)}"
            )
            traceback.print_exc()
            
    def display_equipment_data(self, equipment_data):
        """在表格中显示设备清单数据"""
        if not equipment_data:
            self.equipment_table.setRowCount(0)
            return
            
        # 设置行数
        self.equipment_table.setRowCount(len(equipment_data))
        
        # 填充表格数据
        for row_idx, item in enumerate(equipment_data):
            # 设备名称
            name_value = item.get("设备名称", "")
            if name_value is None or name_value == "":
                name_value = ""
            name_item = QTableWidgetItem(str(name_value))
            name_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)  # 只读
            self.equipment_table.setItem(row_idx, 0, name_item)
            
            # 品牌
            brand_value = item.get("品牌", "")
            if brand_value is None or brand_value == "":
                brand_value = ""
            brand_item = QTableWidgetItem(str(brand_value))
            brand_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.equipment_table.setItem(row_idx, 1, brand_item)
            
            # 规格型号
            spec_value = item.get("规格型号", "")
            if spec_value is None or spec_value == "":
                spec_value = ""
            spec_item = QTableWidgetItem(str(spec_value))
            spec_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.equipment_table.setItem(row_idx, 2, spec_item)
            
            # 技术参数
            params_value = item.get("技术参数", "")
            if params_value is None or params_value == "":
                params_value = ""
            params_item = QTableWidgetItem(str(params_value))
            params_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.equipment_table.setItem(row_idx, 3, params_item)
            
            # 数量
            quantity_value = item.get("数量", "")
            if quantity_value is None or quantity_value == "":
                quantity_value = ""
            quantity_item = QTableWidgetItem(str(quantity_value))
            quantity_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.equipment_table.setItem(row_idx, 4, quantity_item)
            
            # 单位
            unit_value = item.get("单位", "")
            if unit_value is None or unit_value == "":
                unit_value = ""
            unit_item = QTableWidgetItem(str(unit_value))
            unit_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.equipment_table.setItem(row_idx, 5, unit_item)
            
            # 子系统
            system_value = item.get("子系统", "")
            if system_value is None or system_value == "":
                system_value = ""
            system_item = QTableWidgetItem(str(system_value))
            system_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.equipment_table.setItem(row_idx, 6, system_item)
    
    def search_project(self):
        """执行项目查询"""
        project_number = self.project_input.text().strip()
        
        if not project_number:
            QMessageBox.warning(self, "输入错误", "请输入项目编号")
            return
            
        try:
            # 显示加载中消息
            QApplication.setOverrideCursor(Qt.WaitCursor)
            
            # 调用控制器执行查询
            project_data = self.controller.search_project_data(project_number)
            
            # 恢复光标
            QApplication.restoreOverrideCursor()
            
            # 显示结果
            if project_data:
                self._display_results(project_data)
                QMessageBox.information(
                    self, 
                    "查询成功", 
                    f"找到 {len(project_data)} 个相关项目"
                )
            else:
                QMessageBox.information(
                    self, 
                    "查询结果", 
                    "未找到相关项目"
                )
        except Exception as e:
            # 恢复光标
            QApplication.restoreOverrideCursor()
            
            QMessageBox.critical(
                self, 
                "查询失败", 
                f"查询过程中发生错误: {str(e)}"
            )
            traceback.print_exc()
            
    def _display_results(self, project_data):
        """在表格中显示项目数据"""
        if not project_data:
            self.project_table.setRowCount(0)
            return
            
        # 设置行数
        self.project_table.setRowCount(len(project_data))
        
        # 填充表格数据
        for row_idx, item in enumerate(project_data):
            # 项目名称
            field_id = self.field_mapping.get(FormFields.MainForm.PROJECT_NAME, "")
            name_value = item.get(field_id, "")
            # 不使用 or 操作符，而是明确处理None和空字符串
            if name_value is None or name_value == "":
                name_value = ""
            name_item = QTableWidgetItem(str(name_value))
            name_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)  # 只读
            self.project_table.setItem(row_idx, 0, name_item)
            
            # 场站
            field_id = self.field_mapping.get(FormFields.MainForm.STATION, "")
            station_value = item.get(field_id, "")
            if station_value is None or station_value == "":
                station_value = ""
            station_item = QTableWidgetItem(str(station_value))
            station_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.project_table.setItem(row_idx, 1, station_item)
            
            # 项目编号
            field_id = self.field_mapping.get(FormFields.MainForm.PROJECT_NUMBER, "")
            number_value = item.get(field_id, "")
            if number_value is None or number_value == "":
                number_value = ""
            number_item = QTableWidgetItem(str(number_value))
            number_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.project_table.setItem(row_idx, 2, number_item)
            
            # 深化设计编号
            field_id = self.field_mapping.get(FormFields.MainForm.SHENHUA_NUMBER, "")
            design_value = item.get(field_id, "")
            if design_value is None or design_value == "":
                design_value = ""
            design_item = QTableWidgetItem(str(design_value))
            design_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.project_table.setItem(row_idx, 3, design_item)
            
            # 客户名称
            field_id = self.field_mapping.get(FormFields.MainForm.CLIENT_NAME, "")
            client_value = item.get(field_id, "")
            if client_value is None or client_value == "":
                client_value = ""
            client_item = QTableWidgetItem(str(client_value))
            client_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.project_table.setItem(row_idx, 4, client_item)
    
    def clear_results(self):
        """清空结果和表格"""
        # 清空控制器中的数据
        self.controller.clear_data()
        
        # 清空表格
        self.project_table.setRowCount(0)
        self.equipment_table.setRowCount(0)
        
        # 禁用按钮
        self.update_button_states()
        
        # 显示消息
        QMessageBox.information(self, "清空成功", "已清空所有数据")
        
    def generate_io_table(self):
        """生成IO点表并导出到Excel"""
        if not self.controller.current_equipment_data:
            QMessageBox.warning(self, "错误", "没有可用的设备数据")
            return
            
        try:
            # 选择保存文件的路径
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "保存IO点表",
                "",
                "Excel文件 (*.xlsx *.xls)"
            )
            
            if not file_path:
                return  # 用户取消了操作
                
            # 显示加载中消息
            QApplication.setOverrideCursor(Qt.WaitCursor)
            
            # 调用控制器生成IO点表
            result = self.controller.generate_io_table(file_path)
            
            # 恢复光标
            QApplication.restoreOverrideCursor()
            
            if result:
                QMessageBox.information(
                    self,
                    "导出成功",
                    f"IO点表已成功导出到:\n{file_path}"
                )
            else:
                QMessageBox.warning(
                    self,
                    "导出失败",
                    "IO点表导出失败"
                )
        except Exception as e:
            # 恢复光标
            QApplication.restoreOverrideCursor()
            
            QMessageBox.critical(
                self,
                "错误",
                f"生成IO点表时发生错误: {str(e)}"
            )
            traceback.print_exc()
            
    def generate_fat_io_table(self):
        """生成FAT点表并上传到简道云"""
        try:
            # 显示加载中光标
            QApplication.setOverrideCursor(Qt.WaitCursor)
            
            # 调用控制器生成FAT点表
            self.controller.generate_fat_io_table(self)
            
            # 恢复光标
            QApplication.restoreOverrideCursor()
            
        except ValueError as e:
            # 恢复光标
            QApplication.restoreOverrideCursor()
            QMessageBox.warning(
                self,
                "输入错误",
                str(e)
            )
        except Exception as e:
            # 恢复光标
            QApplication.restoreOverrideCursor()
            QMessageBox.critical(
                self,
                "错误",
                f"生成FAT点表时发生错误: {str(e)}"
            )
            traceback.print_exc()
            
    def generate_hmi_io_table(self):
        """生成HMI上位点表并上传到简道云"""
        try:
            # 显示加载中光标
            QApplication.setOverrideCursor(Qt.WaitCursor)
            
            # 调用控制器生成HMI点表
            self.controller.generate_hmi_io_table(self)
            
            # 恢复光标
            QApplication.restoreOverrideCursor()
            
        except ValueError as e:
            # 恢复光标
            QApplication.restoreOverrideCursor()
            QMessageBox.warning(
                self,
                "输入错误",
                str(e)
            )
        except Exception as e:
            # 恢复光标
            QApplication.restoreOverrideCursor()
            QMessageBox.critical(
                self,
                "错误",
                f"生成HMI点表时发生错误: {str(e)}"
            )
            traceback.print_exc()
            
    def generate_plc_io_table(self):
        """生成PLC下位点表并上传到简道云"""
        try:
            # 显示加载中光标
            QApplication.setOverrideCursor(Qt.WaitCursor)
            
            # 调用控制器生成PLC点表
            self.controller.generate_plc_io_table(self)
            
            # 恢复光标
            QApplication.restoreOverrideCursor()
            
        except ValueError as e:
            # 恢复光标
            QApplication.restoreOverrideCursor()
            QMessageBox.warning(
                self,
                "输入错误",
                str(e)
            )
        except Exception as e:
            # 恢复光标
            QApplication.restoreOverrideCursor()
            QMessageBox.critical(
                self,
                "错误",
                f"生成PLC点表时发生错误: {str(e)}"
            )
            traceback.print_exc()
            
    def upload_io_table(self):
        """上传已补全信息的IO点表"""
        if not self.controller.current_equipment_data:
            QMessageBox.warning(self, "错误", "没有可用的设备数据")
            return
            
        try:
            # 选择要上传的文件
            file_path, _ = QFileDialog.getOpenFileName(
                self,
                "选择IO点表",
                "",
                "Excel文件 (*.xlsx *.xls)"
            )
            
            if not file_path:
                return  # 用户取消了操作
                
            # 创建进度对话框
            progress = QProgressDialog("正在上传IO点表...", "取消", 0, 100, self)
            progress.setWindowTitle("上传进度")
            progress.setWindowModality(Qt.WindowModal)
            progress.setMinimumDuration(0)
            progress.setValue(0)
            progress.show()
            
            # 定义进度更新回调函数
            def update_progress(text):
                # 防止在关闭窗口后继续更新进度
                if progress.wasCanceled():
                    return
                    
                # 更新进度文本
                progress.setLabelText(text)
                
                # 更新进度条值
                QApplication.processEvents()  # 确保UI更新
                
            # 调用控制器上传IO点表
            success, result = self.controller.upload_io_table(file_path, update_progress)
            
            # 关闭进度对话框
            progress.close()
            
            if success:
                # 创建上传信息显示框
                file_name = os.path.basename(file_path)
                
                # 检查是否已存在上传信息框
                if hasattr(self, "upload_info_frame"):
                    self.upload_info_frame.deleteLater()
                
                # 创建信息框
                self.upload_info_frame = QFrame(self.centralWidget())
                self.upload_info_frame.setFrameShape(QFrame.StyledPanel)
                self.upload_info_frame.setFrameShadow(QFrame.Raised)
                self.upload_info_frame.setStyleSheet("background-color: #e6f7ff; border: 1px solid #91d5ff; border-radius: 4px; padding: 10px;")
                
                # 创建布局
                info_layout = QHBoxLayout(self.upload_info_frame)
                info_layout.setContentsMargins(10, 5, 10, 5)
                
                # 添加文件图标
                file_icon_label = QLabel()
                file_icon_label.setPixmap(self.style().standardIcon(QStyle.SP_FileIcon).pixmap(24, 24))
                info_layout.addWidget(file_icon_label)
                
                # 添加文件信息
                file_info_label = QLabel(f"已上传文件: <b>{file_name}</b> (共 {result['rows']} 条记录)")
                info_layout.addWidget(file_info_label, 1)
                
                # 添加到主布局的合适位置（在设备组下方）
                self.main_layout.insertWidget(self.main_layout.indexOf(self.equipment_group) + 1, self.upload_info_frame)
                
                QMessageBox.information(
                    self,
                    "上传成功",
                    f"IO点表已成功上传，共 {result['rows']} 条记录"
                )
                
                # 启用其他上传点表按钮
                self.hmi_button.setEnabled(True)
                self.plc_button.setEnabled(True)
                self.fat_button.setEnabled(True)
            else:
                error_msg = "\n".join(result)
                QMessageBox.warning(
                    self,
                    "上传失败",
                    f"IO点表上传失败:\n{error_msg}"
                )
        except Exception as e:
            QMessageBox.critical(
                self,
                "错误",
                f"上传IO点表时发生错误: {str(e)}"
            )
            traceback.print_exc()
            
    def update_button_states(self):
        """更新按钮状态"""
        has_equipment_data = bool(self.controller.current_equipment_data)
        has_io_data = hasattr(self.controller, 'uploaded_io_data') and self.controller.uploaded_io_data is not None
        
        self.generate_button.setEnabled(has_equipment_data)
        self.upload_button.setEnabled(has_equipment_data)
        self.hmi_button.setEnabled(has_equipment_data and has_io_data)
        self.plc_button.setEnabled(has_equipment_data and has_io_data)
        self.fat_button.setEnabled(has_equipment_data and has_io_data)

    def view_detail(self):
        """查看深化清单详情"""
        selected_row = self.project_table.currentRow()
        
        if selected_row < 0:
            QMessageBox.warning(self, "错误", "请先选择一个项目")
            return
            
        try:
            # 显示加载中消息
            QApplication.setOverrideCursor(Qt.WaitCursor)
            
            # 获取选中项目数据
            selected_project = self.controller.project_data[selected_row]
            
            # 加载设备清单数据
            equipment_data, station_info, project_number = self.controller.load_equipment_data(selected_project)
            
            # 恢复光标
            QApplication.restoreOverrideCursor()
            
            # 显示设备清单数据
            if equipment_data:
                self.display_equipment_data(equipment_data)
                
                # 启用生成点表和上传点表按钮
                self.generate_button.setEnabled(True)
                self.upload_button.setEnabled(True)
                
                QMessageBox.information(
                    self, 
                    "加载成功", 
                    f"已加载 {station_info.get('场站', '')} 的设备清单，共 {len(equipment_data)} 条记录"
                )
            else:
                QMessageBox.information(
                    self, 
                    "查询结果", 
                    "未找到相关设备清单数据"
                )
        except Exception as e:
            # 恢复光标
            QApplication.restoreOverrideCursor()
            
            QMessageBox.critical(
                self, 
                "查询失败", 
                f"获取设备清单时发生错误: {str(e)}"
            )
            traceback.print_exc() 