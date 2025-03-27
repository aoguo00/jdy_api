#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
深化设计数据查询工具 - 用户界面模块

此模块实现了应用程序的用户界面，包含以下主要组件：
1. ProjectQueryApp类 - 主应用界面，包含查询和结果展示功能
2. DetailWindow类 - 详情窗口，用于显示深化清单详细信息

该模块负责用户交互和数据展示，采用Tkinter实现GUI界面
提供友好的用户体验和直观的数据展示
"""

import tkinter as tk
from tkinter import ttk, messagebox, font, filedialog
import traceback
import os

# 导入数据访问层和控制器
from io_generator import FormFields
from controller import ProjectController

class DetailWindow:
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
        self.detail_window = tk.Toplevel(parent)
        self.detail_window.title(f"深化清单详情 - {title}")
        self.detail_window.geometry("1000x700")
        self.detail_window.minsize(800, 600)
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
        # 主框架
        main_frame = ttk.Frame(self.detail_window, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题字体
        title_font = font.Font(family="Microsoft YaHei", size=16, weight="bold")
        header_font = font.Font(family="Microsoft YaHei", size=11, weight="bold")
        
        # 信息面板 - 显示场站基本信息
        info_frame = ttk.LabelFrame(main_frame, text="场站信息", padding=10)
        info_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 场站基本信息
        info_grid = ttk.Frame(info_frame)
        info_grid.pack(fill=tk.X, pady=5)
        
        # 显示场站基本信息
        row = 0
        col = 0
        for field, value in station_info.items():
            if col > 2:  # 每行最多3个字段
                row += 1
                col = 0
                
            label = ttk.Label(info_grid, text=f"{field}:", font=header_font)
            label.grid(row=row, column=col*2, sticky=tk.W, padx=5, pady=2)
            
            value_label = ttk.Label(info_grid, text=value)
            value_label.grid(row=row, column=col*2+1, sticky=tk.W, padx=5, pady=2)
            
            col += 1
        
        # 返回按钮
        back_button = ttk.Button(info_frame, text="返回上一页", command=self.detail_window.destroy)
        back_button.pack(side=tk.RIGHT, padx=5, pady=5)
        
        # 深化清单数据表格
        detail_frame = ttk.LabelFrame(main_frame, text="深化清单详情", padding=10)
        detail_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 创建带滚动条的表格框架
        table_frame = ttk.Frame(detail_frame)
        table_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建滚动条
        scrollbar_y = ttk.Scrollbar(table_frame, orient=tk.VERTICAL)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        
        scrollbar_x = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 设置列 - 使用FormFields中的字段
        columns = FormFields.ShenhuaForm.get_all_fields()
        
        # 创建表格
        self.detail_tree = ttk.Treeview(
            table_frame, 
            columns=columns,
            show="headings",
            yscrollcommand=scrollbar_y.set,
            xscrollcommand=scrollbar_x.set
        )
        
        # 配置滚动条
        scrollbar_y.config(command=self.detail_tree.yview)
        scrollbar_x.config(command=self.detail_tree.xview)
        
        # 设置列标题和宽度
        for col in columns:
            self.detail_tree.heading(col, text=col)
            self.detail_tree.column(col, width=100, minwidth=80)
        
        # 设置某些列的特定宽度
        if FormFields.ShenhuaForm.EQUIPMENT_NAME in columns:
            self.detail_tree.column(FormFields.ShenhuaForm.EQUIPMENT_NAME, width=150)
        if FormFields.ShenhuaForm.SPEC_MODEL in columns:
            self.detail_tree.column(FormFields.ShenhuaForm.SPEC_MODEL, width=150)
        if FormFields.ShenhuaForm.TECH_PARAMS in columns:
            self.detail_tree.column(FormFields.ShenhuaForm.TECH_PARAMS, width=200)
        
        self.detail_tree.pack(fill=tk.BOTH, expand=True)
        
        # 填充数据
        self._display_detail_data(detail_data)
        
        # 显示记录数
        status_label = ttk.Label(detail_frame, text=f"共 {len(detail_data)} 条记录", anchor=tk.W)
        status_label.pack(side=tk.LEFT, padx=5, pady=(5, 0))
        
    def _display_detail_data(self, detail_data):
        """在表格中显示深化清单数据"""
        if not detail_data:
            return
            
        # 获取所有字段
        fields = FormFields.ShenhuaForm.get_all_fields()
        
        # 填充表格数据
        for item in detail_data:
            row_data = []
            for field_name in fields:
                field_id = self.field_mapping.get(field_name)
                if not field_id:
                    row_data.append("-")
                    continue
                    
                value = item.get(field_id, "")
                if value is None or value == "":
                    value = "-"
                row_data.append(value)
                
            self.detail_tree.insert("", tk.END, values=row_data)

class ProjectQueryApp:
    """项目查询应用主界面类"""
    def __init__(self, root, api_client, field_mapping, shenhua_field_mapping):
        """
        初始化项目查询应用
        
        Args:
            root: Tkinter根窗口
            api_client: 简道云API客户端实例
            field_mapping: 字段映射字典
            shenhua_field_mapping: 深化清单字段映射字典
        """
        self.root = root
        self.api_client = api_client
        self.field_mapping = field_mapping
        self.shenhua_field_mapping = shenhua_field_mapping
        
        # 创建控制器
        self.controller = ProjectController(api_client, field_mapping, shenhua_field_mapping)
        
        # 设置UI
        self.setup_ui()
        
        # 注册关闭窗口的处理函数
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
    def on_closing(self):
        """处理窗口关闭事件"""
        self.root.destroy()
        
    def setup_ui(self):
        """设置应用界面"""
        # 设置窗口
        self.root.title("深化设计数据查询工具")
        self.root.geometry("1000x700")  # 增加窗口大小
        self.root.minsize(900, 600)
        
        # 设置样式
        self.style = ttk.Style()
        self.style.theme_use('clam')  # 使用clam主题，看起来更现代
        
        # 配置标题字体 - 使用更适合中文显示的字体
        title_font = font.Font(family="Microsoft YaHei", size=16, weight="bold")
        header_font = font.Font(family="Microsoft YaHei", size=11, weight="bold")
        
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建顶部标题
        title_label = ttk.Label(main_frame, text="深化设计数据查询", font=title_font)
        title_label.pack(pady=(0, 10))
        
        # 创建查询框架
        query_frame = ttk.LabelFrame(main_frame, text="查询条件", padding=10)
        query_frame.pack(fill=tk.X, pady=5)
        
        # 创建查询输入区域
        input_frame = ttk.Frame(query_frame)
        input_frame.pack(fill=tk.X, pady=(5, 0))
        
        # 项目编号标签和输入框
        proj_label = ttk.Label(input_frame, text="项目编号:", font=header_font)
        proj_label.pack(side=tk.LEFT, padx=5)
        
        self.project_input = ttk.Entry(input_frame, width=30, font=("Microsoft YaHei", 11))
        self.project_input.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        self.project_input.insert(0, "OPP.25011100829")  # 默认值
        
        # 创建按钮框架 - 作为大容器
        button_frame = ttk.Frame(query_frame)
        button_frame.pack(fill=tk.X, pady=5)
        
        # 配置按钮布局为网格，使按钮能够均匀分布
        button_frame.columnconfigure(0, weight=1)  # 查询按钮列
        button_frame.columnconfigure(1, weight=1)  # 清空按钮列
        button_frame.columnconfigure(2, weight=1)  # 生成点表按钮列
        button_frame.columnconfigure(3, weight=1)  # 上传点表按钮列
        button_frame.columnconfigure(4, weight=1)  # 生成HMI点表按钮列
        button_frame.columnconfigure(5, weight=1)  # 生成PLC点表按钮列
        
        # 查询按钮
        self.search_button = ttk.Button(button_frame, text="查询", command=self.search_project, width=8)
        self.search_button.grid(row=0, column=0, padx=3, pady=5, sticky='ew')
        
        # 清空按钮
        self.clear_button = ttk.Button(button_frame, text="清空", command=self.clear_results, width=8)
        self.clear_button.grid(row=0, column=1, padx=3, pady=5, sticky='ew')
        
        # 生成点表按钮
        self.generate_io_button = ttk.Button(button_frame, text="生成点表", command=self.generate_io_table, width=8)
        self.generate_io_button.grid(row=0, column=2, padx=3, pady=5, sticky='ew')
        
        # 上传点表按钮
        self.upload_io_button = ttk.Button(button_frame, text="上传点表", command=self.upload_io_table, width=8)
        self.upload_io_button.grid(row=0, column=3, padx=3, pady=5, sticky='ew')
        
        # 生成HMI点表按钮
        self.generate_hmi_io_button = ttk.Button(button_frame, text="上传HMI点表", command=self.generate_hmi_io_table, width=12)
        self.generate_hmi_io_button.grid(row=0, column=4, padx=3, pady=5, sticky='ew')
        
        # 生成PLC点表按钮
        self.generate_plc_io_button = ttk.Button(button_frame, text="上传PLC点表", command=self.generate_plc_io_table, width=12)
        self.generate_plc_io_button.grid(row=0, column=5, padx=3, pady=5, sticky='ew')
        
        # 文件信息Frame - 创建但初始不显示
        self.file_info_frame = ttk.Frame(button_frame)
        
        # 文件名标签（初始为空）
        self.file_name_label = ttk.Label(self.file_info_frame, text="")
        self.file_name_label.pack(side=tk.RIGHT)
        
        # 文件图标标签（初始为空）
        self.file_icon_label = ttk.Label(self.file_info_frame, text="", font=("Microsoft YaHei", 12))
        self.file_icon_label.pack(side=tk.RIGHT, padx=(0, 5))
        
        # 创建垂直分割的窗格 - 上面显示项目，下面显示设备清单
        self.paned_window = ttk.PanedWindow(main_frame, orient=tk.VERTICAL)
        self.paned_window.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 创建项目结果区域
        projects_frame = ttk.LabelFrame(self.paned_window, text="项目列表", padding=10)
        self.paned_window.add(projects_frame, weight=1)
        
        # 创建设备清单区域
        equipment_frame = ttk.LabelFrame(self.paned_window, text="设备清单", padding=10)
        self.paned_window.add(equipment_frame, weight=2)
        
        # === 项目列表表格 ===
        # 创建表格
        columns = FormFields.MainForm.get_all_fields()
        
        # 创建带滚动条的表格框架
        projects_table_frame = ttk.Frame(projects_frame)
        projects_table_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建滚动条
        scrollbar_y = ttk.Scrollbar(projects_table_frame, orient=tk.VERTICAL)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        
        scrollbar_x = ttk.Scrollbar(projects_table_frame, orient=tk.HORIZONTAL)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 创建表格
        self.results_tree = ttk.Treeview(
            projects_table_frame, 
            columns=columns,
            show="headings",
            yscrollcommand=scrollbar_y.set,
            xscrollcommand=scrollbar_x.set
        )
        
        # 配置滚动条
        scrollbar_y.config(command=self.results_tree.yview)
        scrollbar_x.config(command=self.results_tree.xview)
        
        # 设置列标题和宽度
        for col in columns:
            self.results_tree.heading(col, text=col)
            self.results_tree.column(col, width=150, minwidth=100)
        
        self.results_tree.pack(fill=tk.BOTH, expand=True)
        
        # === 设备清单表格 ===
        # 创建表格框架
        equipment_table_frame = ttk.Frame(equipment_frame)
        equipment_table_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建滚动条
        equipment_scrollbar_y = ttk.Scrollbar(equipment_table_frame, orient=tk.VERTICAL)
        equipment_scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        
        equipment_scrollbar_x = ttk.Scrollbar(equipment_table_frame, orient=tk.HORIZONTAL)
        equipment_scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 设置列 - 使用FormFields中的字段
        equipment_columns = FormFields.ShenhuaForm.get_all_fields()
        
        # 创建表格
        self.equipment_tree = ttk.Treeview(
            equipment_table_frame, 
            columns=equipment_columns,
            show="headings",
            yscrollcommand=equipment_scrollbar_y.set,
            xscrollcommand=equipment_scrollbar_x.set
        )
        
        # 配置滚动条
        equipment_scrollbar_y.config(command=self.equipment_tree.yview)
        equipment_scrollbar_x.config(command=self.equipment_tree.xview)
        
        # 设置列标题和宽度
        for col in equipment_columns:
            self.equipment_tree.heading(col, text=col)
            self.equipment_tree.column(col, width=100, minwidth=80)
        
        # 设置某些列的特定宽度
        if FormFields.ShenhuaForm.EQUIPMENT_NAME in equipment_columns:
            self.equipment_tree.column(FormFields.ShenhuaForm.EQUIPMENT_NAME, width=150)
        if FormFields.ShenhuaForm.SPEC_MODEL in equipment_columns:
            self.equipment_tree.column(FormFields.ShenhuaForm.SPEC_MODEL, width=150)
        if FormFields.ShenhuaForm.TECH_PARAMS in equipment_columns:
            self.equipment_tree.column(FormFields.ShenhuaForm.TECH_PARAMS, width=200)
        
        self.equipment_tree.pack(fill=tk.BOTH, expand=True)
        
        # 设置状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("准备就绪")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 绑定回车键到查询
        self.project_input.bind("<Return>", lambda event: self.search_project())
        
        # 绑定项目表格选中事件
        self.results_tree.bind("<<TreeviewSelect>>", self.on_project_selected)
        
        # 初始化按钮状态
        self.update_button_states()
        
    def on_project_selected(self, event):
        """处理项目表格选中事件"""
        # 获取选中的项目
        selection = self.results_tree.selection()
        if not selection:
            return
            
        # 获取选中项目的索引
        item_id = selection[0]
        item_idx = self.results_tree.index(item_id)
        
        if item_idx < 0 or item_idx >= len(self.controller.project_data):
            messagebox.showwarning("错误", "无法获取所选项目的详细信息")
            return
        
        # 获取选中的项目数据
        project = self.controller.project_data[item_idx]
        
        # 更新状态
        self.status_var.set("正在获取设备清单数据...")
        self.root.update()
        
        # 清空设备清单表格
        for item in self.equipment_tree.get_children():
            self.equipment_tree.delete(item)
        
        try:
            # 显示加载进度窗口
            progress_window = tk.Toplevel(self.root)
            progress_window.title("数据加载中")
            progress_window.geometry("300x100")
            progress_window.transient(self.root)
            progress_window.grab_set()
            
            # 设置窗口在主窗口中央显示
            progress_window.withdraw()  # 先隐藏窗口
            progress_window.update()    # 更新窗口信息
            
            # 计算窗口位置
            x = self.root.winfo_x() + (self.root.winfo_width() - 300) // 2
            y = self.root.winfo_y() + (self.root.winfo_height() - 100) // 2
            progress_window.geometry(f"300x100+{x}+{y}")
            
            progress_window.deiconify()  # 显示窗口
            
            # 获取场站名称
            field_id = self.field_mapping.get(FormFields.MainForm.STATION)
            station_name = project.get(field_id, "[未知场站]")
            
            progress_label = ttk.Label(progress_window, text=f"正在加载 {station_name} 的设备清单...", font=("Microsoft YaHei", 10))
            progress_label.pack(pady=20)
            progress_window.update()
            
            # 调用控制器加载设备数据
            equipment_list, station_name, project_number = self.controller.load_equipment_data(project)
            
            # 关闭进度窗口
            progress_window.destroy()
            
            if not equipment_list:
                messagebox.showinfo("提示", f"未找到 {station_name} 的设备清单数据")
                self.status_var.set("未找到设备清单数据")
                return
            
            # 显示设备清单数据
            self._display_equipment_data(equipment_list)
            
            # 更新状态
            self.status_var.set(f"已加载 {station_name} 的设备清单数据 ({len(equipment_list)} 条记录)")
            
        except Exception as e:
            # 清理可能存在的进度窗口
            if 'progress_window' in locals() and progress_window.winfo_exists():
                progress_window.destroy()
                
            # 详细记录异常信息
            error_details = traceback.format_exc()
            
            messagebox.showerror("错误", f"获取设备清单数据时发生错误: {str(e)}\n\n详细错误信息:\n{error_details}")
            self.status_var.set("获取设备清单数据失败")
    
    def _display_equipment_data(self, equipment_data):
        """在设备清单表格中显示设备数据"""
        if not equipment_data:
            return
            
        # 获取所有字段
        fields = FormFields.ShenhuaForm.get_all_fields()
        
        # 显示场站信息
        if equipment_data and "_station" in equipment_data[0]:
            station_name = equipment_data[0]["_station"]
            # 更新设备清单框架的标题
            for frame in self.paned_window.panes():
                if hasattr(frame, 'cget') and frame.cget('text') == "设备清单":
                    frame.configure(text=f"设备清单 - {station_name}")
                    break
        
        # 填充表格数据
        for item in equipment_data:
            row_data = []
            for field_name in fields:
                value = item.get(field_name, "")
                if value is None or value == "":
                    value = "-"
                row_data.append(value)
                
            self.equipment_tree.insert("", tk.END, values=row_data)
    
    def search_project(self):
        """执行项目查询"""
        # 获取项目编号
        project_number = self.project_input.get().strip()
        
        # 验证输入
        if not project_number:
            messagebox.showwarning("输入错误", "请输入项目编号")
            return
            
        # 更新状态
        self.status_var.set(f"正在查询项目: {project_number}...")
        self.root.update()
        
        # 执行查询
        try:
            # 清空现有结果
            for item in self.results_tree.get_children():
                self.results_tree.delete(item)
            
            # 使用控制器执行查询
            project_data = self.controller.search_project_data(project_number)
            
            if not project_data:
                messagebox.showinfo("查询结果", f"未找到与项目编号 '{project_number}' 匹配的记录")
                self.status_var.set("查询完成，未找到匹配记录")
                return
                
            # 显示结果
            self._display_results(project_data)
            self.status_var.set(f"查询完成，找到 {len(project_data)} 条记录")
            
        except Exception as e:
            messagebox.showerror("查询错误", f"查询时发生错误: {str(e)}")
            self.status_var.set("查询失败")
    
    def _display_results(self, project_data):
        """在表格中显示查询结果"""
        # 获取所有字段
        fields = FormFields.MainForm.get_all_fields()
        
        # 填充表格数据
        for project in project_data:
            # 创建行数据列表
            row_data = []
            
            # 按列顺序整理数据
            for field_name in fields:
                field_id = self.field_mapping.get(field_name)
                if not field_id:
                    row_data.append("[无数据]")
                    continue
                    
                field_value = project.get(field_id, "[无数据]")
                if field_value is None or field_value == "":
                    field_value = "[无数据]"
                row_data.append(field_value)
            
            # 将行数据添加到表格
            self.results_tree.insert("", tk.END, values=row_data)
    
    def clear_results(self):
        """清空查询结果和输入"""
        self.project_input.delete(0, tk.END)
        
        # 清空项目表格
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
            
        # 清空设备清单表格
        for item in self.equipment_tree.get_children():
            self.equipment_tree.delete(item)
            
        # 使用控制器清空数据
        self.controller.clear_data()
        
        # 更新按钮状态
        self.update_button_states()
        
        self.status_var.set("已清空查询结果")

    def generate_io_table(self):
        """
        生成IO点表并导出到Excel
        """
        try:
            # 使用控制器验证数据是否可用
            if not self.controller.current_equipment_data:
                messagebox.showwarning("警告", "请先查询项目并选择一个项目加载设备数据！")
                return
                
            # 获取保存路径
            project_number = self.project_input.get().strip()
            default_filename = f"{project_number}_IO点表.xlsx"
            output_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel文件", "*.xlsx")],
                initialfile=default_filename
            )
            
            if not output_path:
                return  # 用户取消
                
            # 显示导出进度窗口
            export_window = tk.Toplevel(self.root)
            export_window.title("正在导出Excel")
            export_window.geometry("300x100")
            export_window.transient(self.root)
            export_window.grab_set()
            
            # 设置窗口在主窗口中央显示
            export_window.withdraw()  # 先隐藏窗口
            export_window.update()    # 更新窗口信息
            
            # 计算窗口位置
            x = self.root.winfo_x() + (self.root.winfo_width() - 300) // 2
            y = self.root.winfo_y() + (self.root.winfo_height() - 100) // 2
            export_window.geometry(f"300x100+{x}+{y}")
            
            export_window.deiconify()  # 显示窗口
            
            export_label = ttk.Label(export_window, text="正在生成IO点表，请稍候...", font=("Microsoft YaHei", 10))
            export_label.pack(pady=20)
            export_window.update()
            
            try:
                # 使用控制器生成IO点表
                export_success = self.controller.generate_io_table(output_path)
                
                # 关闭导出进度窗口
                export_window.destroy()
                
                if export_success:
                    # 显示成功消息
                    messagebox.showinfo("成功", "导出成功")
                    
                    # 尝试打开文件
                    try:
                        os.startfile(output_path)  # 在Windows上打开文件
                    except Exception:
                        # 如果打开失败，不进行任何处理
                        pass
                else:
                    messagebox.showerror("错误", "导出IO点表失败！")
                    
            except PermissionError as e:
                export_window.destroy()
                messagebox.showerror(
                    "权限错误", 
                    f"无法写入文件: {output_path}\n\n可能的原因:\n- 文件已被其他程序打开\n- 没有写入权限\n\n请关闭已打开的Excel文件，或选择其他位置保存。"
                )
            except Exception as e:
                export_window.destroy()
                messagebox.showerror("文件错误", f"文件访问错误: {str(e)}")
                
        except Exception as e:
            # 清理可能存在的导出窗口
            if 'export_window' in locals() and export_window.winfo_exists():
                export_window.destroy()
                
            error_details = traceback.format_exc()
            messagebox.showerror("错误", f"生成IO点表时发生错误:\n{str(e)}\n\n详细错误信息:\n{error_details}") 
            
    def generate_hmi_io_table(self):
        """
        生成HMI上位点表并上传到简道云
        """
        try:
            # 使用控制器调用逻辑方法
            self.controller.generate_hmi_io_table(self.root)
        except ValueError as e:
            messagebox.showwarning("警告", str(e))
        except Exception as e:
            error_details = traceback.format_exc()
            messagebox.showerror("错误", f"生成HMI点表时发生错误:\n{str(e)}\n\n详细错误信息:\n{error_details}")
        
    def generate_plc_io_table(self):
        """
        生成PLC下位点表并上传到简道云
        """
        try:
            # 使用控制器调用逻辑方法
            self.controller.generate_plc_io_table(self.root)
        except ValueError as e:
            messagebox.showwarning("警告", str(e))
        except Exception as e:
            error_details = traceback.format_exc()
            messagebox.showerror("错误", f"生成PLC点表时发生错误:\n{str(e)}\n\n详细错误信息:\n{error_details}")
        
    def upload_io_table(self):
        """
        上传已补全信息的IO点表Excel文件
        """
        try:
            # 检查是否已加载设备数据
            if not self.controller.current_equipment_data:
                messagebox.showwarning("警告", "请先查询项目并选择一个项目加载设备数据！")
                return
                
            # 获取项目编号，用于文件名参考
            project_number = self.project_input.get().strip()
            if not project_number:
                project_number = "未知项目"
                
            # 打开文件选择对话框
            input_path = filedialog.askopenfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel文件", "*.xlsx")],
                title="选择已补全的IO点表Excel文件"
            )
            
            if not input_path:
                return  # 用户取消
                
            # 显示加载进度窗口
            upload_window = tk.Toplevel(self.root)
            upload_window.title("正在上传文件")
            upload_window.geometry("300x100")
            upload_window.transient(self.root)
            upload_window.grab_set()
            
            # 设置窗口在主窗口中央显示
            upload_window.withdraw()  # 先隐藏窗口
            upload_window.update()    # 更新窗口信息
            
            # 计算窗口位置
            x = self.root.winfo_x() + (self.root.winfo_width() - 300) // 2
            y = self.root.winfo_y() + (self.root.winfo_height() - 100) // 2
            upload_window.geometry(f"300x100+{x}+{y}")
            
            upload_window.deiconify()  # 显示窗口
            
            upload_label = ttk.Label(upload_window, text="正在读取Excel文件，请稍候...", font=("Microsoft YaHei", 10))
            upload_label.pack(pady=20)
            
            # 创建进度更新回调
            def update_progress(text):
                upload_label.config(text=text)
                upload_window.update()
            
            try:
                # 调用控制器上传IO表
                success, result = self.controller.upload_io_table(input_path, update_progress)
                
                # 关闭上传进度窗口
                upload_window.destroy()
                
                if not success:
                    # 创建错误消息
                    error_message = "上传文件验证失败:\n\n" + "\n\n".join(result)
                    
                    # 创建详细错误窗口
                    error_window = tk.Toplevel(self.root)
                    error_window.title("文件验证错误")
                    error_window.geometry("600x400")
                    error_window.transient(self.root)
                    error_window.grab_set()
                    
                    # 设置窗口在主窗口中央显示
                    x = self.root.winfo_x() + (self.root.winfo_width() - 600) // 2
                    y = self.root.winfo_y() + (self.root.winfo_height() - 400) // 2
                    error_window.geometry(f"600x400+{x}+{y}")
                    
                    # 创建带滚动条的文本框
                    text_frame = ttk.Frame(error_window)
                    text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
                    
                    scrollbar = ttk.Scrollbar(text_frame)
                    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
                    
                    error_text = tk.Text(text_frame, wrap=tk.WORD, yscrollcommand=scrollbar.set)
                    error_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
                    scrollbar.config(command=error_text.yview)
                    
                    # 插入错误信息
                    error_text.insert(tk.END, error_message)
                    error_text.config(state=tk.DISABLED)  # 设置为只读
                    
                    # 关闭按钮
                    close_button = ttk.Button(error_window, text="关闭", command=error_window.destroy)
                    close_button.pack(pady=10)
                    
                    return
                
                # 验证通过，更新按钮状态
                self.update_button_states()
                
                # 显示成功消息
                messagebox.showinfo("成功", f"文件验证并上传成功！\n共读取了 {result['rows']} 行数据。")
                
                # 更新状态栏
                self.status_var.set(f"已上传点表文件: {result['filename']}")
                
            except Exception as e:
                if upload_window.winfo_exists():
                    upload_window.destroy()
                messagebox.showerror("文件错误", f"无法读取或验证文件: {str(e)}")
                
        except Exception as e:
            # 清理可能存在的上传窗口
            if 'upload_window' in locals() and upload_window.winfo_exists():
                upload_window.destroy()
                
            error_details = traceback.format_exc()
            messagebox.showerror("错误", f"上传点表时发生错误:\n{str(e)}\n\n详细错误信息:\n{error_details}")

    def update_button_states(self):
        """根据数据状态更新按钮的启用/禁用状态"""
        # 安全检查：确保属性存在
        if not hasattr(self, 'generate_hmi_io_button') or not hasattr(self, 'generate_plc_io_button'):
            return
            
        # 如果已上传点表数据，启用HMI和PLC点表生成按钮，否则禁用
        if self.controller.uploaded_io_data is not None:
            self.generate_hmi_io_button.config(state="normal")
            self.generate_plc_io_button.config(state="normal")
            
            # 显示文件名信息
            if hasattr(self, 'file_name_label') and hasattr(self.controller, 'uploaded_io_file_path'):
                # 设置文件信息区域的列权重
                self.root.update_idletasks()  # 确保UI已更新
                button_frame = self.file_info_frame.master
                button_frame.columnconfigure(6, weight=1)
                
                # 显示文件信息框架
                file_name = os.path.basename(self.controller.uploaded_io_file_path)
                self.file_name_label.config(text=f"已上传: {file_name}")
                self.file_icon_label.config(text="📊")
                self.file_info_frame.grid(row=0, column=6, padx=3, pady=5, sticky='e')
        else:
            self.generate_hmi_io_button.config(state="disabled")
            self.generate_plc_io_button.config(state="disabled")
            
            # 移除文件信息框架
            if hasattr(self, 'file_info_frame'):
                self.file_info_frame.grid_forget()
                
                # 清空文件信息显示
                if hasattr(self, 'file_name_label'):
                    self.file_name_label.config(text="")
                    self.file_icon_label.config(text="") 