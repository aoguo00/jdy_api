#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
深化设计数据查询工具 - 业务逻辑控制器模块

此模块实现了应用程序的业务逻辑，负责数据处理、计算和验证，
将UI层与数据访问层分离，提高代码的可维护性和可测试性。
"""

import os

# 导入数据访问层和数据服务
from data_service import  DataServiceFactory

class ProjectController:
    """项目控制器类，处理业务逻辑"""
    
    def __init__(self, api_client, field_mapping, shenhua_field_mapping):
        """
        初始化项目控制器
        
        Args:
            api_client: 简道云API客户端实例
            field_mapping: 字段映射字典
            shenhua_field_mapping: 深化清单字段映射字典
        """
        # 使用工厂创建数据服务实例
        self.project_data_service = DataServiceFactory.create_project_data_service(
            api_client, field_mapping, shenhua_field_mapping
        )
        self.excel_data_service = DataServiceFactory.create_excel_data_service()
        self.io_point_data_service = DataServiceFactory.create_io_point_data_service()
        
        # 保存上传的IO点表数据
        self.uploaded_io_data = None
        self.uploaded_io_file_path = None
    
    def search_project_data(self, project_number):
        """
        执行项目查询
        
        Args:
            project_number: 项目编号
            
        Returns:
            项目数据列表
        """
        if not project_number:
            raise ValueError("项目编号不能为空")
            
        # 调用项目数据服务查询数据
        return self.project_data_service.search_project_data(project_number)
    
    def clear_data(self):
        """清空所有数据"""
        self.project_data_service.clear_data()
        self.uploaded_io_data = None
        self.uploaded_io_file_path = None
    
    def load_equipment_data(self, project):
        """
        加载设备清单数据
        
        Args:
            project: 项目数据
            
        Returns:
            equipment_list: 设备列表
            station_name: 场站名称
            project_number: 项目编号
        """
        return self.project_data_service.load_equipment_data(project)
    
    @property
    def current_equipment_data(self):
        """获取当前设备数据"""
        return self.project_data_service.current_equipment_data
    
    @property
    def project_data(self):
        """获取项目数据"""
        return self.project_data_service.project_data
    
    def generate_io_table(self, output_path):
        """
        生成IO点表并导出到Excel
        
        Args:
            output_path: Excel输出路径
            
        Returns:
            是否成功生成
        """
        if not self.current_equipment_data:
            raise ValueError("没有可用的设备数据")
        
        # 检查文件是否可写
        try:
            with open(output_path, 'a') as f:
                pass  # 只是测试文件是否可写
        except PermissionError:
            raise PermissionError(f"无法写入文件: {output_path}")
        except Exception as e:
            raise Exception(f"文件访问错误: {str(e)}")
            
        # 使用Excel数据服务导出到Excel
        return self.excel_data_service.export_io_table_to_excel(self.current_equipment_data, output_path)
    
    def upload_io_table(self, input_path, callback_update_progress=None):
        """
        上传已补全信息的IO点表Excel文件
        
        Args:
            input_path: Excel文件路径
            callback_update_progress: 更新进度的回调函数
            
        Returns:
            是否成功上传，以及可能的错误信息
        """
        if not self.current_equipment_data:
            raise ValueError("没有可用的设备数据")
            
        # 更新进度
        if callback_update_progress:
            callback_update_progress("正在读取Excel文件，请稍候...")
            
        # 使用Excel数据服务读取Excel文件
        df = self.excel_data_service.read_io_table(input_path)
        
        # 更新进度
        if callback_update_progress:
            callback_update_progress("正在验证数据，请稍候...")
        
        # 使用Excel数据服务验证数据
        validation_result = self.excel_data_service.validate_io_table(df)
        
        if validation_result["has_errors"]:
            return False, validation_result["errors"]
        
        # 验证通过，继续处理
        # 存储上传的点表数据
        self.uploaded_io_data = df
        self.uploaded_io_file_path = input_path
        
        return True, {"rows": len(df), "filename": os.path.basename(input_path)}
    
    def generate_hmi_io_table(self, root_window):
        """
        生成HMI上位点表并上传到简道云
        
        Args:
            root_window: 主窗口，用于显示进度或错误
        """
        if not self.current_equipment_data:
            raise ValueError("没有可用的设备数据")
            
        if self.uploaded_io_data is None:
            raise ValueError("请先上传已补全信息的IO点表")
            
        # 使用IO点表数据服务生成HMI点表
        hmi_data = self.io_point_data_service.generate_hmi_table(self.uploaded_io_data)
        
        # 导入上传模块
        from upload import upload_hmi_table
        
        # 调用生成并上传HMI点表的函数
        upload_hmi_table(
            io_data=hmi_data,
            root_window=root_window
        )
    
    def generate_plc_io_table(self, root_window):
        """
        生成PLC下位点表并上传到简道云
        
        Args:
            root_window: 主窗口，用于显示进度或错误
        """
        if not self.current_equipment_data:
            raise ValueError("没有可用的设备数据")
            
        if self.uploaded_io_data is None:
            raise ValueError("请先上传已补全信息的IO点表")
            
        # 使用IO点表数据服务生成PLC点表
        plc_data = self.io_point_data_service.generate_plc_table(self.uploaded_io_data)
        
        # 导入上传模块
        from upload import upload_plc_table
        
        # 调用生成并上传PLC点表的函数
        upload_plc_table(
            io_data=plc_data,
            root_window=root_window
        ) 