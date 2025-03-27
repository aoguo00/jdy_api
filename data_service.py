#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
深化设计数据查询工具 - 数据服务模块

此模块实现了应用程序的数据获取、转换和处理功能，
作为控制器层和数据访问层之间的中间层，
负责数据的格式化、转换和验证。
"""

import pandas as pd
from io_generator import IOChannelCalculator, FormFields

class ProjectDataService:
    """
    项目数据服务类
    负责项目和设备数据的获取和处理
    """
    
    def __init__(self, api_client, field_mapping, shenhua_field_mapping):
        """
        初始化项目数据服务
        
        Args:
            api_client: 简道云API客户端实例
            field_mapping: 字段映射字典
            shenhua_field_mapping: 深化清单字段映射字典
        """
        self.api_client = api_client
        self.field_mapping = field_mapping
        self.shenhua_field_mapping = shenhua_field_mapping
        
        # 保存项目数据
        self.project_data = []
        
        # 保存所有设备数据，用于本地筛选
        self.all_equipment_data = []
        
        # 保存所有深化清单数据（所有项目）
        self.all_shenhua_data = []
        
        # 保存当前选中项目的设备数据
        self.current_equipment_data = []
    
    def search_project_data(self, project_number):
        """
        执行项目查询
        
        Args:
            project_number: 项目编号
            
        Returns:
            项目数据列表
        """
        # 调用API查询数据
        self.project_data = self.api_client.search_project_data(project_number)
        return self.project_data
    
    def clear_data(self):
        """清空所有数据"""
        self.project_data = []
        self.all_equipment_data = []
        self.all_shenhua_data = []
        self.current_equipment_data = []
    
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
        # 获取项目ID和场站信息
        data_id = project.get('_id', '')
        if not data_id:
            raise ValueError("无法获取项目数据ID (_id)")
        
        # 获取场站名称和项目编号
        field_id = self.field_mapping.get(FormFields.MainForm.STATION)  # 使用FormFields中定义的常量
        station_name = project.get(field_id, "[未知场站]")
        
        field_id = self.field_mapping.get(FormFields.MainForm.PROJECT_NUMBER)  # 使用FormFields中定义的常量
        project_number = project.get(field_id, "[未知项目]")
        
        # 如果缓存为空则加载所有数据
        if not self.all_shenhua_data:
            # 获取所有深化清单数据
            self.all_shenhua_data = self.api_client.get_all_shenhua_data()
        
        # 筛选出当前选中项目的数据
        selected_data = None
        for item in self.all_shenhua_data:
            if item.get('_id') == data_id:
                selected_data = item
                break
        
        if not selected_data or self.api_client.subform_field_id not in selected_data:
            return None, station_name, project_number
        
        # 获取子表单数据
        detail_data = selected_data.get(self.api_client.subform_field_id, [])
        
        if not detail_data:
            return None, station_name, project_number
        
        # 将获取到的数据转换为显示格式并添加到设备列表中
        equipment_list = []
        for item in detail_data:
            formatted_item = {}
            for field_name, field_id in self.shenhua_field_mapping.items():
                formatted_item[field_name] = item.get(field_id, "")
            
            # 添加该设备数据所属的场站信息，便于区分
            formatted_item["_station"] = station_name
            formatted_item["_project"] = project_number
            formatted_item["_data_id"] = data_id
            equipment_list.append(formatted_item)
        
        # 保存当前选中项目的设备数据，用于生成IO点表
        self.current_equipment_data = equipment_list
        
        return equipment_list, station_name, project_number


class ExcelDataService:
    """
    Excel数据服务类
    负责Excel相关的数据处理、验证和导出
    """
    
    @staticmethod
    def export_io_table_to_excel(equipment_data, output_path):
        """
        将IO点表导出到Excel
        
        Args:
            equipment_data: 设备数据列表
            output_path: 输出文件路径
            
        Returns:
            是否成功导出
        """
        # 计算IO通道数量和数据类型
        channel_data = IOChannelCalculator.calculate_channels(equipment_data)
        
        # 导出到Excel，传递设备列表获得更详细的报表
        return IOChannelCalculator.export_to_excel(channel_data, output_path, equipment_data)
    
    @staticmethod
    def validate_io_table(df):
        """
        验证IO点表数据
        
        Args:
            df: pandas DataFrame，包含IO点表数据
            
        Returns:
            验证结果字典
        """
        # 验证必填字段
        required_fields = ["变量名称（HMI）", "变量描述"]
        
        # 创建错误信息列表
        missing_fields = []
        invalid_power_type = []
        invalid_wire_type_bool = []
        invalid_wire_type_real = []
        missing_values_real = []
        
        # 验证每一行
        for idx, row in df.iterrows():
            row_num = idx + 2  # Excel行号从2开始（跳过表头）
            
            # 获取此行的数据类型
            data_type = row.get("数据类型")
            
            # 1. 检查必填字段
            for field in required_fields:
                if field in df.columns and pd.isna(row[field]):
                    missing_fields.append(f"第{row_num}行: {field}为空")
            
            # 2. 验证供电类型
            power_type = row.get("供电类型（有源/无源）")
            module_type = row.get("模块类型", "")
            # 对于AO类型模块，不进行供电类型验证
            if module_type != "AO":
                if pd.isna(power_type) or str(power_type).strip() == "":
                    invalid_power_type.append(f"第{row_num}行: 供电类型为空")
                elif str(power_type) not in ["有源", "无源"]:
                    invalid_power_type.append(f"第{row_num}行: 供电类型必须是'有源'或'无源'，当前值: {power_type}")
            
            # 3. 验证线制
            wire_type = row.get("线制")
            if data_type == "BOOL":
                if pd.isna(wire_type) or str(wire_type).strip() == "":
                    invalid_wire_type_bool.append(f"第{row_num}行: 线制为空")
                elif str(wire_type) not in ["常开", "常闭"]:
                    invalid_wire_type_bool.append(f"第{row_num}行: BOOL类型的线制必须是'常开'或'常闭'，当前值: {wire_type}")
            elif data_type == "REAL":
                if pd.isna(wire_type) or str(wire_type).strip() == "":
                    invalid_wire_type_real.append(f"第{row_num}行: 线制为空")
                elif str(wire_type) not in ["2线制", "二线制", "三线制", "四线制","两线制"]:
                    invalid_wire_type_real.append(f"第{row_num}行: REAL类型的线制必须是'2线制'、'二线制'、'三线制'或'四线制'，当前值: {wire_type}")
            
            # 4. 如果是REAL类型，验证设定值是否填写
            if data_type == "REAL":
                set_point_fields = ["SLL设定值", "SL设定值", "SH设定值", "SHH设定值"]
                for field in set_point_fields:
                    if field in df.columns and pd.isna(row[field]):
                        missing_values_real.append(f"第{row_num}行: {field}为空")
        
        # 合并所有错误信息
        all_errors = []
        
        if missing_fields:
            all_errors.append("必填字段缺失:\n" + "\n".join(missing_fields))
        
        if invalid_power_type:
            all_errors.append("供电类型错误:\n" + "\n".join(invalid_power_type))
        
        if invalid_wire_type_bool:
            all_errors.append("数字量线制错误:\n" + "\n".join(invalid_wire_type_bool))
        
        if invalid_wire_type_real:
            all_errors.append("模拟量线制错误:\n" + "\n".join(invalid_wire_type_real))
        
        if missing_values_real:
            all_errors.append("模拟量设定值缺失:\n" + "\n".join(missing_values_real))
        
        return {
            "has_errors": len(all_errors) > 0,
            "errors": all_errors
        }
    
    @staticmethod
    def read_io_table(file_path, sheet_name="IO点表"):
        """
        读取IO点表Excel文件
        
        Args:
            file_path: Excel文件路径
            sheet_name: 工作表名称
            
        Returns:
            pandas DataFrame
        """
        return pd.read_excel(file_path, sheet_name=sheet_name)


class IOPointDataService:
    """
    IO点表数据服务类
    负责IO点表相关的数据处理和转换
    """
    
    @staticmethod
    def generate_hmi_table(io_data):
        """
        生成HMI上位点表数据
        
        Args:
            io_data: 原始IO点表数据(DataFrame)
            
        Returns:
            转换后的HMI点表数据(DataFrame)
        """
        # 这里实现HMI点表生成逻辑
        # 可以进行字段映射、数据转换等操作
        
        # 暂时返回原始数据，后续实现具体转换逻辑
        return io_data
    
    @staticmethod
    def generate_plc_table(io_data):
        """
        生成PLC下位点表数据
        
        Args:
            io_data: 原始IO点表数据(DataFrame)
            
        Returns:
            转换后的PLC点表数据(DataFrame)
        """
        # 这里实现PLC点表生成逻辑
        # 可以进行字段映射、数据转换等操作
        
        # 暂时返回原始数据，后续实现具体转换逻辑
        return io_data


# 工厂类，用于创建和管理各种数据服务实例
class DataServiceFactory:
    """
    数据服务工厂类
    用于创建和管理各种数据服务实例
    """
    
    @staticmethod
    def create_project_data_service(api_client, field_mapping, shenhua_field_mapping):
        """创建项目数据服务实例"""
        return ProjectDataService(api_client, field_mapping, shenhua_field_mapping)
    
    @staticmethod
    def create_excel_data_service():
        """创建Excel数据服务实例"""
        return ExcelDataService()
    
    @staticmethod
    def create_io_point_data_service():
        """创建IO点表数据服务实例"""
        return IOPointDataService() 