#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
深化设计数据查询工具 - 主程序入口
该文件是系统的主入口点，负责初始化应用程序并启动界面
"""

import sys
import traceback
from PySide6.QtWidgets import QApplication, QMessageBox

# 导入配置
from config.api_config import (
    API_KEY, APP_ID, ENTRY_ID, FIELD_MAPPING, 
    SHENHUA_FIELD_MAPPING, SHENHUA_SUBFORM_FIELD_ID,
    get_app_id, get_entry_id, get_field_mapping
)
from config.settings import UI_TITLE, UI_WIDTH, UI_HEIGHT

# 从api模块导入JianDaoYunAPI
from api import JianDaoYunAPI
from ui_pyside import ProjectQueryApp

def main():
    """
    主程序入口函数
    负责初始化API客户端并启动用户界面
    包含基本的错误处理机制
    """
    print("深化设计数据查询工具启动中...")
    print("=" * 50)
    
    try:
        # 检查PySide6是否已安装
        try:
            import PySide6
            print("正在启动界面...")
        except ImportError:
            print("导入PySide6失败，请确保已安装PySide6。")
            print("尝试使用命令安装: pip install PySide6")
            return 1
            
        # 检查API凭证是否已设置
        if not API_KEY:
            QMessageBox.critical(None, "配置错误", "API凭证未配置\n请在 config/api_config.py 文件中设置有效的 API_KEY")
            return 1
        
        # 初始化API客户端
        api_client = JianDaoYunAPI(
            api_key=API_KEY,
            app_id=APP_ID,
            entry_id=ENTRY_ID,
            field_mapping=FIELD_MAPPING,
            subform_field_id=SHENHUA_SUBFORM_FIELD_ID
        )
        
        # 创建QApplication实例
        app = QApplication(sys.argv)
        app.setApplicationName(UI_TITLE)
        
        # 创建应用实例
        main_window = ProjectQueryApp(
            api_client=api_client,
            field_mapping=FIELD_MAPPING,
            shenhua_field_mapping=SHENHUA_FIELD_MAPPING
        )
        main_window.show()
        
        # 启动主循环
        return app.exec()
        
    except KeyboardInterrupt:
        print("程序被用户中断")
        return 0
    except Exception as e:
        error_message = f"程序发生错误: {str(e)}\n\n{traceback.format_exc()}"
        print(error_message)
        QMessageBox.critical(None, "程序错误", error_message)
        return 1

if __name__ == "__main__":
    sys.exit(main()) 