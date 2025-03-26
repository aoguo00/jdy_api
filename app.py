#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
深化设计数据查询工具 - 主程序入口
该文件是系统的主入口点，负责初始化应用程序并启动界面
"""

import tkinter as tk
from tkinter import messagebox

# 导入配置
from config.api_config import (
    API_KEY, APP_ID, ENTRY_ID, FIELD_MAPPING, 
    SHENHUA_FIELD_MAPPING, SHENHUA_SUBFORM_FIELD_ID,
    get_app_id, get_entry_id, get_field_mapping
)
from config.settings import UI_TITLE, UI_WIDTH, UI_HEIGHT

# 从api模块导入JianDaoYunAPI
from api import JianDaoYunAPI
from ui import ProjectQueryApp

def main():
    """
    主程序入口函数
    负责初始化API客户端并启动用户界面
    包含基本的错误处理机制
    """
    try:
        # 检查API凭证是否已设置
        if not API_KEY:
            messagebox.showerror("配置错误", "API凭证未配置\n请在 config/api_config.py 文件中设置有效的 API_KEY")
            return
        
        # 初始化API客户端
        api_client = JianDaoYunAPI(
            api_key=API_KEY,
            app_id=APP_ID,
            entry_id=ENTRY_ID,
            field_mapping=FIELD_MAPPING,
            subform_field_id=SHENHUA_SUBFORM_FIELD_ID
        )
        
        # 创建Tkinter根窗口
        root = tk.Tk()
        root.title(UI_TITLE)
        root.geometry(f"{UI_WIDTH}x{UI_HEIGHT}")
        
        # 创建应用实例
        app = ProjectQueryApp(
            root=root,
            api_client=api_client,
            field_mapping=FIELD_MAPPING,
            shenhua_field_mapping=SHENHUA_FIELD_MAPPING
        )
        
        # 启动主循环
        root.mainloop()
        
    except KeyboardInterrupt:
        pass  # 忽略键盘中断
    except Exception as e:
        messagebox.showerror("程序错误", f"程序发生错误: {str(e)}")

if __name__ == "__main__":
    main() 