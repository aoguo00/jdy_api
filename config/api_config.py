"""
API配置
这个文件包含所有与API交互相关的配置信息
包括API凭证、应用信息以及字段映射关系
"""

# ===================== 通用 API 配置 =====================

# 简道云API凭证（公共部分）
API_KEY = "WuVMLm7r6s1zzFTkGyEYXQGxEZ9mLj3h"  # API密钥

# API端点
API_BASE_URL = "https://api.jiandaoyun.com/api"
API_V5_BASE_URL = "https://api.jiandaoyun.com/api/v5"  # 新版API

# ===================== 应用配置 =====================

# 应用配置字典 - 包含所有应用及其相关表单的配置
APPLICATIONS = {
    # 深化设计应用
    "深化设计": {
        "APP_ID": "67d13e0bb840cdf11eccad1e",
        "表单": {
            "主表单": {
                "ENTRY_ID": "67d7f0ed97abe5bfc70d8aed",
                "字段映射": {
                    "项目名称": "_widget_1635777114903",
                    "项目编号": "_widget_1635777114935",
                    "深化设计编号": "_widget_1636359817201",
                    "客户名称": "_widget_1635777114972",
                    "场站": "_widget_1635777114991"
                },
                "子表单": {
                    "深化清单": {
                        "字段ID": "_widget_1635777115095",
                        "字段映射": {
                            "设备名称": "_widget_1635777115211",
                            "品牌": "_widget_1635777115248",
                            "规格型号": "_widget_1635777115287",
                            "技术参数": "_widget_1641439264111",
                            "数量": "_widget_1635777485580",
                            "单位": "_widget_1654703913698",
                            "子系统": "_widget_1636353456514",
                            "备注": "_widget_1635777854826",
                            "技术备注": "_widget_1666709244379",
                            "合同内外": "_widget_1684760244471"
                        }
                    }
                }
            }
        }
    },
    
    # 文件上传应用
    "文件上传": {
        "APP_ID": "67d13e0bb840cdf11eccad1e",
        "表单": {
            "文件目录": {
                "ENTRY_ID": "67de5421feff5fa84e9b0a24",
                "字段映射": {
                    "上传字段": "_widget_1742623777364",
                    "文件类型": "_widget_1742624119037"
                }
            }
        }
    }
}

# ===================== 便捷访问函数 =====================

def get_app_id(app_name):
    """获取指定应用的APP_ID"""
    return APPLICATIONS.get(app_name, {}).get("APP_ID")

def get_entry_id(app_name, form_name):
    """获取指定应用下指定表单的ENTRY_ID"""
    return APPLICATIONS.get(app_name, {}).get("表单", {}).get(form_name, {}).get("ENTRY_ID")

def get_field_mapping(app_name, form_name):
    """获取指定应用下指定表单的字段映射"""
    return APPLICATIONS.get(app_name, {}).get("表单", {}).get(form_name, {}).get("字段映射", {})

def get_field_id(app_name, form_name, field_name):
    """获取指定应用下指定表单中指定字段的ID"""
    field_mapping = get_field_mapping(app_name, form_name)
    return field_mapping.get(field_name)

# 为了向后兼容，保留原有变量名
# 深化设计应用
APP_ID = get_app_id("深化设计")
ENTRY_ID = get_entry_id("深化设计", "主表单")
FIELD_MAPPING = get_field_mapping("深化设计", "主表单")
SHENHUA_SUBFORM_FIELD_ID = APPLICATIONS["深化设计"]["表单"]["主表单"]["子表单"]["深化清单"]["字段ID"]
SHENHUA_FIELD_MAPPING = APPLICATIONS["深化设计"]["表单"]["主表单"]["子表单"]["深化清单"]["字段映射"]

