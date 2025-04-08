# 深化设计数据查询工具

## 项目介绍

这是一个用于查询简道云API中深化设计数据的工具，提供友好的用户界面，支持项目查询和详细数据展示功能。本工具还能够生成PLC点表和亚控HMI点表，实现自动化数据处理。

## 功能特点

- 根据项目编号查询相关项目信息
- 显示项目基本信息和相关深化清单数据
- 自动生成PLC点表和亚控HMI点表
- 支持Excel导出功能
- 完善的错误处理和用户提示
- 模块化设计，易于维护和扩展
- 基于PySide6的现代用户界面，提供良好跨平台体验

## 系统要求

- Python 3.6+
- 操作系统: Windows, macOS, Linux
- 依赖的第三方库: 
  - requests 及其依赖 (urllib3, certifi, charset-normalizer, idna)
  - requests-toolbelt
  - pandas
  - openpyxl
  - xlrd, xlwt (用于Excel处理)
  - pywin32 (仅Windows系统需要)
  - PySide6 (用于用户界面)

## 安装依赖

```bash
# 安装所有依赖
pip install -r requirements.txt
```

## 配置说明

在运行程序之前，请先检查并配置以下参数:

1. 在 `config/api_config.py` 文件中配置简道云API参数:
   - API_KEY: 简道云API密钥
   - APP_ID: 应用ID
   - ENTRY_ID: 表单ID
   - FIELD_MAPPING: 字段映射
   - SHENHUA_FIELD_MAPPING: 深化清单字段映射
   - SHENHUA_SUBFORM_FIELD_ID: 深化清单子表单字段ID

## 运行程序

```bash
# 使用启动脚本
python run.py

# 或直接运行
python app.py
```

## 项目结构

```
项目根目录/
├── app.py                   # 主程序入口，初始化应用并启动界面
├── run.py                   # 启动脚本
├── api.py                   # API模块，封装与简道云API的交互
├── ui_pyside.py             # PySide6用户界面模块
├── controller.py            # 控制器模块，处理业务逻辑
├── data_service.py          # 数据服务层，处理数据加工
├── io_generator.py          # IO点表生成器，包含字段定义和IO计算逻辑
├── hmi_generator.py         # HMI点表生成器，用于生成亚控HMI点表
├── plc_generator.py         # PLC点表生成器，用于生成PLC点表
├── upload.py                # 文件上传模块，负责将点表上传到简道云
├── config/                  # 配置目录
│   ├── api_config.py        # API配置文件，包含API凭证和字段映射
│   └── settings.py          # 系统设置文件，包含路径和UI配置
├── templates/               # 模板文件目录，存放点表模板
├── requirements.txt         # 依赖库列表
├── README.md                # 项目说明文档
└── 字段.txt                 # 简道云字段映射参考文件
```

## 模块说明

### app.py
主程序入口，负责初始化和启动应用。
- 创建简道云API客户端实例
- 初始化PySide6应用程序
- 创建应用界面实例并启动
- 处理全局异常

### run.py
启动脚本，简化应用启动。
- 检查依赖是否安装
- 提供友好的启动信息
- 错误处理和异常捕获

### ui_pyside.py
PySide6用户界面模块，实现GUI界面:
- ProjectQueryApp: 主界面类
  - 项目查询功能
  - 结果展示
  - 深化清单展示
  - IO表生成功能
  - 导出功能
- DetailWindow: 详情窗口类
  - 展示单个项目的深化清单详情
  - 展示场站基本信息

### api.py
API模块，封装与简道云API的交互。
- JianDaoYunAPI类: 封装简道云API访问方法
  - 项目数据查询
  - 深化清单详情获取
  - 批量数据获取

### controller.py
控制器模块，处理业务逻辑。
- ProjectController: 项目控制器
  - 查询项目
  - 获取深化清单
  - 处理数据逻辑

### data_service.py
数据服务层，处理数据加工。
- 数据处理和转换
- 业务规则实现

### io_generator.py
IO点表生成器，包含字段定义和IO计算逻辑。
- FormFields: 表单字段数据结构，定义了所有表单字段
  - MainForm: 主表单字段定义
  - ShenhuaForm: 深化清单子表单字段定义
- IOChannelModels: IO通道模型配置类
- IOChannelCalculator: IO通道计算工具类
  - IO通道计算
  - Modbus地址计算
  - Excel导出功能

### hmi_generator.py
HMI点表生成器，负责生成亚控HMI点表。
- HMIGenerator: HMI点表生成器类
  - generate_hmi_table: 生成HMI布尔类型点表
  - generate_io_real: 生成HMI实数类型点表
  - 支持Excel格式处理和进度显示

### plc_generator.py
PLC点表生成器，负责生成PLC点表。
- PLCGenerator: PLC点表生成器类
  - generate_plc_table: 生成PLC点表
  - 支持Excel格式处理和进度显示

### upload.py
文件上传模块，负责将生成的点表上传到简道云平台。
- upload_file: 上传文件到简道云
- upload_hmi_table: 生成并上传HMI点表
- upload_plc_table: 生成并上传PLC点表
- upload_fat_table: 生成并上传FAT点表

### config/
配置目录，包含系统配置文件。
- api_config.py: API配置文件
  - API凭证：API_KEY, APP_ID, ENTRY_ID
  - 字段映射
- settings.py: 系统设置文件
  - UI配置
  - 文件路径

### templates/
模板文件目录，存放生成点表所需的模板文件。
- PLC点表模板
- HMI点表模板

## 工作流程

1. 用户启动应用程序(run.py或app.py)
2. 在主界面输入项目编号进行查询
3. 查询结果显示匹配的项目信息
4. 用户选择项目查看深化清单详情
5. 用户可以生成IO点表、PLC点表或HMI点表
6. 用户可以将生成的点表上传到简道云平台
7. 结果可导出为Excel文件保存

## 错误排查

如果程序运行时遇到问题:

1. 确认配置文件中的API凭证是否正确
2. 检查网络连接是否正常
3. 查看控制台输出的错误信息
4. 确保Excel模板文件存在且格式正确
5. Windows系统需确保安装了pywin32库
6. 确保安装了xlrd和xlwt库用于Excel处理
7. 确保正确安装了PySide6库 (可使用 `pip install PySide6` 安装)

## 开发者说明

代码采用模块化设计，将数据访问、用户界面和配置分离，便于后期维护和扩展:

1. 数据结构: 使用FormFields类组织表单字段，便于增删改查
2. 数据访问: JianDaoYunAPI类封装所有API交互
3. 用户界面: ui_pyside.py模块实现基于PySide6的现代化用户界面
4. 业务逻辑: 通过Controller控制器层处理业务流程
5. 点表生成: 独立的生成器模块，处理不同类型点表的生成逻辑

## 许可证

MIT License