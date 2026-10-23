# 东软-薪福通工资表格式转换工具

## 概述

将东软人事系统导出的工资表（.xls / .xlsx 格式）转换为薪福通标准模板格式。支持文件选择、拖放、批量转换，并提供 GUI 图形界面操作。

## 目录结构

```
converter/
├── excel_converter_gui.py          # 主程序（GUI 版）
├── excel_converter_gui-comment.py  # 带详细注释的同版本（含变更记录）
├── runXinFuTong.ps1                # Windows 启动脚本
├── staff_info/                     # 员工信息格式转换工具（独立子工具）
├── requirements.txt                # Python 依赖
└── README.md
```

## 功能

- 自动读取东软工资表（支持 .xls 和 .xlsx）
- 按预设字段映射转换列名和格式
- 自动从身份证信息查询手机号（支持 CSV/XLSX 电话查询表）
- 检测有数据的动态列并输出
- 提取发薪月份、人员信息校验
- 比对新增人员并自动创建花名册 sheet
- 可编辑的字段映射配置（通过 GUI 配置编辑器）
- 批量处理多个文件

## 安装与使用

```bash
# 安装依赖
pip install -r requirements.txt

# 运行
python excel_converter_gui.py

# 或使用 Windows 启动脚本
powershell -ExecutionPolicy ByPass -c ".\runXinFuTong.ps1"
```

## 依赖

- xlrd - 读取旧版 .xls 文件
- xlsxwriter - 写入 .xlsx 文件
- openpyxl - 读写 .xlsx 文件
- tkinterdnd2 - 拖拽文件支持

## 配置

配置文件 `dongr-xinft_config.json` 自动生成在脚本同目录，包含：
- 字段映射（源列名 → 目标列名）
- 电话查询表配置（CSV/XLSX 列名映射）
- 源表格式配置（标题行、表头行数、数据起始行等）
- 固定列和动态列配置
