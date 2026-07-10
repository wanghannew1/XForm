# 员工信息格式转换工具

## 概述

将东软人事系统导出的 CSV 员工信息文件，转换为薪福通 `员工导入模板202512.xlsx` 格式。

支持身份证号校验、手机号格式清洗、重复数据去重（保留最新记录）、新增人员比对等功能。

## 功能

- 读取 CSV 员工数据（支持 gb18030 编码）
- 身份证号格式和校验码验证
- 手机号格式清洗和有效性校验（号段、假号码检测）
- 按身份证号去重，保留最新经办日期的记录
- 自动生成 `异常数据` sheet（标黄显示异常项）
- 比对上次生成的文件，自动识别新增人员（标绿）
- 支持命令行模式和 GUI 拖拽模式

## 使用方式

### GUI 模式（默认）

```bash
python employee_info_converter.py
```

### 命令行模式

```bash
python employee_info_converter.py --cli --csv 数据文件.csv
python employee_info_converter.py --cli --csv 数据文件.csv --compare   # 比对新增人员
python employee_info_converter.py --cli -i 数据.csv -o 输出.xlsx -c    # 完整参数
```

## 依赖

- openpyxl - 读写 .xlsx 模板文件
- tkinterdnd2 - GUI 拖拽支持（可选，命令行模式不需要）

## 配置

配置文件 `employee_info_converter_config.json` 自动生成，包含：
- `csv_encoding` - CSV 文件编码（默认 gb18030）
- `id_card_col` / `name_col` / `phone_col` - 各列索引
- `sheet_name` - 模板中花名册 sheet 名称

## 文件说明

- `employee_info_converter.py` - 主程序
- `员工导入模板202512.xlsx` - 薪福通标准模板文件（需放置在脚本同目录）
- `Cleaned/` - 转换结果输出目录
