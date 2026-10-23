# XForm - 工资表处理工具集

一套 Python 工具集，用于处理东软人事系统导出的 Excel 工资表，涵盖**格式转换、人员差异对比、打印布局调整**等场景。

## 项目结构

```
xform/
├── converter/                    # 东软 → 薪福通 工资表格式转换工具
│   ├── excel_converter_gui.py
│   ├── excel_converter_gui-comment.py
│   ├── runXinFuTong.ps1
│   ├── staff_info/               # 员工信息格式转换工具（CSV→模板）
│   ├── README.md
│   └── requirements.txt
│
├── PayrollComparator/            # 工资表人员差异对比工具
│   ├── PayrollComparator.py
│   ├── prcomp.py
│   ├── README.md
│   └── requirements.txt
│
└── PrintLayoutTool/              # 工资表打印格式批量调整工具
    └── zhanglihui20250521/
        └── win32com-wps/         # win32com 版（需安装 Excel/WPS）
```

## 子项目说明

| 子项目 | 功能 | 主要依赖 |
|--------|------|----------|
| `converter/` | 东软工资表 → 薪福通模板格式转换，含电话自动匹配、增员识别 | xlrd, xlsxwriter, openpyxl |
| `converter/staff_info/` | CSV 员工信息清洗、身份证校验、去重、比对新增人员 | openpyxl |
| `PayrollComparator/` | 对比两月工资表人员差异（新增/减少），生成差异报告 | pandas, openpyxl |
| `PrintLayoutTool/` | 批量调整工资表打印布局、隐藏/删除列、修改签字文本 | pywin32 |

每个子项目均有独立的 `README.md` 和 `requirements.txt`，可按需进入对应目录查看详情。

## 系统要求

- Python 3.8+
- Windows 系统（PrintLayoutTool 需 Microsoft Excel 或 WPS）

## 快速开始

```bash
# 进入对应子项目目录
cd converter

# 安装依赖
pip install -r requirements.txt

# 运行
python excel_converter_gui.py
```

各工具的详细使用说明见对应子项目的 README.md。
