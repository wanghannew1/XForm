# legacy - 历史遗留版本

本目录包含早期开发的工资表处理工具原型，均已由新版替代，仅供参考。

## 文件说明

| 文件 | 说明 | 技术栈 |
|------|------|--------|
| `XForm.py` | 最早的 PyQt5 版本 | PyQt5, xlwings |
| `XForm2.py` | tkinter + xlwings 重构版 | tkinter, xlwings |
| `XForm-openpyxl.py` | 纯 openpyxl 版本（不需 Excel 环境） | openpyxl, xlrd |
| `XFormtest.ipynb` | 功能测试/原型 Notebook | Jupyter |

## 当前推荐

- **打印格式调整** → `PrintLayoutTool/` 目录下的新版
- **格式转换** → `converter/` 目录下的东软-薪福通转换工具
- **人员对比** → `PayrollComparator/` 目录下的差异对比工具
