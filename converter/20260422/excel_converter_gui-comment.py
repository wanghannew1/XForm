#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
东软-薪福通工资表格式转换工具 - GUI 版本
支持文件选择、拖放、批量转换

================================================================================
变更记录（带详细注释版本）
================================================================================

【问题1】第1行大类列头生成错误
--------------------------------------
现象：第1行（薪福通大类列头）显示为列名而不是类别
     例如：第6列显示"基本工资"而不是"总计"
根因：final_headers_1 = column_config['fixed_columns'].copy() 复制了列名
修复：将列名通过fixed_category映射转换为大类
     final_headers_1 = [fixed_category.get(col, col) for col in fixed_columns]

【问题2】未配置列（未知列）数据不写入
--------------------------------------
现象：第25-32列只有列头，没有数据
根因：原代码先检查"if col_name in field_index: continue"
     unknown列也在field_index中，直接被跳过，第709行根本不会执行
修复：
  1. 删除field_index检查
  2. 调整执行顺序：先跳过固定列 -> 处理动态列 -> 处理未知列
  3. 添加output_unknown_columns开关控制是否输出未知列

【问题3】输出未知列开关
--------------------------------------
新增配置项：output_unknown_columns
  True = 输出未知列（原始表有数据但不在配置中的列）
  False = 不输出未知列（默认关闭）
涉及位置：
  - get_default_config() 默认False
  - dongr-xinft_config.json 默认false
  - convert_single_file() 两处逻辑
  - ConfigEditor 复选框

【问题4】字段映射tab页滚轮失效
--------------------------------------
修复：在setup_mapping_tab()中添加鼠标滚轮事件绑定
     def on_mousewheel(event):
         self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
     self.canvas.bind_all("<MouseWheel>", on_mousewheel)

【问题5】配置编辑器bug（取消输出未知列报错）
--------------------------------------
错误信息：AttributeError: 'ConfigEditor' object has no attribute 'header_row_entry'
根因：用户未访问源表配置选项卡时，相关widget未创建
修复：在confirm()方法中使用getattr和hasattr进行防御性检查

【问题6】固定列添加功能缺失
--------------------------------------
修复：
  1. 在按钮栏添加"添加"按钮
  2. 新增fixed_add()方法，弹出对话框输入新列名

【问题7】方案A - 配置合并
--------------------------------------
将 DEFAULT_FIELD_MAPPING (24项) 和 dynamic_columns (46项) 合并为 column_mapping (约70项)
删除旧的 source_field_mapping 和 dynamic_columns 配置项
配置编辑器界面重构：
  - 删除原来的"字段映射"tab（旧样式）和"动态列"tab
  - 新建统一的"字段映射"tab（Treeview样式）
  - 加载column_mapping数据，支持添加/编辑/删除/排序

【问题8】动态列不输出bug
--------------------------------------
根因：detect_dynamic_and_unknown_columns() 错误地将所有column_mapping中的source都加入fixed_source_columns
     导致动态列也被视为"已处理"而跳过
修复：只将target在fixed_columns中的列视为固定列
     fixed_target_set = set(column_config.get('fixed_columns', []))
     for src, tgt in source_to_target.items():
         if tgt in fixed_target_set:  # 只添加目标在fixed_columns中的映射
             fixed_source_columns.add(src)
             fixed_source_columns.add(tgt)

================================================================================
"""

import os
import re
import csv
import json
import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import tkinterdnd2
from tkinterdnd2 import DND_FILES
import xlrd
import xlsxwriter
from typing import Dict, List, Tuple, Any, Optional
import openpyxl
from openpyxl.styles import Font, Alignment


# ==================== 配置部分 ====================

CONFIG_FILE = 'dongr-xinft_config.json'

# ================================================================================
# 【问题7 - 方案A：配置合并】将DEFAULT_FIELD_MAPPING和dynamic_columns合并为column_mapping
# ================================================================================
DEFAULT_COLUMN_MAPPING = [
    # 固定列映射（约21项）- 这些列的目标名在fixed_columns中
    {"source": "职工号", "target": "工号"},
    {"source": "姓名", "target": "姓名"},
    {"source": "身份证", "target": "证件号码"},
    {"source": "联系电话", "target": "手机号"},
    {"source": "工资年月", "target": "发薪月份"},
    {"source": "基本工资", "target": "基本工资"},
    {"source": "应发工资", "target": "应发工资"},
    {"source": "实发合计", "target": "实发金额"},
    {"source": "社保基数", "target": "社保基数"},
    {"source": "医保基数", "target": "医保基数"},
    {"source": "工伤基数", "target": "工伤基数"},
    {"source": "公积金基数", "target": "公积金基数"},
    {"source": "扣款明细_养老_单位", "target": "扣单位养老保险"},
    {"source": "扣款明细_养老_个人", "target": "扣个人养老保险"},
    {"source": "扣款明细_失业_单位", "target": "扣单位失业保险"},
    {"source": "扣款明细_失业_个人", "target": "扣个人失业保险"},
    {"source": "扣款明细_医疗_单位", "target": "扣单位医疗保险"},
    {"source": "扣款明细_医疗_个人", "target": "扣个人医疗保险"},
    {"source": "扣款明细_工伤险_单位", "target": "扣单位工伤险"},
    {"source": "扣款明细_公积金_单位", "target": "扣单位住房公积金"},
    {"source": "扣款明细_公积金_个人", "target": "扣个人住房公积金"},
    {"source": "个人所得税", "target": "扣缴税金（扣个人所得税）"},
    {"source": "补发工资", "target": "补发工资"},
    # 动态列映射（约46项，有数据时输出）
    {"source": "扣款明细_补缴及退款保险差额（单位）", "target": "补扣单位保险"},
    {"source": "扣款明细_补缴及退款保险差额（个人）", "target": "补扣个人保险"},
    {"source": "扣款明细_大病险合计", "target": "医保大额"},
    {"source": "公务员医疗补助", "target": "扣单位公务员医疗补助"},
    {"source": "扣款", "target": "扣款"},
    {"source": "病假扣款", "target": "病假扣款"},
    {"source": "事假扣款", "target": "事假扣款"},
    {"source": "大众班车费（扣）", "target": "大众班车费（扣）"},
    {"source": "个人缴纳现金（交）", "target": "个人缴纳现金（交）"},
    {"source": "偶然所得税（税）", "target": "偶然所得税（税）"},
    {"source": "扣伙食补助", "target": "扣伙食补助"},
    {"source": "补收单位服务费", "target": "补收单位服务费"},
    {"source": "补收个人服务费", "target": "补收个人服务费"},
    {"source": "采暖费（免）", "target": "采暖费"},
    {"source": "奖金", "target": "奖金"},
    {"source": "加班费", "target": "加班费"},
    {"source": "餐补", "target": "餐补"},
    {"source": "独生子女费（免）", "target": "独生子女费（免）"},
    {"source": "有毒有害津贴", "target": "有毒有害津贴"},
    {"source": "双一流", "target": "双一流"},
    {"source": "房租", "target": "房租"},
    {"source": "交通费用", "target": "交通费用"},
    {"source": "岗位工资", "target": "岗位工资"},
    {"source": "绩效奖金", "target": "绩效奖金"},
    {"source": "一次性奖金", "target": "一次性奖金"},
    {"source": "津贴", "target": "津贴"},
    {"source": "合计夜餐费", "target": "合计夜餐费"},
    {"source": "合计加班费", "target": "合计加班费"},
    {"source": "补发", "target": "补发"},
    {"source": "补贴2（餐现）", "target": "补贴2（餐现）"},
    {"source": "补发（体检费）", "target": "补发（体检费）"},
    {"source": "年终奖", "target": "年终奖"},
    {"source": "高危补助", "target": "高危补助"},
    {"source": "体检费", "target": "体检费"},
    {"source": "补贴", "target": "补贴"},
    {"source": "业绩津贴", "target": "业绩津贴"},
    {"source": "伙食费", "target": "伙食费"},
    {"source": "采访补助", "target": "采访补助"},
    {"source": "返餐费（返）", "target": "返餐费（返）"},
    {"source": "福利费（返）", "target": "福利费（返）"},
    {"source": "个人其他调整", "target": "个人其他调整"},
    {"source": "其他工资", "target": "其他工资"},
    {"source": "其他免税", "target": "其他免税"},
    {"source": "其他返单位", "target": "其他返单位"},
    {"source": "交通补贴", "target": "交通补贴"},
    {"source": "特殊班次补助", "target": "特殊班次补助"},
]

TARGET_HEADER_ROW1 = []
TARGET_HEADER_ROW2 = []
TARGET_FIELD_INDEX = {}

def build_target_headers(column_config: Dict = None):
    """
    构建目标Excel文件的表头（用于生成转换后的文件）
    
    参数:
        column_config: Dict - 列配置字典，包含 fixed_columns、column_mapping 等
    
    全局变量:
        TARGET_HEADER_ROW1: 第1行大类表头（总计、扣缴基数、应扣项、其它等）
        TARGET_HEADER_ROW2: 第2行列名表头（工号、姓名、基本工资等）
        TARGET_FIELD_INDEX: 列名到索引的映射字典
    
    处理逻辑：
    1. 从column_config获取fixed_columns（固定列）和column_mapping（字段映射）
    2. 生成第2行表头（列名）：
       - 先添加所有固定列
       - 再添加column_mapping中的动态列（target不在fixed_columns中）
    3. 生成第1行表头（大类）：
       - 使用fixed_category映射将列名转换为大类名称
       - 不在映射中的列归类为"其它"
    4. 构建列名到索引的映射（用于数据写入定位）
    
    分类映射规则（fixed_category）：
    - 工号、姓名、证件号码、手机号、发薪月份 -> 保持原名
    - 基本工资、补发工资、应发工资、实发金额 -> 总计
    - 社保基数、医保基数、工伤基数、公积金基数 -> 扣缴基数
    - 各类保险和公积金扣款、个税 -> 应扣项
    - 其他 -> 其它
    """
    global TARGET_HEADER_ROW1, TARGET_HEADER_ROW2, TARGET_FIELD_INDEX
    
    # 获取列配置
    if column_config is None:  # 使用默认配置
        column_config = get_default_config().get('column_config', {})
    
    fixed_columns = column_config.get('fixed_columns', [])
    
    # 从column_mapping获取动态列（target不在fixed_columns中）
    column_mapping = column_config.get('column_mapping', DEFAULT_COLUMN_MAPPING)
    fixed_set = set(fixed_columns)  # set(): 从列表创建集合（用于快速查找）
    # list comprehension: [item for item in column_mapping if condition] 筛选动态列
    dynamic_columns = [item for item in column_mapping if item['target'] not in fixed_set]
    
    # 生成第2行表头（列名）：固定列 + 动态列
    TARGET_HEADER_ROW2 = fixed_columns.copy()
    for item in dynamic_columns:
        TARGET_HEADER_ROW2.append(item['target'])
    
    # 固定列对应的大类映射（问题1修复：使用此映射生成第1行表头）
    fixed_category = {
        '工号': '工号', '姓名': '姓名', '证件号码': '证件号码', '手机号': '手机号', '发薪月份': '发薪月份',
        '基本工资': '总计', '补发工资': '总计', '应发工资': '总计', '实发金额': '总计',
        '社保基数': '扣缴基数', '医保基数': '扣缴基数', '工伤基数': '扣缴基数', '公积金基数': '扣缴基数',
        '扣单位养老保险': '应扣项', '扣个人养老保险': '应扣项', '扣单位失业保险': '应扣项', '扣个人失业保险': '应扣项',
        '扣单位医疗保险': '应扣项', '扣个人医疗保险': '应扣项', '扣单位工伤险': '应扣项', '扣单位住房公积金': '应扣项',
        '扣个人住房公积金': '应扣项', '扣缴税金（扣个人所得税）': '应扣项', '2025年个税补退': '应扣项'
    }
    
    # 生成第1行表头（大类）
    TARGET_HEADER_ROW1 = []
    for col in TARGET_HEADER_ROW2:
        if col in fixed_category:  # 在大类映射中
            TARGET_HEADER_ROW1.append(fixed_category[col])
        else:
            TARGET_HEADER_ROW1.append('其它')
    
    # 构建列名到索引的映射（用于数据写入定位）
    TARGET_FIELD_INDEX = {name: idx for idx, name in enumerate(TARGET_HEADER_ROW2)}


# 运行时字段映射（从配置加载）
FIELD_MAPPING = {}


# ==================== 配置加载/保存 ====================

def get_default_config() -> Dict:
    """
    获取默认配置字典
    
    返回值:
        Dict - 默认配置，包含以下键值：
        - phone_lookup_config: 电话查询配置
          * csv: CSV格式的列名配置（身份证列、联系电话列）
          * xlsx: XLSX格式的列名配置（身份证列、手机号列、Sheet名）
        - source_config: 源表解析配置
          * title_row: 标题行索引
          * header_row_count: 表头行数
          * data_start_row: 数据起始行
          * skip_keyword: 跳过关键字（遇到此关键字停止读取）
          * title_col: 标题所在列
        - recent_csv: 最近使用的人员信息文件路径
        - column_config: 列配置
          * output_unknown_columns: 是否输出未知列（默认False）
          * fixed_columns: 固定列列表（始终输出）
          * column_mapping: 字段映射列表（source->target）
    """
    return {
        "phone_lookup_config": {
            "csv": {
                "id_card_column": "公民身份号码",
                "phone_column": "联系电话"
            },
            "xlsx": {
                "id_card_column": "*证件号码",
                "phone_column": "*手机号",
                "sheet_name": "花名册"
            }
        },
        "source_config": {
            "title_row": 0,
            "header_row_count": 5,
            "data_start_row": 5,
            "skip_keyword": "合计",
            "title_col": 2
        },
        "recent_csv": {
            "path": "",
            "created_time": ""
        },
        "column_config": {
            "output_unknown_columns": False,
            "fixed_columns": [
                "工号", "姓名", "证件号码", "手机号", "发薪月份",
                "基本工资", "补发工资", "应发工资", "实发金额",
                "社保基数", "医保基数", "工伤基数", "公积金基数",
                "扣单位养老保险", "扣个人养老保险", "扣单位失业保险", "扣个人失业保险",
                "扣单位医疗保险", "扣个人医疗保险", "扣单位工伤险", "扣单位住房公积金",
                "扣个人住房公积金", "扣缴税金（扣个人所得税）", "2025年个税补退"
            ],
            "column_mapping": DEFAULT_COLUMN_MAPPING
        }
    }


def load_config() -> Dict:
    """
    从JSON文件加载应用程序配置
    
    返回值:
        Dict - 加载的配置字典
    
    处理逻辑：
    1. 查找配置文件（dongr-xinft_config.json，与脚本同目录）
    2. 如果文件存在且可读：解析JSON并返回
    3. 如果文件不存在或解析失败：
       - 获取默认配置
       - 保存默认配置到文件
       - 返回默认配置
    
    异常处理：
    - 捕获所有异常，使用默认配置作为后备
    """
    config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), CONFIG_FILE)
    if os.path.exists(config_path):  # 配置文件存在
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    # 配置文件不存在或解析失败，使用默认配置并保存
    default_cfg = get_default_config()
    save_config(default_cfg)
    return default_cfg


def save_config(config: Dict):
    """
    将配置保存到JSON文件
    
    参数:
        config: Dict - 要保存的配置字典
    
    处理逻辑：
    1. 构建配置文件路径（与脚本同目录）
    2. 以UTF-8编码写入JSON文件
    3. 使用ensure_ascii=False支持中文
    4. 使用indent=4格式化输出
    
    异常处理：
    - 打印错误信息到控制台（不抛出异常）
    """
    config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), CONFIG_FILE)
    try:
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"保存配置失败: {e}")


def get_file_create_time(path: str) -> Optional[str]:
    """
    获取文件的创建/修改时间
    
    参数:
        path: str - 文件路径
    
    返回值:
        Optional[str] - 格式化的时间字符串（YYYY-MM-DD HH:MM:SS），失败返回None
    
    说明：
    - 实际获取的是文件的最后修改时间（mtime）
    - 用于显示CSV文件的创建时间供用户参考
    """
    if not path or not os.path.exists(path):  # 路径无效
        return None
    try:
        stat = os.stat(path)
        mtime = datetime.datetime.fromtimestamp(stat.st_mtime)
        return mtime.strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        return None


# ================================================================================
# 【问题7 - 方案A：配置合并】从column_mapping构建运行时FIELD_MAPPING字典
# ================================================================================
def load_config_and_apply():
    """
    加载配置并应用到运行时变量
    
    返回值:
        Dict - 加载的配置字典
    
    处理逻辑：
    1. 加载配置文件
    2. 从column_mapping构建FIELD_MAPPING字典（全局变量）
       - 格式：{源列名: 目标列名}
    3. 构建目标表头（调用build_target_headers）
    4. 返回配置字典
    
    全局变量:
        FIELD_MAPPING - 运行时字段映射字典
    """
    global FIELD_MAPPING
    config = load_config()
    column_config = config.get('column_config', get_default_config()['column_config'])
    column_mapping = column_config.get('column_mapping', DEFAULT_COLUMN_MAPPING)
    FIELD_MAPPING = {item['source']: item['target'] for item in column_mapping}
    build_target_headers(column_config)
    return config


# ==================== 转换逻辑 ====================

def extract_year_month(title: str) -> str:
    """
    从标题文本提取发薪月份
    
    参数:
        title: str - 源Excel文件的标题文本，通常格式如"XX单位2024年1月工资表"
    
    返回值:
        str - 格式化后的发薪月份，如"202401"（YYYYMM格式），提取失败返回空字符串
    
    正则表达式:
        (\\d{4})年(\\d{1,2})月
        - (\\d{4}): 捕获4位年份
        - (\\d{1,2}): 捕获1-2位月份
    
    示例:
        "XX单位2024年1月工资表" -> "202401"
        "2024年12月工资表" -> "202412"
        "无年份" -> ""
    """
    # 使用正则表达式提取年份和月份
    match = re.search(r'(\d{4})年(\d{1,2})月', title)
    if match:  # 正则匹配成功
        year, month = match.groups()
        # 月份补零：1月 -> "01", 12月 -> "12"
        return f"{year}{month.zfill(2)}"
    return ''


def read_source_file(filepath: str) -> Tuple[Dict[str, Any], List[List[Any]], List[Dict[str, Any]]]:
    """
    读取源Excel文件
    
    参数:
        filepath: str - 源Excel文件路径（.xls或.xlsx）
    
    返回值:
        Tuple[Dict, List, List]:
        - meta: 源文件元数据字典（包含title, year_month, filename）
        - header_rows: 表头行列表（5行）
        - data_rows: 数据行列表
    
    处理逻辑：
    - 根据文件扩展名选择读取方法
    - .xlsx 调用 _read_source_file_xlsx
    - .xls 调用 _read_source_file_xls
    """
    config = load_config()
    if filepath.lower().endswith('.xlsx'):  # XLSX格式
        return _read_source_file_xlsx(filepath, config)
    else:  # XLS格式
        return _read_source_file_xls(filepath, config)


def _read_source_file_xls(filepath: str, config: Dict = None) -> Tuple[Dict[str, Any], List[List[Any]], List[Dict[str, Any]]]:
    """
    读取XLS格式的源Excel文件
    
    参数:
        filepath: str - 源Excel文件路径（.xls格式）
        config: Dict - 配置字典（可选，默认加载配置）
    
    返回值:
        Tuple[Dict, List, List]:
        - meta: 元数据 {title, year_month, filename}
        - header_rows: 表头行列表
        - data_rows: 数据行列表
    
    处理逻辑：
    1. 读取配置确定读取参数（title_row, title_col, header_count, data_start, skip_keyword）
    2. 使用xlrd打开XLS文件
    3. 提取标题（用于获取发薪月份）
    4. 读取表头行（前5行）
    5. 读取数据行：
       - 从data_start开始
       - 遇到skip_keyword（默认"合计"）停止
       - 第一列为数字（工号）才读取
    """
    if config is None:
        config = load_config()
    
    # 从配置获取读取参数
    source_config = config.get('source_config', {})
    title_row = source_config.get('title_row', 0)
    title_col = source_config.get('title_col', 2)
    header_count = source_config.get('header_row_count', 5)
    data_start = source_config.get('data_start_row', 5)
    skip_keyword = source_config.get('skip_keyword', '合计')
    
    # 打开XLS文件
    book = xlrd.open_workbook(filepath)
    sheet = book.sheet_by_index(0)
    
    # 提取标题并计算发薪月份
    title = sheet.cell_value(title_row, title_col) if sheet.nrows > title_row else ''
    year_month = extract_year_month(title)
    
    # 构建元数据
    meta = {
        'title': title,
        'year_month': year_month,
        'filename': os.path.basename(filepath),
    }
    
    # 读取表头行
    header_rows = []
    for row_idx in range(header_count):
        row_data = [sheet.cell_value(row_idx, col_idx) for col_idx in range(sheet.ncols)]
        header_rows.append(row_data)
    
    # 读取数据行
    data_rows = []
    for row_idx in range(data_start, sheet.nrows):
        # 获取第一列值（工号列）
        first_cell = sheet.cell_value(row_idx, 0)
        # 遇到跳过关键字（"合计"行）停止读取
        if skip_keyword and first_cell == skip_keyword:  # 遇到合计行
            break
        # 判断是否为有效数据行：工号为数字
        if isinstance(first_cell, float) or (isinstance(first_cell, str) and first_cell.isdigit()):  # 工号为数字
            row_data = [sheet.cell_value(row_idx, col_idx) for col_idx in range(sheet.ncols)]
            data_rows.append(row_data)
    
    return meta, header_rows, data_rows


def _read_source_file_xlsx(filepath: str, config: Dict = None) -> Tuple[Dict[str, Any], List[List[Any]], List[Dict[str, Any]]]:
    """
    读取XLSX格式的源Excel文件
    
    参数:
        filepath: str - 源Excel文件路径（.xlsx格式）
        config: Dict - 配置字典（可选，默认加载配置）
    
    返回值:
        Tuple[Dict, List, List]:
        - meta: 元数据 {title, year_month, filename}
        - header_rows: 表头行列表
        - data_rows: 数据行列表
    
    处理逻辑：
    1. 读取配置确定读取参数
    2. 使用openpyxl打开XLSX文件（data_only=True获取计算后的值）
    3. 提取标题并计算发薪月份
    4. 读取表头行（前5行）
    5. 读取数据行
    
    注意：
    - openpyxl使用1-based索引（行列从1开始）
    - xlrd使用0-based索引（行列从0开始）
    - 需要做+1转换
    """
    if config is None:  # 未提供配置
        config = load_config()
    
    # 从配置获取读取参数
    source_config = config.get('source_config', {})
    title_row = source_config.get('title_row', 0)
    title_col = source_config.get('title_col', 2)
    header_count = source_config.get('header_row_count', 5)
    data_start = source_config.get('data_start_row', 5)
    skip_keyword = source_config.get('skip_keyword', '合计')
    

    # 打开XLSX文件，data_only=True获取公式计算后的值而非公式
    book = openpyxl.load_workbook(filepath, data_only=True)
    sheet = book.active
    
    # 提取标题（+1转换为openpyxl的1-based索引）
    title = sheet.cell(title_row + 1, title_col + 1).value if sheet.max_row > title_row else ''
    year_month = extract_year_month(str(title) if title else '')
    
    # 构建元数据
    meta = {
        'title': str(title) if title else '',
        'year_month': year_month,
        'filename': os.path.basename(filepath),
    }
    
    # 读取表头行（1到header_count）
    header_rows = []
    for row_idx in range(1, header_count + 1):
        row_data = [sheet.cell(row_idx, col_idx).value for col_idx in range(1, sheet.max_column + 1)]
        header_rows.append(row_data)
    
    # 读取数据行（从data_start+1开始）
    data_rows = []
    for row_idx in range(data_start + 1, sheet.max_row + 1):
        first_cell = sheet.cell(row_idx, 1).value
        # 遇到跳过关键字停止读取
        if skip_keyword and first_cell == skip_keyword:  # 遇到合计行
            break
        # 判断是否为有效数据行
        if isinstance(first_cell, (int, float)) or (isinstance(first_cell, str) and first_cell.isdigit()):  # 工号为数字
            row_data = [sheet.cell(row_idx, col_idx).value for col_idx in range(1, sheet.max_column + 1)]
            data_rows.append(row_data)
    
    book.close()
    return meta, header_rows, data_rows


def extract_source_headers(header_rows: List[List[Any]]) -> List[Tuple[str, int]]:
    """
    从源Excel表头行提取列名
    
    参数:
        header_rows: List[List[Any]] - 源文件的表头行列表（通常为5行）
                     header_rows[0]: 空白行
                     header_rows[1]: 空行
                     header_rows[2]: 主表头（大类，如"基本工资"、"养老保险"）
                     header_rows[3]: 子表头（如"单位"、"个人"）
                     header_rows[4]: 第三级表头（如具体项目名称）
    
    返回值:
        List[Tuple[str, int]] - 列名和列索引的列表，如 [("职工号", 0), ("基本工资_单位", 1), ...]
    
    处理逻辑：
    1. 读取主表头(第3行)、子表头(第4行)、第三表头(第5行)
    2. 遍历每一列，判断是否跨列：
       - 如果主表头值非空，检查向右连续空值的数量确定跨列数
    3. 根据跨列数处理：
       - 跨列数=1：单列，直接使用主表头值
       - 跨列数>1：多列，拼接列名：
         * 格式：主表头_子表头_第三表头
         * 清理主表头多余空格
         * 子表头为空但第三表头不为空时，继承上一列的子表头值
    
    示例：
        主表头: "基本工资" | "" | "" | "养老保险"
        子表头: "单位"     | "个人" | "单位" | "个人"
        第三表头: ""      | ""    | ""    | ""
        结果: [("基本工资_单位", 0), ("基本工资_个人", 1), ("养老保险_单位", 2), ("养老保险_个人", 3)]
    """
    # 获取三层表头
    main_header = header_rows[2]
    sub_header = header_rows[3] if len(header_rows) > 3 else [''] * len(main_header)
    third_header = header_rows[4] if len(header_rows) > 4 else [''] * len(main_header)
    
    headers = []
    col_idx = 0
    
    # 遍历所有列（按主表头跳转，非逐列）
    while col_idx < len(main_header):
        # 获取当前列的主表头值（去空格）
        main_val = str(main_header[col_idx]).strip() if main_header[col_idx] else ''
        
        # ===========================================================
        # 【外层】判断当前列的主表头是否为空
        # 作用：跳过当前位置的空列
        # 
        # continue：跳过当前列（子表头已在上一轮主表头循环中处理成完整列名，避免重复）
        # 
        # 场景示例（子表头已在上一轮主表头循环中被处理）：
        #   主表头: "基本工资" | ""       | ""        | "养老保险"
        #   子表头: "单位"     | "个人"  | "单位"   | "个人"
        #   索引:     0     |  1      |  2       |   3
        #
        # 处理流程：
        #   第1次循环 col_idx=0：主表头="基本工资"，span=3
        #     -> 为列0生成 "基本工资_单位"
        #     -> 为列1生成 "基本工资_个人"  (子表头"个人"在此处理)
        #     -> 为列2生成 "基本工资_单位"
        #   第2次循环 col_idx=1：主表头=""（空）
        #     -> 此时子表头"个人"已在第1次循环中处理成"基本工资_个人"
        #     -> continue 跳过，避免重复生成
        # ===========================================================
        if not main_val:  # 主表头为空 -> 跳过当前列
            col_idx += 1
            continue  # 子表头已在上一轮处理成列名，跳过避免重复
        
        # ===========================================================
        # 【内层】从当前列之后查找下一个有内容的主表头
        # 作用：计算当前主表头跨度（当前列到下一个非空列之间的距离）
        # 两个"空"判断对象不同：
        #   - 外层：判断的是"当前列"的主表头
        #   - 内层：查找的是"当前列之后"的列
        # ===========================================================
        next_main_col = col_idx + 1
        while next_main_col < len(main_header):
            next_main = str(main_header[next_main_col]).strip() if main_header[next_main_col] else ''
            if next_main:  # 找到下一个非空主表头
                break  # 已找到目标，跳出while循环（不是逐列遍历）
            next_main_col += 1
        
        span = next_main_col - col_idx  # 当前主表头覆盖的列数
        
        if span > 1:  # 跨列字段
            # 跨列字段：遍历范围内的每一列
            # 清理主表头的多余空格
            main_val_clean = re.sub(r'\s+', '', main_val)
            prev_sub_val = ''  # 用于继承上一列的 sub_val
            for c in range(col_idx, next_main_col):
                sub_val = str(sub_header[c]).strip() if sub_header[c] else ''
                third_val = str(third_header[c]).strip() if third_header[c] else ''
                
                # 如果 sub_val 为空但 third_val 不为空，继承上一列的 sub_val
                # 判断依据：某些表格中子表头只在第一列有值，后续列需要继承
                if not sub_val and third_val:  # 子表头为空但有第三表头
                    sub_val = prev_sub_val
                
                # 更新 prev_sub_val
                if sub_val:  # 有子表头值
                    prev_sub_val = sub_val
                
                # 拼接列名：主表头_子表头(_第三表头)
                if sub_val:  # 有子表头
                    if third_val:  # 有第三表头
                        col_name = f"{main_val_clean}_{sub_val}_{third_val}"
                    else:
                        col_name = f"{main_val_clean}_{sub_val}"
                else:
                    col_name = main_val_clean
                
                headers.append((col_name, c))
        else:
            # 单列字段：直接使用主表头值
            headers.append((main_val, col_idx))
        
        # 跳转到下一个主表头位置（非逐列+1，而是按主表头跳转）
        # 例如：第1次循环 span=3，col_idx 从0直接跳到3（跳过中间的空列）
        col_idx = next_main_col
    
    return headers


def map_row_to_target(row: List[Any], source_headers: List[Tuple[str, int]]) -> List[Any]:
    """
    将源数据行映射到目标格式
    
    参数:
        row: List[Any] - 源Excel的一行数据
        source_headers: List[Tuple[str, int]] - 源列名和索引列表
    
    返回值:
        List[Any] - 目标格式的一行数据
    
    处理逻辑：
    1. 构建源数据字典 {列名: 值}
    2. 遍历FIELD_MAPPING，将源列值映射到目标列位置
    3. 返回29列的目标行
    
    说明：
    - 此函数现已较少使用，convert_single_file中直接处理
    """
    source_data = {}
    for col_name, col_idx in source_headers:
        if col_idx < len(row):  # 列索引在数据范围内
            source_data[col_name] = row[col_idx]
    
    target_row = [''] * 29
    
    for source_name, target_name in FIELD_MAPPING.items():
        if source_name in source_data:  # 源列有数据
            target_idx = TARGET_FIELD_INDEX.get(target_name)
            if target_idx is not None:  # 目标列存在
                target_row[target_idx] = source_data[source_name]
    
    return target_row


def clean_excel_formula(value: str) -> str:
    """
    清理Excel公式值
    
    参数:
        value: str - 可能包含Excel公式的值
    
    返回值:
        str - 清理后的值
    
    处理逻辑：
    1. 空值返回空字符串
    2. 如果值以='开头并以"结尾（Excel公式格式），提取内部内容
    3. 否则去除首尾空格
    
    示例：
        =\"张三\" -> 张三
        " 张三 " -> 张三
    """
    if not value:  # 空值
        return ''
    if value.startswith('="') and value.endswith('"'):  # Excel公式格式
        return value[2:-1]
    return value.strip()


def load_phone_lookup(csv_path: str, id_col_name: str = None, phone_col_name: str = None) -> Dict[str, str]:
    """
    从人员信息文件加载电话查询字典
    
    参数:
        csv_path: str - CSV/XLSX文件路径
        id_col_name: str - 身份证列名（可选，从配置读取）
        phone_col_name: str - 手机号列名（可选，从配置读取）
    
    返回值:
        Dict[str, str] - 身份证号到手机号的映射 {证件号码: 手机号}
    
    处理逻辑：
    1. 检查文件是否存在
    2. 根据文件格式选择读取方法：
       - .xlsx: 调用load_phone_lookup_from_xlsx
       - .csv: 直接读取CSV
    3. 从配置获取列名映射
    4. 构建lookup字典
    
    编码：
    - CSV使用GB18030编码（支持中文）
    """
    global FIELD_MAPPING
    lookup = {}
    if not csv_path or not os.path.exists(csv_path):  # 文件不存在
        return lookup
    config = load_config()
    phone_config = config.get('phone_lookup_config', {})
    
    # XLSX格式：从配置获取列名，调用专用函数
    if csv_path.endswith('.xlsx'):  # XLSX格式
        xlsx_config = phone_config.get('xlsx', {})
        if id_col_name is None:  # 未指定身份证列
            id_col_name = xlsx_config.get('id_card_column', '*证件号码')
        if phone_col_name is None:  # 未指定手机号列
            phone_col_name = xlsx_config.get('phone_column', '*手机号')
        sheet_name = xlsx_config.get('sheet_name', '花名册')
        return load_phone_lookup_from_xlsx(csv_path, id_col_name, phone_col_name, sheet_name)
    
    # CSV格式：直接从配置获取列名
    csv_config = phone_config.get('csv', {})
    if id_col_name is None:  # 未指定身份证列
        id_col_name = csv_config.get('id_card_column', '公民身份号码')
    if phone_col_name is None:  # 未指定手机号列
        phone_col_name = csv_config.get('phone_column', '联系电话')
    
    # 读取CSV文件
    try:
        with open(csv_path, 'r', encoding='gb18030') as f:
            reader = csv.reader(f)
            header = next(reader)
            # 查找身份证列和手机号列的索引
            # enumerate: 带索引遍历表头行
            id_col = -1
            phone_col = -1
            for i, col_name in enumerate(header):
                if col_name == id_col_name:  # 找到身份证列
                    id_col = i
                elif col_name == phone_col_name:  # 找到手机号列
                    phone_col = i
                if id_col >= 0 and phone_col >= 0:  # 两列都找到
                    break
            if id_col < 0 or phone_col < 0:  # 列未找到
                return lookup
            # 构建lookup字典
            for row in reader:
                if len(row) > max(id_col, phone_col):  # 行数据足够长
                    id_card = clean_excel_formula(row[id_col])
                    phone = clean_excel_formula(row[phone_col])
                    if id_card and phone:  # 证件号和手机号都有值
                        lookup[id_card] = phone
    except Exception:
        pass
    return lookup


def load_phone_lookup_from_xlsx(xlsx_path: str, id_col_name: str, 
                                 phone_col_name: str, sheet_name: str) -> Dict[str, str]:
    """从XLSX文件读取电话查询表"""
    lookup = {}
    if not xlsx_path or not os.path.exists(xlsx_path):  # 文件不存在
        return lookup
    
    try:
        wb = openpyxl.load_workbook(xlsx_path)
        ws = wb[sheet_name]
        
        # 读取表头行
        header = [cell.value for cell in ws[1]]
        id_col = header.index(id_col_name) if id_col_name in header else -1
        phone_col = header.index(phone_col_name) if phone_col_name in header else -1
        
        if id_col < 0 or phone_col < 0:
            return lookup
        
        # 遍历数据行（从第2行开始）
        for row in ws.iter_rows(min_row=2, values_only=True):
            if len(row) > max(id_col, phone_col):
                id_card = str(row[id_col]).strip() if row[id_col] else ''
                phone = str(row[phone_col]).strip() if row[phone_col] else ''
                if id_card and phone:
                    lookup[id_card] = phone
    except Exception:
        pass
    return lookup


def load_new_employees_from_template(xlsx_path: str) -> Dict[str, Dict]:
    """
    从员工导入模板读取新增人员信息
    
    参数:
        xlsx_path: str - 员工导入模板文件路径
    
    返回值:
        Dict[str, Dict] - 新员工信息字典
        格式：{证件号码: {name: 姓名, phone: 手机号}}
    
    处理逻辑：
    1. 优先从"新增人员"Sheet读取，如果没有则从"花名册"Sheet读取
    2. 自动识别姓名、手机号、证件号码列（模糊匹配）
    3. 遍历数据行，构建lookup字典
    4. 相同证件号码只保留第一条记录
    
    列名识别规则：
    - 姓名列：包含"姓名"的列
    - 手机号列：包含"手机"的列
    - 证件号码列：精确匹配"证件号码"或包含"证件"的列
    """
    lookup = {}
    if not xlsx_path or not os.path.exists(xlsx_path):
        return lookup
    
    try:
        wb = openpyxl.load_workbook(xlsx_path)
        
        # 选择Sheet：优先"新增人员"，其次"花名册"
        sheet_name = None
        if '新增人员' in wb.sheetnames:
            sheet_name = '新增人员'
        elif '花名册' in wb.sheetnames:
            sheet_name = '花名册'
        
        if not sheet_name:
            return lookup
        
        ws = wb[sheet_name]
        header = [cell.value for cell in ws[1]]
        
        # 自动识别列索引（模糊匹配）
        # next() + generator expression: next((expr for item in iterable if condition), default)
        # 找第一个匹配的索引，找不到返回默认值
        name_col = next((i for i, h in enumerate(header) if h and '姓名' in str(h)), 1)
        phone_col = next((i for i, h in enumerate(header) if h and '手机' in str(h)), 2)
        id_col = next((i for i, h in enumerate(header) if '证件号码' in str(h) or h == '证件号码'), None)
        if id_col is None:
            id_col = next((i for i, h in enumerate(header) if h and '证件' in str(h)), 3)
        
        # 遍历数据行
        for row in ws.iter_rows(min_row=2, values_only=True):
            if len(row) > max(name_col, phone_col, id_col):
                # ternary expression: value if condition else default
                name = str(row[name_col]).strip() if row[name_col] else ''
                phone = str(row[phone_col]).strip() if row[phone_col] else ''
                id_card = str(row[id_col]).strip() if row[id_col] else ''
                
                # 相同证件号码只保留第一条
                if id_card and id_card not in lookup:
                    lookup[id_card] = {'name': name, 'phone': phone}
    except Exception:
        pass
    
    return lookup


def has_value(value: Any) -> bool:
    """
    判断值是否有实际数据
    
    参数:
        value: Any - 任意类型的值
    
    返回值:
        bool - 是否有实际数据
    
    判断规则：
    - None -> False
    - 数字(int/float) -> True
    - 字符串 -> 去除空格后非空为True
    - 其他类型 -> True
    
    用途：
    - 用于判断源表的某列是否有数据（决定是否输出动态列）
    """
    if value is None:
        return False
    if isinstance(value, (int, float)):
        return True
    if isinstance(value, str):
        return value.strip() != ''
    return True


# ================================================================================
# 【问题8修复】检测动态列和未知列
# 只将target在fixed_columns中的列视为"固定列"，其他在column_mapping中的列为"动态列"
# 不在column_mapping中且有数据的列为"未知列"
# ================================================================================
def detect_dynamic_and_unknown_columns(source_headers: List[Tuple[str, int]], 
                                        data_rows: List[List[Any]],
                                        column_config: Dict) -> Tuple[List[Tuple[str, str]], List[Tuple[str, str]]]:
    # 从column_mapping构建source到target的映射字典
    column_mapping = column_config.get('column_mapping', DEFAULT_COLUMN_MAPPING)
    source_to_target = {item['source']: item['target'] for item in column_mapping}
    known_sources = set(source_to_target.keys())
    
    # ================================================================================
    # 【问题8修复关键】只将target在fixed_columns中的列视为固定列
    # 原错误逻辑：将所有column_mapping中的source和target都加入fixed_source_columns
    # 修复后：只将target在fixed_columns中的映射加入fixed_source_columns
    # ================================================================================
    fixed_target_set = set(column_config.get('fixed_columns', []))
    fixed_source_columns = set()
    for src, tgt in source_to_target.items():
        if tgt in fixed_target_set:  # 只添加目标在fixed_columns中的映射
            fixed_source_columns.add(src)
            fixed_source_columns.add(tgt)
    
    # dynamic_cols_with_data: 有数据的动态列列表 [源列名, 目标列名]
    dynamic_cols_with_data = []
    # unknown_cols_with_data: 有数据的未知列列表
    unknown_cols_with_data = []
    
    # source_col_has_data: 标记源列是否有数据（用于判断是否输出）
    source_col_has_data = {}
    for col_name, col_idx in source_headers:
        source_col_has_data[col_name] = False
        for data_row in data_rows[:10]:
            if col_idx < len(data_row) and has_value(data_row[col_idx]):
                source_col_has_data[col_name] = True
                break
    
    for col_name, col_idx in source_headers:
        if col_name in fixed_source_columns:
            continue
        
        if col_name in known_sources:
            if source_col_has_data.get(col_name, False):
                target = source_to_target.get(col_name, col_name)
                dynamic_cols_with_data.append((col_name, target))
        else:
            if source_col_has_data.get(col_name, False):
                unknown_cols_with_data.append((col_name, col_name))
    
        return dynamic_cols_with_data, unknown_cols_with_data


# ================================================================================
# 核心转换函数：convert_single_file
# ================================================================================
def convert_single_file(input_path: str, output_dir: str, phone_lookup: Dict[str, str] = None, 
                       new_employees: Dict[str, Dict] = None) -> Tuple[bool, str, Dict]:
    """
    转换单个Excel文件
    
    参数说明:
        - input_path: 源Excel文件路径（.xls或.xlsx）
        - output_dir: 输出目录路径
        - phone_lookup: 从start_convert传入的电话查询字典 {证件号码: 手机号}
        - new_employees: 从start_convert传入的新员工信息字典 {证件号码: {name, phone}}
    
    处理流程:
        1. 读取源文件：read_source_file() -> meta, header_rows, data_rows
        2. 提取列名：extract_source_headers(header_rows) -> source_headers
        3. 检测动态列和未知列：detect_dynamic_and_unknown_columns()
        4. 生成输出表头：final_headers_1(大类) + final_headers_2(列名)
        5. 逐行转换数据：
           a. 通过FIELD_MAPPING写入固定列数据
           b. 遍历源表列，写入动态列数据（column_mapping中且不在fixed_columns中）
           c. 根据output_unknown开关决定是否写入未知列
        6. 补充手机号（从phone_lookup查询）
        7. 补充发薪月份（从标题提取）
        8. 写入输出Excel文件
    
    返回值:
        Tuple[bool, str, Dict]:
            - 成功时: (True, "成功转换 X 行数据", {total_rows, warnings, output_path})
            - 失败时: (False, "转换失败: 错误信息", {error})
    """
    # phone_lookup: 从函数参数传入（start_convert中加载的CSV文件数据）
    if phone_lookup is None:
        phone_lookup = {}
    
    # column_config: 从配置文件加载（包含fixed_columns、column_mapping、output_unknown_columns）
    config = load_config()
    column_config = config.get('column_config', get_default_config()['column_config'])
    
    filename = os.path.basename(input_path)
    name_without_ext = os.path.splitext(filename)[0]
    
    counter = 1
    while True:
        if counter == 1:  # 第一个文件
            output_filename = f"{name_without_ext}_converted.xlsx"
        else:  # 后续文件
            output_filename = f"{name_without_ext}_converted_{counter}.xlsx"
        
        output_path = os.path.join(output_dir, output_filename)
        
        if not os.path.exists(output_path):  # 文件不存在
            break
        counter += 1
    
    warnings = []
    try:
        meta, header_rows, data_rows = read_source_file(input_path)
        source_headers = extract_source_headers(header_rows)
        
        # dynamic_cols: detect_dynamic_and_unknown_columns返回的动态列列表（有数据的非固定列）
        # unknown_cols: detect_dynamic_and_unknown_columns返回的未知列列表（不在配置中的列）
        dynamic_cols, unknown_cols = detect_dynamic_and_unknown_columns(
            source_headers, data_rows, column_config
        )
        
        # ================================================================================
        # 【问题1修复】定义目标列名对应的大类（用于第1行表头）
        # 固定列：工号/姓名/证件号码/手机号/发薪月份 -> 保持原名
        # 总计：基本工资/补发工资/应发工资/实发金额 -> 总计
        # 扣缴基数：社保基数/医保基数/工伤基数/公积金基数 -> 扣缴基数
        # 应扣项：各类保险和公积金扣款 -> 应扣项
        # 其它：不在此映射中的列 -> 其它
        # ================================================================================
        fixed_category = {
    '工号': '工号', '姓名': '姓名', '证件号码': '证件号码', '手机号': '手机号', '发薪月份': '发薪月份',
    '基本工资': '总计', '补发工资': '总计', '应发工资': '总计', '实发金额': '总计',
    '社保基数': '扣缴基数', '医保基数': '扣缴基数', '工伤基数': '扣缴基数', '公积金基数': '扣缴基数',
    '扣单位养老保险': '应扣项', '扣个人养老保险': '应扣项', '扣单位失业保险': '应扣项', '扣个人失业保险': '应扣项',
    '扣单位医疗保险': '应扣项', '扣个人医疗保险': '应扣项', '扣单位工伤险': '应扣项', '扣单位住房公积金': '应扣项',
    '扣个人住房公积金': '应扣项', '扣缴税金（扣个人所得税）': '应扣项', '2025年个税补退': '应扣项'
}

        # ================================================================================
        # 【问题1修复】生成输出表头
        # final_headers_1: 第1行大类名称（如"总计"、"应扣项"）
        # final_headers_2: 第2行列名（如"基本工资"、"扣单位养老保险"）
        # 顺序：固定列 -> 动态列（有数据时） -> 未知列（开关开启时）
        # ================================================================================
        fixed_columns = column_config['fixed_columns']
        
        final_headers_1 = [fixed_category.get(col, col) for col in fixed_columns]
        final_headers_2 = fixed_columns.copy()
        
        for _, target in dynamic_cols:
            final_headers_2.append(target)
            final_headers_1.append(fixed_category.get(target, '其它'))
        
        # ================================================================================
        # 【问题3】输出未知列开关
        # ================================================================================
        output_unknown = column_config.get('output_unknown_columns', False)  # dict.get(键, 默认值)：安全取值
        if output_unknown:  # 开关开启时追加未知列
            for _, target in unknown_cols:
                final_headers_2.append(target)
                final_headers_1.append('其它')
        
        # dict comprehension: {name: idx for idx, name in enumerate(final_headers_2)} 构建列名→索引映射
        field_index = {name: idx for idx, name in enumerate(final_headers_2)}
        
        workbook = xlsxwriter.Workbook(output_path)
        worksheet = workbook.add_worksheet('Sheet1')
        
        worksheet.set_column('A:A', 12)
        worksheet.set_column('B:B', 10)
        worksheet.set_column('C:C', 22)
        worksheet.set_column('D:D', 12)
        worksheet.set_column('E:AC', 10)
        
        # enumerate: 带索引遍历，可同时获取索引和值
        for col_idx, header in enumerate(final_headers_1):
            worksheet.write(0, col_idx, header)
        
        for col_idx, header in enumerate(final_headers_2):
            worksheet.write(1, col_idx, header)
        
        # id_card_col_idx: 从source_headers中查找的身份证列索引
        # name_col_idx: 从source_headers中查找的姓名列索引
        id_card_col_idx = None
        name_col_idx = None
        for col_name, col_idx in source_headers:
            if col_name in ('身份证', '证件号码'):
                id_card_col_idx = col_idx
            if col_name == '姓名':
                name_col_idx = col_idx
        
        # current_id_cards: 用于收集数据行中的证件号，后面与new_employees匹配
        current_id_cards = []
        
        # ================================================================================
        # 数据写入准备：从column_mapping构建source_to_target映射
        # ================================================================================
        # fixed_col_set: 从fixed_columns列表构建的集合（用于判断是否跳过）
        fixed_col_set = set(column_config.get('fixed_columns', []))
        # column_mapping: 从配置加载的字段映射列表
        column_mapping = column_config.get('column_mapping', DEFAULT_COLUMN_MAPPING)  # dict.get(键, 默认值)
        # dict comprehension: {item['source']: item['target'] for item in column_mapping} 构建 {源列名: 目标列名} 映射
        source_to_target = {item['source']: item['target'] for item in column_mapping}
        # output_unknown: 从配置获取的是否输出未知列开关
        output_unknown = column_config.get('output_unknown_columns', False)
        # set comprehension + ternary: {src for src, tgt in unknown_cols} if output_unknown else set()
        # 如果开关开启，从unknown_cols提取源列名构建集合，否则空集合
        unknown_col_names = {src for src, tgt in unknown_cols} if output_unknown else set()
        
        # ================================================================================
        # 【问题2修复】数据写入逻辑 - 按顺序处理各类列
        # 1. 先通过FIELD_MAPPING写入固定列数据（第787-792行）
        # 2. 再遍历原始表列，处理动态列和未知列（第794行开始）
        # 执行顺序：跳过固定列 -> 处理动态列 -> 处理未知列
        # ================================================================================
        # enumerate: 带索引遍历数据行
        for row_idx, data_row in enumerate(data_rows):
            target_row = [''] * len(final_headers_2)  # list: 初始化空列表
            
            # 写入固定列数据（通过FIELD_MAPPING）
            # dict.items(): 遍历键值对
            for source_name, target_name in FIELD_MAPPING.items():
                for col_name, col_idx in source_headers:
                    if col_name == source_name and col_idx < len(data_row):
                        target_idx = field_index.get(target_name)  # dict.get(): 安全取值
                        if target_idx is not None:
                            target_row[target_idx] = data_row[col_idx]
            
            # 写入动态列和未知列数据
            for col_name, col_idx in source_headers:
                if col_idx >= len(data_row):
                    continue
                value = data_row[col_idx]
                if not has_value(value):
                    continue
                
                # 1. 跳过固定列
                if col_name in fixed_col_set:
                    continue
                
                # 2. 处理动态列（在column_mapping中且不在fixed_columns中）
                if col_name in source_to_target:
                    target_name = source_to_target[col_name]
                    if target_name in field_index:
                        target_row[field_index[target_name]] = value
                elif col_name in unknown_col_names:
                    if col_name in final_headers_2:
                        idx = final_headers_2.index(col_name)
                        target_row[idx] = value
            
            id_card = ''
            if id_card_col_idx is not None and id_card_col_idx < len(data_row):
                id_card = str(data_row[id_card_col_idx]).strip()
            
            name = ''
            if name_col_idx is not None and name_col_idx < len(data_row):
                name = str(data_row[name_col_idx]).strip() if data_row[name_col_idx] else ''
            
            id_card = id_card.strip() if id_card else ''
            
            if id_card:
                current_id_cards.append(id_card)
            
            phone = phone_lookup.get(id_card, '')
            
            if '手机号' in field_index:
                target_row[field_index['手机号']] = phone
            if '发薪月份' in field_index:
                target_row[field_index['发薪月份']] = meta['year_month']
            
            row_missing = []
            if not name:
                row_missing.append('姓名')
            if not id_card:
                row_missing.append('身份证号')
            if not phone:
                row_missing.append('手机号')
            
            if row_missing:
                warnings.append({
                    'row': row_idx + 1,
                    'missing_fields': row_missing,
                    'name': name if name else f'(第{row_idx + 1}行)'
                })
            
            for col_idx, value in enumerate(target_row):
                if value is not None and value != '':
                    worksheet.write(row_idx + 2, col_idx, value)
        
        workbook.close()
        
        if new_employees and current_id_cards:
            # set operations: set1 & set2 求交集
            # set(collection): 从可迭代对象创建集合（去重）
            intersection = set(current_id_cards) & set(new_employees.keys())
            
            if intersection:  # 有交集则写入新员工Sheet
                wb_out = openpyxl.load_workbook(output_path)
                
                if '花名册' in wb_out.sheetnames:
                    del wb_out['花名册']
                ws_new = wb_out.create_sheet('花名册')
                
                headers = ['*姓名', '*手机号', '*证件类型', '*证件号码']
                for col, header in enumerate(headers, start=1):
                    cell = ws_new.cell(row=1, column=col, value=header)
                    cell.font = Font(bold=True)
                    cell.alignment = Alignment(horizontal='center')
                
                row_num = 2
                for id_card in intersection:
                    emp = new_employees[id_card]
                    ws_new.cell(row=row_num, column=1, value=emp['name'])
                    ws_new.cell(row=row_num, column=2, value=emp['phone'])
                    ws_new.cell(row=row_num, column=3, value='居民身份证')
                    ws_new.cell(row=row_num, column=4, value=id_card)
                    
                    for col in range(1, 5):
                        ws_new.cell(row=row_num, column=col).alignment = Alignment(horizontal='center')
                    row_num += 1
                
                ws_new.column_dimensions['A'].width = 25
                ws_new.column_dimensions['B'].width = 13
                ws_new.column_dimensions['C'].width = 13
                ws_new.column_dimensions['D'].width = 13
                ws_new.row_dimensions[1].height = 15
                
                wb_out.save(output_path)
                detail = {
                    'total_rows': len(data_rows),
                    'warnings': warnings,
                    'output_path': output_path,
                    'new_employee_count': len(intersection),
                    'dynamic_columns': len(dynamic_cols),
                    'unknown_columns': len(unknown_cols)
                }
                return True, f"成功转换 {len(data_rows)} 行数据，新增 {len(intersection)} 名人员", detail
        
        detail = {
            'total_rows': len(data_rows),
            'warnings': warnings,
            'output_path': output_path,
            'dynamic_columns': len(dynamic_cols),
            'unknown_columns': len(unknown_cols)
        }
        return True, f"成功转换 {len(data_rows)} 行数据", detail
    
    except Exception as e:
        detail = {'error': str(e)}
        return False, f"转换失败: {str(e)}", detail


def get_file_info(filepath: str) -> Optional[Dict[str, Any]]:
    """
    获取文件信息（用于文件列表显示）
    
    参数:
        filepath: str - Excel文件路径
    
    返回值:
        Optional[Dict] - 文件信息字典，失败返回None
        包含键值：
        - path: 文件完整路径
        - filename: 文件名
        - title: 标题文本
        - year_month: 发薪月份（YYYYMM格式）
        - data_rows: 数据行数
        - file_size: 文件大小（字节）
    
    异常处理：
    - 捕获所有异常，失败返回None
    """
    try:
        meta, _, data_rows = read_source_file(filepath)
        file_size = os.path.getsize(filepath)
        return {
            'path': filepath,
            'filename': meta['filename'],
            'title': meta['title'],
            'year_month': meta['year_month'],
            'data_rows': len(data_rows),
            'file_size': file_size,
        }
    except Exception as e:
        return None


# ==================== GUI 部分 ====================

# ================================================================================
# 配置编辑器 - 支持编辑固定列、字段映射、源表配置、电话查询配置
# 【问题7 - 方案A】界面重构：删除旧的"字段映射"和"动态列"tab
# 新建统一的"字段映射"tab，加载column_mapping数据
# ================================================================================
class ConfigEditor(tk.Toplevel):
    """
    配置编辑器窗口 - 用于编辑应用程序配置项
    
    属性:
        parent: tk.Tk - 父窗口
        config: Dict - 初始配置字典
        result: Dict - 确定后的配置字典（用户点击确定时设置）
        fixed_columns: List[str] - 固定列名列表
        column_mapping: List[Dict] - 字段映射列表 [{source, target}, ...]
        output_unknown_columns: bool - 是否输出未知列开关
    """
    
    def __init__(self, parent, config):
        """
        初始化配置编辑器
        
        参数:
            parent: tk.Tk - 父窗口对象
            config: Dict - 应用程序配置字典（将复制一份用于编辑）
        """
        super().__init__(parent)
        self.title("配置编辑器")
        # ================================================================================
        # 窗口高度从550调整为600，解决确认按钮显示不全的问题
        # ================================================================================
        self.geometry("800x600")
        self.resizable(True, True)
        self.transient(parent)
        self.grab_set()
        
        self.config = config.copy()
        self.result = None
        self.mapping_entries = []
        
        # ================================================================================
        # 【问题7 - 方案A】加载column_mapping用于字段映射编辑
        # ================================================================================
        column_config = self.config.get('column_config', get_default_config()['column_config'])
        self.fixed_columns = column_config.get('fixed_columns', []).copy()
        self.column_mapping = column_config.get('column_mapping', DEFAULT_COLUMN_MAPPING).copy()
        self.output_unknown_columns = column_config.get('output_unknown_columns', False)
        
        self.setup_ui()
    
    def setup_ui(self):
        """
        设置配置编辑器UI布局
        
        创建Notebook组件，包含4个选项卡：
        - 电话查询配置：CSV/XLSX格式的身份证和手机号列名配置
        - 源表配置：标题行、表头行数、数据起始行等配置
        - 固定列：固定输出列列表和未知列开关
        - 字段映射：源列名到目标列名的映射关系
        
        底部添加"重置为默认"、"取消"、"确定"三个按钮
        """
        notebook = ttk.Notebook(self)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 创建4个选项卡Frame
        self.tab_phone = ttk.Frame(notebook)
        self.tab_source = ttk.Frame(notebook)
        self.tab_fixed = ttk.Frame(notebook)
        self.tab_mapping = ttk.Frame(notebook)
        
        # 将选项卡添加到Notebook
        notebook.add(self.tab_phone, text="电话查询配置")
        notebook.add(self.tab_source, text="源表配置")
        notebook.add(self.tab_fixed, text="固定列")
        notebook.add(self.tab_mapping, text="字段映射")
        
        # 设置每个选项卡的内容
        self.setup_phone_tab()
        self.setup_source_tab()
        self.setup_fixed_tab()
        self.setup_mapping_tab()
        
        # 底部按钮栏
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(btn_frame, text="重置为默认", command=self.reset_to_default).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=self.cancel).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_frame, text="确定", command=self.confirm).pack(side=tk.RIGHT, padx=5)
    
    # ================================================================================
    # 【问题7 - 方案A】字段映射tab - Treeview样式
    # 显示column_mapping数据，动态统计显示固定列和动态列数量
    # 支持添加/编辑/删除/上移/下移操作
    # ================================================================================
    def setup_mapping_tab(self):
        """
        设置"字段映射"选项卡内容
        
        使用Treeview组件显示column_mapping数据，支持以下操作：
        - 添加：新增映射关系
        - 编辑：修改选中映射
        - 上移/下移：调整顺序
        - 删除：移除选中映射
        
        动态统计显示：
        - 总映射数 = 固定列映射 + 动态列映射
        - 固定列：target在fixed_columns中的映射，始终输出
        - 动态列：target不在fixed_columns中的映射，有数据时才输出
        """
        frame = ttk.Frame(self.tab_mapping, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        
        # ================================================================================
        # 动态统计显示：字段映射（70项，固定列24项+动态列46项；动态列为非固定列配置项，有数据时输出）
        # 计算公式：固定列数量 = len(fixed_columns)，动态列数量 = 总映射 - 固定列数量
        # 判断依据：column_mapping中target在fixed_columns中为固定列，否则为动态列
        # ================================================================================
        total_mapping = len(self.column_mapping)
        fixed_count = len(self.fixed_columns)
        dynamic_count = total_mapping - fixed_count
        ttk.Label(frame, text=f"字段映射（{total_mapping}项，固定列{fixed_count}项+动态列{dynamic_count}项；动态列为非固定列配置项，有数据时输出）").pack(pady=(0, 5))
        
        tree_frame = ttk.Frame(frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        columns = ('source', 'target')
        self.mapping_tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=18)
        self.mapping_tree.heading('source', text='源列名（原始表）')
        self.mapping_tree.heading('target', text='目标列名（输出表）')
        self.mapping_tree.column('source', width=250)
        self.mapping_tree.column('target', width=250)
        self.mapping_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.mapping_tree.yview)
        self.mapping_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, pady=(5, 0))
        ttk.Button(btn_frame, text="添加", command=self.mapping_add).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="编辑", command=self.mapping_edit).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="上移", command=self.mapping_move_up).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="下移", command=self.mapping_move_down).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="删除", command=self.mapping_delete).pack(side=tk.LEFT, padx=2)
        
        self.refresh_mapping_list()
    
    def refresh_mapping_list(self):
        """
        刷新字段映射列表显示
        
        处理逻辑：
        1. 清空Treeview所有项
        2. 重新插入column_mapping中的所有项
        """
        for item in self.mapping_tree.get_children():
            self.mapping_tree.delete(item)
        for item in self.column_mapping:
            self.mapping_tree.insert('', tk.END, values=(item['source'], item['target']))
    
    # 注意：mapping_add, mapping_edit, mapping_move_up, mapping_move_down, mapping_delete
    # 在之前已添加注释
    def mapping_add(self):
        dialog = tk.Toplevel(self)
        dialog.title("添加字段映射")
        dialog.geometry("400x150")
        dialog.transient(self)
        dialog.grab_set()
        
        ttk.Label(dialog, text="源列名:").grid(row=0, column=0, sticky=tk.W, padx=10, pady=10)
        src_entry = ttk.Entry(dialog, width=30)
        src_entry.grid(row=0, column=1, sticky=tk.W, padx=10)
        
        ttk.Label(dialog, text="目标列名:").grid(row=1, column=0, sticky=tk.W, padx=10, pady=10)
        tgt_entry = ttk.Entry(dialog, width=30)
        tgt_entry.grid(row=1, column=1, sticky=tk.W, padx=10)
        
        def save():
            src = src_entry.get().strip()
            tgt = tgt_entry.get().strip()
            if src and tgt:
                self.column_mapping.append({'source': src, 'target': tgt})
                self.refresh_mapping_list()
            dialog.destroy()
        
        btn_frame = ttk.Frame(dialog)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=10)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="确定", command=save).pack(side=tk.LEFT, padx=5)
    
    def mapping_edit(self):
        """
        编辑选中字段映射
        
        弹出对话框修改源列名和目标列名：
        - 预填当前值
        - 保存后更新column_mapping并刷新列表
        """
        selected = self.mapping_tree.selection()
        if not selected:
            return
        item = selected[0]
        values = self.mapping_tree.item(item, 'values')
        
        dialog = tk.Toplevel(self)
        dialog.title("编辑字段映射")
        dialog.geometry("400x150")
        dialog.transient(self)
        dialog.grab_set()
        
        ttk.Label(dialog, text="源列名:").grid(row=0, column=0, sticky=tk.W, padx=10, pady=10)
        src_entry = ttk.Entry(dialog, width=30)
        src_entry.insert(0, values[0])
        src_entry.grid(row=0, column=1, sticky=tk.W, padx=10)
        
        ttk.Label(dialog, text="目标列名:").grid(row=1, column=0, sticky=tk.W, padx=10, pady=10)
        tgt_entry = ttk.Entry(dialog, width=30)
        tgt_entry.insert(0, values[1])
        tgt_entry.grid(row=1, column=1, sticky=tk.W, padx=10)
        
        def save():
            src = src_entry.get().strip()
            tgt = tgt_entry.get().strip()
            if src and tgt:
                idx = self.column_mapping.index({'source': values[0], 'target': values[1]})
                self.column_mapping[idx] = {'source': src, 'target': tgt}
                self.refresh_mapping_list()
            dialog.destroy()
        
        btn_frame = ttk.Frame(dialog)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=10)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="确定", command=save).pack(side=tk.LEFT, padx=5)
    
    def mapping_move_up(self):
        """
        将选中映射上移一位
        
        交换选中项与上一项的位置
        """
        selected = self.mapping_tree.selection()
        if selected:
            item = selected[0]
            values = self.mapping_tree.item(item, 'values')
            idx = self.column_mapping.index({'source': values[0], 'target': values[1]})
            if idx > 0:
                self.column_mapping[idx-1], self.column_mapping[idx] = self.column_mapping[idx], self.column_mapping[idx-1]
                self.refresh_mapping_list()
                self.mapping_tree.selection_set(self.mapping_tree.get_children()[idx-1])
    
    def mapping_move_down(self):
        """
        将选中映射下移一位
        
        交换选中项与下一项的位置
        """
        selected = self.mapping_tree.selection()
        if selected:
            item = selected[0]
            values = self.mapping_tree.item(item, 'values')
            idx = self.column_mapping.index({'source': values[0], 'target': values[1]})
            if idx < len(self.column_mapping) - 1:
                self.column_mapping[idx], self.column_mapping[idx+1] = self.column_mapping[idx+1], self.column_mapping[idx]
                self.refresh_mapping_list()
                self.mapping_tree.selection_set(self.mapping_tree.get_children()[idx+1])
    
    def mapping_delete(self):
        """
        删除选中映射
        
        从column_mapping中移除选中项
        """
        selected = self.mapping_tree.selection()
        if selected:
            item = selected[0]
            values = self.mapping_tree.item(item, 'values')
            self.column_mapping.remove({'source': values[0], 'target': values[1]})
            self.refresh_mapping_list()
    
    def _on_frame_configure(self, event=None):
        """
        Canvas框架配置回调
        
        用于更新Canvas滚动区域
        """
        self.canvas.configure(scrollregion=self.canvas.bbox('all'))
    
    def _on_canvas_configure(self, event):
        """
        Canvas尺寸配置回调
        
        用于响应窗口大小变化
        """
        self.canvas.itemconfig(self.mapping_frame_id, width=event.width)
        src_entry.pack(side=tk.LEFT, padx=5)
        src_entry.insert(0, src)
        
        tgt_entry = ttk.Entry(row_frame, width=30)
        tgt_entry.pack(side=tk.LEFT, padx=5)
        tgt_entry.insert(0, tgt)
        
        var = tk.BooleanVar(value=False)
        chk = ttk.Checkbutton(row_frame, variable=var)
        chk.pack(side=tk.LEFT, padx=5)
        
        self.mapping_entries.append((var, src_entry, tgt_entry))
    
    def add_mapping(self):
        self.add_mapping_row()
    
    def delete_mapping(self):
        to_delete = []
        for i, (var, src_entry, tgt_entry) in enumerate(self.mapping_entries):
            if var.get():
                to_delete.append(i)
        
        for i in reversed(to_delete):
            var, src_entry, tgt_entry = self.mapping_entries[i]
            src_entry.master.destroy()
            del self.mapping_entries[i]
    
    def setup_phone_tab(self):
        """
        设置"电话查询配置"选项卡
        
        UI组件：
        - CSV格式配置：身份证列名、联系电话列名
        - XLSX格式配置：身份证列名、手机号列名、Sheet名称
        
        数据来源：
        - CSV默认：公民身份号码、联系电话
        - XLSX默认：*证件号码、*手机号、花名册
        
        控件：
        - csv_id_card_entry: CSV身份证列名输入框
        - csv_phone_entry: CSV联系电话列名输入框
        - xlsx_id_card_entry: XLSX身份证列名输入框
        - xlsx_phone_entry: XLSX手机号列名输入框
        - xlsx_sheet_name_entry: XLSX Sheet名称输入框
        """
        frame = ttk.Frame(self.tab_phone, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)
        
        phone_config = self.config.get('phone_lookup_config', {})
        csv_config = phone_config.get('csv', {})
        xlsx_config = phone_config.get('xlsx', {})
        
        # CSV格式配置区域
        ttk.Label(frame, text="CSV 格式配置", font=('Microsoft YaHei', 10, 'bold')).grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=(20, 10))
        
        ttk.Label(frame, text="身份证列名:").grid(row=1, column=0, sticky=tk.W, pady=10)
        self.csv_id_card_entry = ttk.Entry(frame, width=30)
        self.csv_id_card_entry.grid(row=1, column=1, sticky=tk.W, padx=10)
        self.csv_id_card_entry.insert(0, csv_config.get('id_card_column', '公民身份号码'))
        ttk.Label(frame, text="← 公民身份号码", foreground='gray').grid(row=1, column=2, sticky=tk.W)
        
        ttk.Label(frame, text="联系电话列名:").grid(row=2, column=0, sticky=tk.W, pady=10)
        self.csv_phone_entry = ttk.Entry(frame, width=30)
        self.csv_phone_entry.grid(row=2, column=1, sticky=tk.W, padx=10)
        self.csv_phone_entry.insert(0, csv_config.get('phone_column', '联系电话'))
        ttk.Label(frame, text="← 联系电话", foreground='gray').grid(row=2, column=2, sticky=tk.W)
        
        # XLSX格式配置区域
        ttk.Label(frame, text="XLSX 格式配置", font=('Microsoft YaHei', 10, 'bold')).grid(row=3, column=0, columnspan=3, sticky=tk.W, pady=(20, 10))
        
        ttk.Label(frame, text="身份证列名:").grid(row=4, column=0, sticky=tk.W, pady=10)
        self.xlsx_id_card_entry = ttk.Entry(frame, width=30)
        self.xlsx_id_card_entry.grid(row=4, column=1, sticky=tk.W, padx=10)
        self.xlsx_id_card_entry.insert(0, xlsx_config.get('id_card_column', '*证件号码'))
        ttk.Label(frame, text="← *证件号码", foreground='gray').grid(row=4, column=2, sticky=tk.W)
        
        ttk.Label(frame, text="联系电话列名:").grid(row=5, column=0, sticky=tk.W, pady=10)
        self.xlsx_phone_entry = ttk.Entry(frame, width=30)
        self.xlsx_phone_entry.grid(row=5, column=1, sticky=tk.W, padx=10)
        self.xlsx_phone_entry.insert(0, xlsx_config.get('phone_column', '*手机号'))
        ttk.Label(frame, text="← *手机号", foreground='gray').grid(row=5, column=2, sticky=tk.W)
        
        ttk.Label(frame, text="Sheet名称:").grid(row=6, column=0, sticky=tk.W, pady=10)
        self.xlsx_sheet_name_entry = ttk.Entry(frame, width=30)
        self.xlsx_sheet_name_entry.grid(row=6, column=1, sticky=tk.W, padx=10)
        self.xlsx_sheet_name_entry.insert(0, xlsx_config.get('sheet_name', '花名册'))
        ttk.Label(frame, text="← 花名册", foreground='gray').grid(row=6, column=2, sticky=tk.W)
    
    def setup_source_tab(self):
        """
        设置"源表配置"选项卡
        
        UI组件（用于指定源Excel文件格式）：
        - title_row: 标题行索引（提取发薪月份）
        - title_col: 标题所在列
        - header_row_count: 表头行数
        - data_start_row: 数据起始行
        - skip_keyword: 跳过关键字（遇此行停止读取）
        
        默认值：
        - title_row: 0
        - title_col: 2（第3列C列）
        - header_row_count: 5
        - data_start_row: 5
        - skip_keyword: 合计
        """
        frame = ttk.Frame(self.tab_source, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)
        
        source_config = self.config.get('source_config', {})
        
        ttk.Label(frame, text="标题行索引:").grid(row=0, column=0, sticky=tk.W, pady=10)
        self.title_row_entry = ttk.Entry(frame, width=15)
        self.title_row_entry.grid(row=0, column=1, sticky=tk.W, padx=10)
        self.title_row_entry.insert(0, str(source_config.get('title_row', 0)))
        ttk.Label(frame, text="← 标题文本在第几行（0=第1行），用于提取发薪月份", foreground='gray').grid(row=0, column=2, sticky=tk.W)
        
        ttk.Label(frame, text="标题所在列:").grid(row=1, column=0, sticky=tk.W, pady=10)
        self.title_col_entry = ttk.Entry(frame, width=15)
        self.title_col_entry.grid(row=1, column=1, sticky=tk.W, padx=10)
        self.title_col_entry.insert(0, str(source_config.get('title_col', 2)))
        ttk.Label(frame, text="← 标题在第几列（0=第1列，2=第3列即C列）", foreground='gray').grid(row=1, column=2, sticky=tk.W)
        
        ttk.Label(frame, text="表头行数:").grid(row=2, column=0, sticky=tk.W, pady=10)
        self.header_count_entry = ttk.Entry(frame, width=15)
        self.header_count_entry.grid(row=2, column=1, sticky=tk.W, padx=10)
        self.header_count_entry.insert(0, str(source_config.get('header_row_count', 5)))
        ttk.Label(frame, text="← 表头区域占几行，这些行不作为数据读取", foreground='gray').grid(row=2, column=2, sticky=tk.W)
        
        ttk.Label(frame, text="数据起始行:").grid(row=3, column=0, sticky=tk.W, pady=10)
        self.data_start_entry = ttk.Entry(frame, width=15)
        self.data_start_entry.grid(row=3, column=1, sticky=tk.W, padx=10)
        self.data_start_entry.insert(0, str(source_config.get('data_start_row', 5)))
        ttk.Label(frame, text="← 从第几行开始读取数据（0=第1行）", foreground='gray').grid(row=3, column=2, sticky=tk.W)
        
        ttk.Label(frame, text="跳过关键字:").grid(row=4, column=0, sticky=tk.W, pady=10)
        self.skip_keyword_entry = ttk.Entry(frame, width=15)
        self.skip_keyword_entry.grid(row=4, column=1, sticky=tk.W, padx=10)
        self.skip_keyword_entry.insert(0, source_config.get('skip_keyword', '合计'))
        ttk.Label(frame, text='← 遇到包含此文字的行时停止读取（通常用于跳过"合计"行）', foreground='gray').grid(row=4, column=2, sticky=tk.W)
    
    # ================================================================================
    # 固定列tab - Listbox样式
    # 显示fixed_columns数据，动态统计显示固定列数量
    # 支持添加/上移/下移/删除操作
    # 包含"输出未知列"复选框（问题3新增）
    # 
    # 说明：
    # 固定列是按顺序输出、始终显示的列（不论源文件是否有数据）
    # 这些列构成输出表的基础结构
    # ================================================================================
    def setup_fixed_tab(self):
        """
        设置"固定列"选项卡内容
        
        使用Listbox组件显示fixed_columns数据，支持以下操作：
        - 添加：新增固定列名
        - 上移/下移：调整列顺序
        - 删除：移除选中列
        
        顶部包含"输出未知列"复选框（问题3新增）：
        - 开启时：输出源文件中存在但不在column_mapping中的列
        - 关闭时：不输出未知列（默认）
        """
        frame = ttk.Frame(self.tab_fixed, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        
        # 顶部区域：标签 + 复选框
        top_frame = ttk.Frame(frame)
        top_frame.pack(fill=tk.X, pady=(0, 5))
        
        # 动态统计显示：固定列（24项，按顺序输出，始终显示）
        # 判断依据：fixed_columns列表中的列名按顺序作为输出表的前N列
        fixed_count = len(self.fixed_columns)
        ttk.Label(top_frame, text=f"固定列（{fixed_count}项，按顺序输出，始终显示）").pack(side=tk.LEFT)
        
        # ================================================================================
        # 【问题3】输出未知列开关 - 默认关闭（False）
        # 说明：输出源表中存在但不在column_mapping配置中的列
        # 开启条件：用户需要查看源表完整数据时
        # ================================================================================
        self.output_unknown_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(top_frame, text="输出未知列", variable=self.output_unknown_var).pack(side=tk.RIGHT)
        
        list_frame = ttk.Frame(frame)
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.fixed_listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, height=20)
        self.fixed_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.fixed_listbox.yview)
        
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, pady=(5, 0))
        ttk.Button(btn_frame, text="添加", command=self.fixed_add).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="上移", command=self.fixed_move_up).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="下移", command=self.fixed_move_down).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="删除", command=self.fixed_delete).pack(side=tk.LEFT, padx=2)
        
        self.refresh_fixed_list()
        
        if hasattr(self, 'output_unknown_var'):
            self.output_unknown_var.set(self.output_unknown_columns)
    
    def fixed_add(self):
        """
        添加新固定列
        
        弹出对话框让用户输入新列名：
        - 创建模态对话框
        - 输入框聚焦
        - 回车确认，Esc取消
        - 检查列名不重复后添加到fixed_columns
        """
        dialog = tk.Toplevel(self)
        dialog.title("添加固定列")
        dialog.geometry("300x100")
        dialog.transient(self)
        dialog.grab_set()
        
        ttk.Label(dialog, text="列名:").grid(row=0, column=0, sticky=tk.W, padx=10, pady=10)
        entry = ttk.Entry(dialog, width=30)
        entry.grid(row=0, column=1, sticky=tk.W, padx=10)
        entry.focus()
        
        def save():
            col = entry.get().strip()
            if col and col not in self.fixed_columns:
                self.fixed_columns.append(col)
                self.refresh_fixed_list()
            dialog.destroy()
        
        btn_frame = ttk.Frame(dialog)
        btn_frame.grid(row=1, column=0, columnspan=2, pady=10)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="确定", command=save).pack(side=tk.LEFT, padx=5)
        
        # 回车确认，Esc取消
        dialog.bind('<Return>', lambda e: save())
        dialog.bind('<Escape>', lambda e: dialog.destroy())
    
    def refresh_fixed_list(self):
        """
        刷新固定列列表显示
        
        处理逻辑：
        1. 清空Listbox所有项
        2. 重新插入fixed_columns中的所有项
        """
        self.fixed_listbox.delete(0, tk.END)
        for col in self.fixed_columns:
            self.fixed_listbox.insert(tk.END, col)
    
    def fixed_move_up(self):
        """
        将选中列上移一位
        
        交换选中列与上一列的位置
        """
        # 省略注释（方法名已清晰）
        selected = self.fixed_listbox.curselection()
        if selected and selected[0] > 0:
            idx = selected[0]
            self.fixed_columns[idx-1], self.fixed_columns[idx] = self.fixed_columns[idx], self.fixed_columns[idx-1]
            self.refresh_fixed_list()
            self.fixed_listbox.selection_set(idx-1)
    
    def fixed_move_down(self):
        """将选中列下移一位"""
        selected = self.fixed_listbox.curselection()
        if selected and selected[0] < len(self.fixed_columns) - 1:
            idx = selected[0]
            self.fixed_columns[idx], self.fixed_columns[idx+1] = self.fixed_columns[idx+1], self.fixed_columns[idx]
            self.refresh_fixed_list()
            self.fixed_listbox.selection_set(idx+1)
    
    def fixed_delete(self):
        """删除选中列"""
        selected = self.fixed_listbox.curselection()
        if selected:
            idx = selected[0]
            del self.fixed_columns[idx]
            self.refresh_fixed_list()
    
    def reset_to_default(self):
        """
        重置为默认配置
        
        处理逻辑：
        1. 获取默认配置（调用get_default_config()）
        2. 重置fixed_columns和column_mapping
        3. 刷新列表显示
        4. 重置所有UI输入框的值
        
        用途：
        - 用户想恢复出厂设置时使用
        """
        self.config = get_default_config()
        
        self.fixed_columns = self.config.get('column_config', {}).get('fixed_columns', []).copy()
        self.column_mapping = self.config.get('column_config', {}).get('column_mapping', DEFAULT_COLUMN_MAPPING).copy()
        
        self.refresh_fixed_list()
        self.refresh_mapping_list()
        
        # 重置电话查询配置
        phone_config = self.config.get('phone_lookup_config', {})
        csv_config = phone_config.get('csv', {})
        xlsx_config = phone_config.get('xlsx', {})
        self.csv_id_card_entry.delete(0, tk.END)
        self.csv_id_card_entry.insert(0, csv_config.get('id_card_column', '公民身份号码'))
        self.csv_phone_entry.delete(0, tk.END)
        self.csv_phone_entry.insert(0, csv_config.get('phone_column', '联系电话'))
        self.xlsx_id_card_entry.delete(0, tk.END)
        self.xlsx_id_card_entry.insert(0, xlsx_config.get('id_card_column', '*证件号码'))
        self.xlsx_phone_entry.delete(0, tk.END)
        self.xlsx_phone_entry.insert(0, xlsx_config.get('phone_column', '*手机号'))
        self.xlsx_sheet_name_entry.delete(0, tk.END)
        self.xlsx_sheet_name_entry.insert(0, xlsx_config.get('sheet_name', '花名册'))
        
        # 重置源表配置
        source_config = self.config.get('source_config', {})
        self.title_row_entry.delete(0, tk.END)
        self.title_row_entry.insert(0, str(source_config.get('title_row', 0)))
        self.title_col_entry.delete(0, tk.END)
        self.title_col_entry.insert(0, str(source_config.get('title_col', 2)))
        self.header_count_entry.delete(0, tk.END)
        self.header_count_entry.insert(0, str(source_config.get('header_row_count', 5)))
        self.data_start_entry.delete(0, tk.END)
        self.data_start_entry.insert(0, str(source_config.get('data_start_row', 5)))
        self.skip_keyword_entry.delete(0, tk.END)
        self.skip_keyword_entry.insert(0, source_config.get('skip_keyword', '合计'))
    
    def cancel(self):
        """
        取消编辑，关闭窗口
        
        处理逻辑：
        - 设置self.result为None（表示未保存）
        - 销毁窗口
        """
        self.result = None
        self.destroy()
    
    def confirm(self):
        """
        确认配置并保存
        
        从各UI组件获取用户输入，更新config字典：
        1. phone_lookup_config：从电话查询配置选项卡获取CSV/XLSX的列名配置
        2. source_config：从源表配置选项卡获取标题行、表头行数等配置
        3. column_config：从固定列选项卡和字段映射选项卡获取配置
           - output_unknown_columns：输出未知列开关
           - fixed_columns：固定列列表
           - column_mapping：字段映射列表
        
        使用getattr/hasattr进行防御性检查（问题5修复）：
        - 避免用户未访问某些选项卡时，相关widget不存在导致AttributeError
        
        返回值：无（设置self.result后关闭窗口）
        """
        # 获取电话查询配置 - CSV格式
        self.config['phone_lookup_config'] = {
            'csv': {
                'id_card_column': self.csv_id_card_entry.get().strip(),
                'phone_column': self.csv_phone_entry.get().strip()
            },
            'xlsx': {
                'id_card_column': self.xlsx_id_card_entry.get().strip(),
                'phone_column': self.xlsx_phone_entry.get().strip(),
                'sheet_name': self.xlsx_sheet_name_entry.get().strip()
            }
        }
        
        # 获取源表配置 - 使用防御性检查
        # 判断依据：用户可能未访问源表配置选项卡，相关widget可能不存在
        self.config['source_config'] = {
            'title_row': int(getattr(self, 'title_row_entry', None).get().strip() or 0) if hasattr(self, 'title_row_entry') else 0,
            'title_col': int(getattr(self, 'title_col_entry', None).get().strip() or 2) if hasattr(self, 'title_col_entry') else 2,
            'header_row_count': int(getattr(self, 'header_count_entry', None).get().strip() or 5) if hasattr(self, 'header_count_entry') else 5,
            'data_start_row': int(getattr(self, 'data_start_entry', None).get().strip() or 5) if hasattr(self, 'data_start_entry') else 5,
            'skip_keyword': getattr(self, 'skip_keyword_entry', None).get().strip() or '合计' if hasattr(self, 'skip_keyword_entry') else '合计'
        }
        
        # 获取列配置 - 包括固定列、字段映射、未知列开关
        self.config['column_config'] = {
            'output_unknown_columns': getattr(self, 'output_unknown_var', None).get() if hasattr(self, 'output_unknown_var') else False,
            'fixed_columns': self.fixed_columns,
            'column_mapping': self.column_mapping
        }
        
        self.result = self.config
        self.destroy()


class ExcelConverterApp:
    """
    主应用程序窗口 - 东软-薪福通工资表格式转换工具
    
    属性:
        root: tk.Tk - 主窗口
        file_list: List[Dict] - 待转换文件信息列表
        selected_index: int - 当前选中的文件索引
        output_dir: tk.StringVar - 输出目录路径
        export_csv_path: tk.StringVar - 人员信息CSV/XLSX文件路径
        csv_create_time: tk.StringVar - CSV文件创建时间
        app_config: Dict - 应用程序配置
    """
    
    def __init__(self, root):
        """
        初始化主应用程序窗口
        
        参数:
            root: tk.Tk - Tkinter根窗口（由tkinterdnd2.Tk()创建以支持拖放）
        """
        self.root = root
        self.root.title("东软-薪福通工资表格式转换工具")
        self.root.geometry("900x600")
        self.root.resizable(True, True)
        
        # 文件列表：存储待转换文件的详细信息
        self.file_list = []
        self.selected_index = -1
        
        # 输出目录和人员信息文件路径（用于tkinter变量绑定）
        self.output_dir = tk.StringVar()
        self.export_csv_path = tk.StringVar()
        self.csv_create_time = tk.StringVar()
        
        # 加载配置并应用
        self.app_config = load_config_and_apply()
        
        # 初始化UI、拖放、加载最近CSV、设置关闭协议
        self.setup_ui()
        self.setup_drag_drop()
        self.load_recent_csv()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def setup_ui(self):
        """
        设置主窗口UI布局
        
        布局结构：
        1. 标题栏：应用名称
        2. 内容区（左右分栏）：
           - 左侧：文件列表（Listbox + 按钮），支持拖放
           - 右侧：文件预览信息（文件名、标题、发薪月份等）
        3. 底部区域：
           - 输出目录选择
           - 人员信息文件拖放/选择
           - 编辑配置按钮
           - 开始转换按钮
        4. 状态栏：显示当前状态
        5. 日志区：显示转换执行日志
        
        组件说明：
        - file_listbox: Listbox，支持多选，绑定选择事件
        - csv_drop_frame: 支持拖放的Frame，绑定DND_FILES事件
        - convert_btn: 转换按钮，转换时禁用防止重复点击
        - log_text: Text组件，显示INFO/WARNING/ERROR级别日志
        """
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题栏
        title_label = ttk.Label(main_frame, text="东软-薪福通工资表格式转换工具", font=('song', 16, 'bold'))
        title_label.pack(pady=(0, 10))
        
        # 内容区：左右分栏
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # 左侧：文件列表区域
        left_frame = ttk.LabelFrame(content_frame, text="文件列表", padding="5")
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        list_frame = ttk.Frame(left_frame)
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 文件列表：支持多选模式，绑定选择事件
        self.file_listbox = tk.Listbox(list_frame, selectmode=tk.EXTENDED, 
                                        yscrollcommand=scrollbar.set,
                                        font=('song', 10))
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.file_listbox.yview)
        
        self.file_listbox.bind('<<ListboxSelect>>', self.on_file_select)
        
        # 文件列表按钮：添加/移除/清空
        btn_frame = ttk.Frame(left_frame)
        btn_frame.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Button(btn_frame, text="添加文件", command=self.add_files).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="移除选中", command=self.remove_selected).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="清空列表", command=self.clear_list).pack(side=tk.LEFT, padx=2)
        
        drop_hint = ttk.Label(left_frame, text="(或拖放文件到此处)", font=('song', 9), foreground='gray')
        drop_hint.pack(pady=(5, 0))
        
        # 右侧：文件预览区域
        right_frame = ttk.LabelFrame(content_frame, text="文件预览", padding="10")
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))
        
        info_frame = ttk.Frame(right_frame)
        info_frame.pack(fill=tk.BOTH, expand=True)
        
        # 预览信息标签：文件名、标题、发薪月份、数据行数、文件大小
        ttk.Label(info_frame, text="文件名:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.label_filename = ttk.Label(info_frame, text="-", foreground='blue')
        self.label_filename.grid(row=0, column=1, sticky=tk.W, padx=(10, 0), pady=5)
        
        ttk.Label(info_frame, text="标题:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.label_title = ttk.Label(info_frame, text="-", foreground='blue', wraplength=350)
        self.label_title.grid(row=1, column=1, sticky=tk.W, padx=(10, 0), pady=5)
        
        ttk.Label(info_frame, text="发薪月份:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.label_yearmonth = ttk.Label(info_frame, text="-", foreground='blue')
        self.label_yearmonth.grid(row=2, column=1, sticky=tk.W, padx=(10, 0), pady=5)
        
        ttk.Label(info_frame, text="数据行数:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.label_rows = ttk.Label(info_frame, text="-", foreground='blue')
        self.label_rows.grid(row=3, column=1, sticky=tk.W, padx=(10, 0), pady=5)
        
        ttk.Label(info_frame, text="文件大小:").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.label_size = ttk.Label(info_frame, text="-", foreground='blue')
        self.label_size.grid(row=4, column=1, sticky=tk.W, padx=(10, 0), pady=5)
        
        for i in range(5):
            info_frame.rowconfigure(i, pad=5)
        info_frame.columnconfigure(1, pad=10)
        
        # 底部区域：输出目录 + 人员信息
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X, pady=(10, 0))
        
        # 输出目录选择
        ttk.Label(bottom_frame, text="输出目录:").pack(side=tk.LEFT)
        
        output_entry = ttk.Entry(bottom_frame, textvariable=self.output_dir, width=50)
        output_entry.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(bottom_frame, text="浏览...", command=self.browse_output_dir).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(bottom_frame, text="编辑配置", command=self.open_config_editor).pack(side=tk.LEFT, padx=20)
        
        # 开始转换按钮（转换时禁用）
        self.convert_btn = ttk.Button(bottom_frame, text="开始转换", command=self.start_convert)
        self.convert_btn.pack(side=tk.RIGHT)
        
        # 人员信息文件拖放区域
        phone_frame = ttk.Frame(main_frame)
        phone_frame.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Label(phone_frame, text="人员信息:").pack(side=tk.LEFT)
        
        # 支持拖放的Frame：用于放置CSV/员工模板文件
        self.csv_drop_frame = tk.Frame(phone_frame, bd=2, relief=tk.SUNKEN, bg='white', height=30)
        self.csv_drop_frame.pack(side=tk.LEFT, padx=5)
        self.csv_drop_frame.pack_propagate(False)
        self.csv_drop_frame.configure(width=350)
        
        self.csv_drop_label = tk.Label(self.csv_drop_frame, text="拖拽 CSV/员工模板文件到此处 或 点击选择", 
                                        bg='white', fg='gray', font=('song', 9), anchor='w')
        self.csv_drop_label.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 绑定拖放和点击事件
        self.csv_drop_frame.drop_target_register(DND_FILES)
        self.csv_drop_frame.dnd_bind('<<Drop>>', self.on_csv_drop)
        self.csv_drop_frame.bind('<Button-1>', lambda e: self.on_csv_click())
        
        ttk.Button(phone_frame, text="浏览...", command=self.on_csv_click).pack(side=tk.LEFT, padx=5)
        
        csv_info_frame = ttk.Frame(main_frame)
        csv_info_frame.pack(fill=tk.X, pady=(2, 0))
        
        self.csv_info_label = ttk.Label(csv_info_frame, text="", foreground='gray', font=('song', 8))
        self.csv_info_label.pack(side=tk.LEFT)
        
        # 状态栏
        self.status_label = ttk.Label(main_frame, text="就绪", foreground='green')
        self.status_label.pack(pady=(5, 0))
        
        # 执行日志区域
        log_frame = ttk.LabelFrame(main_frame, text="执行日志", padding="5")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        
        log_scroll_y = ttk.Scrollbar(log_frame)
        log_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 日志Text组件：支持滚动，根据级别显示不同颜色
        self.log_text = tk.Text(log_frame, height=8, font=('Consolas', 9), 
                                yscrollcommand=log_scroll_y.set, wrap=tk.WORD)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scroll_y.config(command=self.log_text.yview)
    
    def setup_drag_drop(self):
        """
        设置拖放功能
        
        将文件列表Listbox注册为拖放目标，绑定<<Drop>>事件到on_excel_drop
        """
        self.file_listbox.drop_target_register(DND_FILES)
        self.file_listbox.dnd_bind('<<Drop>>', self.on_excel_drop)
    
    def on_excel_drop(self, event):
        """
        处理Excel文件拖放事件
        
        参数:
            event: tkinterdnd2拖放事件
        
        处理逻辑：
        1. 解析拖放的文件列表
        2. 处理带大括号的路径格式
        3. 过滤只保留.xls/.xlsx文件
        4. 添加到文件列表（不重复）
        """
        files = self.root.tk.splitlist(event.data)
        for file_path in files:
            file_path = file_path.strip()
            if file_path.startswith('{') and file_path.endswith('}'):
                file_path = file_path[1:-1]
            if file_path and (file_path.endswith('.xls') or file_path.endswith('.xlsx')):
                if os.path.exists(file_path):
                    if file_path not in [item['path'] for item in self.file_list]:
                        self.add_file_to_list(file_path)
        return 'action'
    
    def on_csv_drop(self, event):
        """
        处理CSV文件拖放事件
        
        参数:
            event: tkinterdnd2拖放事件
        
        处理逻辑：
        1. 解析拖放的文件
        2. 过滤只保留.csv/.xlsx文件
        3. 调用set_csv_path设置人员信息文件
        """
        files = self.root.tk.splitlist(event.data)
        for file_path in files:
            file_path = file_path.strip()
            if file_path.startswith('{') and file_path.endswith('}'):
                file_path = file_path[1:-1]
            if file_path and (file_path.endswith('.csv') or file_path.endswith('.xlsx')):
                self.set_csv_path(file_path)
                break
        return 'action'
    
    def on_csv_click(self):
        """
        处理人员信息文件点击选择
        
        弹出文件选择对话框，选择后调用set_csv_path
        """
        filepath = filedialog.askopenfilename(
            title="选择人员信息",
            filetypes=[("支持的文件", "*.csv;*.xlsx"), ("CSV 文件", "*.csv"), 
                       ("Excel 文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if filepath:
            self.set_csv_path(filepath)
    
    def set_csv_path(self, filepath):
        """
        设置人员信息文件路径并更新UI显示
        
        参数:
            filepath: str - CSV/XLSX文件路径
        
        更新：
        - export_csv_path变量
        - csv_drop_label显示文件名
        - csv_info_label显示创建时间
        """
        self.export_csv_path.set(filepath)
        filename = os.path.basename(filepath)
        create_time = get_file_create_time(filepath)
        
        if create_time:
            self.csv_drop_label.config(text=f"{filename}", fg='black', bg='white')
            self.csv_info_label.config(text=f"创建时间: {create_time}", foreground='gray')
        else:
            self.csv_drop_label.config(text=f"{filename}", fg='black', bg='white')
            self.csv_info_label.config(text="创建时间: 未知", foreground='red')
    
    def load_recent_csv(self):
        """
        加载最近使用的人员信息文件
        
        从app_config读取recent_csv配置：
        - 如果文件存在：显示文件名和创建时间
        - 如果文件不存在：显示默认提示
        """
        recent_csv = self.app_config.get('recent_csv', {})
        csv_path = recent_csv.get('path', '')
        csv_time = recent_csv.get('created_time', '')
        if csv_path and os.path.exists(csv_path):
            self.export_csv_path.set(csv_path)
            filename = os.path.basename(csv_path)
            if csv_time:
                self.csv_drop_label.config(text=f"{filename}", fg='black', bg='white')
                self.csv_info_label.config(text=f"创建时间: {csv_time}", foreground='gray')
            else:
                self.csv_drop_label.config(text=f"{filename}", fg='black', bg='white')
                self.csv_info_label.config(text="创建时间: 未知", foreground='red')
        else:
            self.csv_drop_label.config(text="拖拽 CSV 文件到此处 或 点击选择", fg='gray', bg='white')
            self.csv_info_label.config(text="", foreground='gray')
    
    def open_config_editor(self):
        """
        打开配置编辑器
        
        弹出配置编辑器窗口，编辑后保存配置并重新加载
        """
        editor = ConfigEditor(self.root, self.app_config)
        self.root.wait_window(editor)
        if editor.result:
            self.app_config = editor.result
            save_config(self.app_config)
            load_config_and_apply()
            messagebox.showinfo("提示", "配置已保存")
    
    def save_recent_csv(self):
        """
        保存最近使用的人员信息文件路径
        
        将当前CSV路径和创建时间保存到配置文件
        """
        csv_path = self.export_csv_path.get()
        if csv_path and os.path.exists(csv_path):
            create_time = get_file_create_time(csv_path)
            self.app_config['recent_csv'] = {
                'path': csv_path,
                'created_time': create_time or ''
            }
            save_config(self.app_config)
    
    def add_files(self):
        """
        添加Excel文件
        
        弹出文件选择对话框，选择后添加到文件列表
        """
        files = filedialog.askopenfilenames(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xls *.xlsx"), ("所有文件", "*.*")]
        )
        for f in files:
            if f not in [item['path'] for item in self.file_list]:
                self.add_file_to_list(f)
    
    def add_file_to_list(self, filepath):
        """
        添加单个文件到列表
        
        参数:
            filepath: str - Excel文件路径
        
        调用get_file_info获取文件信息，并添加到列表显示
        """
        info = get_file_info(filepath)
        if info:
            self.file_list.append(info)
            display_text = f"{info['filename']} ({info['year_month']}) - {info['data_rows']}行"
            self.file_listbox.insert(tk.END, display_text)
    
    def remove_selected(self):
        """
        移除选中的文件
        
        从后向前遍历删除（防止索引变化）
        """
        selected = self.file_listbox.curselection()
        for i in reversed(selected):
            self.file_listbox.delete(i)
            del self.file_list[i]
        self.clear_preview()
    
    def clear_list(self):
        """
        清空文件列表
        """
        self.file_listbox.delete(0, tk.END)
        self.file_list = []
        self.clear_preview()
    
    def clear_preview(self):
        """
        清空文件预览信息
        """
        self.label_filename.config(text="-")
        self.label_title.config(text="-")
        self.label_yearmonth.config(text="-")
        self.label_rows.config(text="-")
        self.label_size.config(text="-")
        self.selected_index = -1
    
    def on_file_select(self, event):
        """
        处理文件列表选择事件
        
        参数:
            event: tkinter事件（<<ListboxSelect>>）
        
        更新右侧预览区域：
        - 文件名
        - 标题
        - 发薪月份
        - 数据行数
        - 文件大小（自动单位转换）
        """
        selection = self.file_listbox.curselection()
        if selection:
            self.selected_index = selection[0]
            info = self.file_list[self.selected_index]
            
            self.label_filename.config(text=info['filename'])
            self.label_title.config(text=info['title'] if info['title'] else "(无标题)")
            self.label_yearmonth.config(text=info['year_month'] if info['year_month'] else "(无法提取)")
            self.label_rows.config(text=str(info['data_rows']))
            
            # 文件大小单位转换：B -> KB -> MB
            size = info['file_size']
            if size < 1024:
                size_str = f"{size} B"
            elif size < 1024 * 1024:
                size_str = f"{size / 1024:.1f} KB"
            else:
                size_str = f"{size / (1024 * 1024):.1f} MB"
            self.label_size.config(text=size_str)
    
    def browse_output_dir(self):
        """
        浏览并选择输出目录
        
        弹出目录选择对话框，选择后设置output_dir变量
        """
        dirpath = filedialog.askdirectory(title="选择输出目录")
        if dirpath:
            self.output_dir.set(dirpath)
    
    def start_convert(self):
        """
        开始执行批量转换
        
        处理流程：
        1. 验证：检查是否有待转换文件
        2. 准备输出目录：创建输出目录和日志目录
        3. 加载人员信息：从CSV/XLSX文件加载电话查询数据和新员工信息
        4. 逐个转换：遍历file_list，调用convert_single_file()
        5. 保存日志：将执行日志追加到日志文件
        6. 显示结果：汇总成功/失败数量，显示结果消息框
        
        注意事项：
        - 转换过程中禁用"开始转换"按钮，防止重复点击
        - 实时更新状态栏显示当前进度
        - 每个文件转换后立即记录日志
        - 警告信息（缺失姓名/身份证/手机号）单独记录
        """
        # 验证：检查是否有待转换文件
        if not self.file_list:
            messagebox.showwarning("提示", "请先添加要转换的文件")
            return
        
        # 获取或设置默认输出目录
        output_dir = self.output_dir.get()
        if not output_dir:
            if self.file_list:
                first_file = self.file_list[0]['path']
                default_dir = os.path.join(os.path.dirname(first_file), 'Converted')
                self.output_dir.set(default_dir)
                output_dir = default_dir
        
        # 创建输出目录和日志目录
        os.makedirs(output_dir, exist_ok=True)
        
        log_dir = os.path.join(output_dir, 'log')
        os.makedirs(log_dir, exist_ok=True)
        log_file = os.path.join(log_dir, f"convert_{datetime.datetime.now().strftime('%Y%m%d')}.log")
        
        # 加载人员信息文件
        export_csv = self.export_csv_path.get()
        # phone_lookup: 从人员信息CSV文件加载的 {证件号码: 手机号} 字典，传入convert_single_file
        phone_lookup = load_phone_lookup(export_csv)
        
        # 如果是XLSX格式，尝试加载新增人员信息
        # new_employees: 从XLSX模板加载的 {证件号码: {name, phone}} 字典，传入convert_single_file
        new_employees = None
        if export_csv and os.path.exists(export_csv):
            if export_csv.endswith('.xlsx'):
                new_employees = load_new_employees_from_template(export_csv)
                if new_employees:
                    self.log(f"已加载新增人员数据: {len(new_employees)} 人", "INFO")
        
        # 清空日志并显示开始信息
        self.log_text.delete(1.0, tk.END)
        self.log(f"开始转换，共 {len(self.file_list)} 个文件", "INFO")
        self.log(f"输出目录: {output_dir}", "INFO")
        self.log(f"日志文件: {log_file}", "INFO")
        
        # 禁用转换按钮，防止重复点击
        self.convert_btn.config(state=tk.DISABLED)
        success_count = 0
        fail_count = 0
        results = []
        
        # 逐个转换文件
        for i, file_info in enumerate(self.file_list):
            # 更新状态栏显示当前进度
            self.status_label.config(text=f"正在转换 ({i+1}/{len(self.file_list)}): {file_info['filename']}")
            self.root.update()
            
            self.log(f"开始处理: {file_info['filename']}", "INFO")
            
            # 调用核心转换函数
            success, msg, detail = convert_single_file(file_info['path'], output_dir, phone_lookup, new_employees)
            
            # 处理转换结果
            if success:
                success_count += 1
                # 记录警告信息（缺失字段）
                if detail.get('warnings'):
                    for w in detail['warnings']:
                        self.log(f"  ⚠️ 第 {w['row']} 行: 缺失 [{', '.join(w['missing_fields'])}]", "WARNING")
                self.log(f"完成: {msg}", "INFO")
            else:
                fail_count += 1
                self.log(f"失败: {msg}", "ERROR")
            
            results.append((file_info['filename'], success, msg))
        
        # 保存日志到文件
        log_content = self.get_log_content()
        try:
            with open(log_file, 'a', encoding='utf-8') as f:
                f.write(log_content)
        except Exception as e:
            self.log(f"保存日志文件失败: {e}", "ERROR")
        
        # 恢复转换按钮
        self.convert_btn.config(state=tk.NORMAL)
        
        # 显示汇总信息
        self.log("-" * 40, "INFO")
        self.log(f"汇总: 成功 {success_count} 个, 失败 {fail_count} 个", "INFO")
        self.log(f"日志已保存: {log_file}", "INFO")
        
        # 构建结果消息
        result_msg = f"转换完成！\n\n成功: {success_count} 个文件\n失败: {fail_count} 个文件\n\n日志已保存至:\n{log_file}"
        if fail_count > 0:
            result_msg += "\n\n失败文件:"
            for fname, success, msg in results:
                if not success:
                    result_msg += f"\n- {fname}: {msg}"
        
        self.status_label.config(text=f"完成 - 成功: {success_count}, 失败: {fail_count}")
        messagebox.showinfo("转换结果", result_msg)
    
    def on_closing(self):
        """
        窗口关闭处理
        
        保存最近使用的CSV路径，然后销毁窗口
        """
        self.save_recent_csv()
        self.root.destroy()
    
    def log(self, message: str, level: str = "INFO"):
        """
        记录日志消息
        
        参数:
            message: str - 日志消息内容
            level: str - 日志级别，可选值：INFO（默认）、WARNING、ERROR
        
        处理流程：
        1. 添加时间戳前缀，格式：[HH:MM:SS] LEVEL message
        2. 插入到log_text组件末尾
        3. 根据级别设置文字颜色：
           - WARNING: 橙色
           - ERROR: 红色
        4. 自动滚动到末尾
        
        颜色配置：
        - INFO: 默认颜色（黑色）
        - WARNING: 橙色foreground
        - ERROR: 红色foreground
        """
        # 添加时间戳并格式化消息
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        formatted = f"[{timestamp}] {level} {message}\n"
        
        # 插入到Text组件末尾
        self.log_text.insert(tk.END, formatted)
        
        # 根据级别设置颜色
        if level == "WARNING":
            self.log_text.tag_add("warning", self.log_text.index(tk.END + "-2l"), tk.END)
            self.log_text.tag_config("warning", foreground="orange")
        elif level == "ERROR":
            self.log_text.tag_add("error", self.log_text.index(tk.END + "-2l"), tk.END)
            self.log_text.tag_config("error", foreground="red")
        
        # 自动滚动到末尾
        self.log_text.see(tk.END)
    
    def get_log_content(self) -> str:
        """
        获取日志文本内容
        
        返回值:
            str - 日志Text组件的全部内容（1.0到END）
        
        用途：
        - 用于保存日志到文件
        """
        return self.log_text.get(1.0, tk.END)


def main():
    """
    程序入口函数
    
    初始化步骤：
    1. 创建支持拖放功能的Tk根窗口（使用tkinterdnd2.Tk替代tkinter.Tk）
    2. 创建ExcelConverterApp主应用程序实例
    3. 启动Tkinter事件循环
    
    说明：
    - tkinterdnd2.Tk 提供拖放功能支持
    - 所有GUI操作通过事件循环处理
    """
    # 创建支持拖放的Tk根窗口
    root = tkinterdnd2.Tk()
    # 创建主应用程序
    app = ExcelConverterApp(root)
    # 启动事件循环
    root.mainloop()


if __name__ == '__main__':
    main()
