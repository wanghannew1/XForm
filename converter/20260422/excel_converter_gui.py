#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
东软-薪福通工资表格式转换工具 - GUI 版本
支持文件选择、拖放、批量转换
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

# 默认的字段映射配置（合并DEFAULT_FIELD_MAPPING和dynamic_columns）
# 结构：[{"source": "源列名", "target": "目标列名"}, ...]
DEFAULT_COLUMN_MAPPING = [
    # 固定列映射（约21项）
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
    global TARGET_HEADER_ROW1, TARGET_HEADER_ROW2, TARGET_FIELD_INDEX
    
    if column_config is None:
        column_config = get_default_config().get('column_config', {})
    
    fixed_columns = column_config.get('fixed_columns', [])
    dynamic_columns = column_config.get('dynamic_columns', [])
    
    TARGET_HEADER_ROW2 = fixed_columns.copy()
    
    for item in dynamic_columns:
        TARGET_HEADER_ROW2.append(item['target'])
    
    fixed_category = {
        '工号': '工号', '姓名': '姓名', '证件号码': '证件号码', '手机号': '手机号', '发薪月份': '发薪月份',
        '基本工资': '总计', '补发工资': '总计', '应发工资': '总计', '实发金额': '总计',
        '社保基数': '扣缴基数', '医保基数': '扣缴基数', '工伤基数': '扣缴基数', '公积金基数': '扣缴基数',
        '扣单位养老保险': '应扣项', '扣个人养老保险': '应扣项', '扣单位失业保险': '应扣项', '扣个人失业保险': '应扣项',
        '扣单位医疗保险': '应扣项', '扣个人医疗保险': '应扣项', '扣单位工伤险': '应扣项', '扣单位住房公积金': '应扣项',
        '扣个人住房公积金': '应扣项', '扣缴税金（扣个人所得税）': '应扣项', '2025年个税补退': '应扣项'
    }
    
    TARGET_HEADER_ROW1 = []
    for col in TARGET_HEADER_ROW2:
        if col in fixed_category:
            TARGET_HEADER_ROW1.append(fixed_category[col])
        else:
            TARGET_HEADER_ROW1.append('其它')
    
    TARGET_FIELD_INDEX = {name: idx for idx, name in enumerate(TARGET_HEADER_ROW2)}


# 运行时字段映射（从配置加载）
FIELD_MAPPING = {}


# ==================== 配置加载/保存 ====================

def get_default_config() -> Dict:
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
    config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), CONFIG_FILE)
    if os.path.exists(config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    default_cfg = get_default_config()
    save_config(default_cfg)
    return default_cfg


def save_config(config: Dict):
    config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), CONFIG_FILE)
    try:
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"保存配置失败: {e}")


def get_file_create_time(path: str) -> Optional[str]:
    if not path or not os.path.exists(path):
        return None
    try:
        stat = os.stat(path)
        mtime = datetime.datetime.fromtimestamp(stat.st_mtime)
        return mtime.strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        return None


def load_config_and_apply():
    global FIELD_MAPPING
    config = load_config()
    column_config = config.get('column_config', get_default_config()['column_config'])
    column_mapping = column_config.get('column_mapping', DEFAULT_COLUMN_MAPPING)
    FIELD_MAPPING = {item['source']: item['target'] for item in column_mapping}
    build_target_headers(column_config)
    return config


# ==================== 转换逻辑 ====================

def extract_year_month(title: str) -> str:
    match = re.search(r'(\d{4})年(\d{1,2})月', title)
    if match:
        year, month = match.groups()
        return f"{year}{month.zfill(2)}"
    return ''


def read_source_file(filepath: str) -> Tuple[Dict[str, Any], List[List[Any]], List[Dict[str, Any]]]:
    config = load_config()
    if filepath.lower().endswith('.xlsx'):
        return _read_source_file_xlsx(filepath, config)
    else:
        return _read_source_file_xls(filepath, config)


def _read_source_file_xls(filepath: str, config: Dict = None) -> Tuple[Dict[str, Any], List[List[Any]], List[Dict[str, Any]]]:
    if config is None:
        config = load_config()
    
    source_config = config.get('source_config', {})
    title_row = source_config.get('title_row', 0)
    title_col = source_config.get('title_col', 2)
    header_count = source_config.get('header_row_count', 5)
    data_start = source_config.get('data_start_row', 5)
    skip_keyword = source_config.get('skip_keyword', '合计')
    
    book = xlrd.open_workbook(filepath)
    sheet = book.sheet_by_index(0)
    
    title = sheet.cell_value(title_row, title_col) if sheet.nrows > title_row else ''
    year_month = extract_year_month(title)
    
    meta = {
        'title': title,
        'year_month': year_month,
        'filename': os.path.basename(filepath),
    }
    
    header_rows = []
    for row_idx in range(header_count):
        row_data = [sheet.cell_value(row_idx, col_idx) for col_idx in range(sheet.ncols)]
        header_rows.append(row_data)
    
    data_rows = []
    for row_idx in range(data_start, sheet.nrows):
        first_cell = sheet.cell_value(row_idx, 0)
        if skip_keyword and first_cell == skip_keyword:
            break
        if isinstance(first_cell, float) or (isinstance(first_cell, str) and first_cell.isdigit()):
            row_data = [sheet.cell_value(row_idx, col_idx) for col_idx in range(sheet.ncols)]
            data_rows.append(row_data)
    
    return meta, header_rows, data_rows


def _read_source_file_xlsx(filepath: str, config: Dict = None) -> Tuple[Dict[str, Any], List[List[Any]], List[Dict[str, Any]]]:
    if config is None:
        config = load_config()
    
    source_config = config.get('source_config', {})
    title_row = source_config.get('title_row', 0)
    title_col = source_config.get('title_col', 2)
    header_count = source_config.get('header_row_count', 5)
    data_start = source_config.get('data_start_row', 5)
    skip_keyword = source_config.get('skip_keyword', '合计')
    

    book = openpyxl.load_workbook(filepath, data_only=True)
    sheet = book.active
    
    title = sheet.cell(title_row + 1, title_col + 1).value if sheet.max_row > title_row else ''
    year_month = extract_year_month(str(title) if title else '')
    
    meta = {
        'title': str(title) if title else '',
        'year_month': year_month,
        'filename': os.path.basename(filepath),
    }
    
    header_rows = []
    for row_idx in range(1, header_count + 1):
        row_data = [sheet.cell(row_idx, col_idx).value for col_idx in range(1, sheet.max_column + 1)]
        header_rows.append(row_data)
    
    data_rows = []
    for row_idx in range(data_start + 1, sheet.max_row + 1):
        first_cell = sheet.cell(row_idx, 1).value
        if skip_keyword and first_cell == skip_keyword:
            break
        if isinstance(first_cell, (int, float)) or (isinstance(first_cell, str) and first_cell.isdigit()):
            row_data = [sheet.cell(row_idx, col_idx).value for col_idx in range(1, sheet.max_column + 1)]
            data_rows.append(row_data)
    
    book.close()
    return meta, header_rows, data_rows
    
    book.close()
    return meta, header_rows, data_rows


def extract_source_headers(header_rows: List[List[Any]]) -> List[Tuple[str, int]]:
    main_header = header_rows[2]
    sub_header = header_rows[3] if len(header_rows) > 3 else [''] * len(main_header)
    third_header = header_rows[4] if len(header_rows) > 4 else [''] * len(main_header)
    
    headers = []
    col_idx = 0
    
    while col_idx < len(main_header):
        main_val = str(main_header[col_idx]).strip() if main_header[col_idx] else ''
        
        if not main_val:
            col_idx += 1
            continue
        
        # 检查是否跨列：向右查找连续的列（这些列的 main_header 为空）
        next_main_col = col_idx + 1
        while next_main_col < len(main_header):
            next_main = str(main_header[next_main_col]).strip() if main_header[next_main_col] else ''
            if next_main:
                break
            next_main_col += 1
        
        span = next_main_col - col_idx
        
        if span > 1:
            # 跨列字段：遍历范围内的每一列
            # 清理主表头的多余空格
            main_val_clean = re.sub(r'\s+', '', main_val)
            prev_sub_val = ''  # 用于继承上一列的 sub_val
            for c in range(col_idx, next_main_col):
                sub_val = str(sub_header[c]).strip() if sub_header[c] else ''
                third_val = str(third_header[c]).strip() if third_header[c] else ''
                
                # 如果 sub_val 为空但 third_val 不为空，继承上一列的 sub_val
                if not sub_val and third_val:
                    sub_val = prev_sub_val
                
                # 更新 prev_sub_val
                if sub_val:
                    prev_sub_val = sub_val
                
                # 拼接列名
                if sub_val:
                    if third_val:
                        col_name = f"{main_val_clean}_{sub_val}_{third_val}"
                    else:
                        col_name = f"{main_val_clean}_{sub_val}"
                else:
                    col_name = main_val_clean
                
                headers.append((col_name, c))
        else:
            # 单列字段
            headers.append((main_val, col_idx))
        
        col_idx = next_main_col
    
    return headers


def map_row_to_target(row: List[Any], source_headers: List[Tuple[str, int]]) -> List[Any]:
    source_data = {}
    for col_name, col_idx in source_headers:
        if col_idx < len(row):
            source_data[col_name] = row[col_idx]
    
    target_row = [''] * 29
    
    for source_name, target_name in FIELD_MAPPING.items():
        if source_name in source_data:
            target_idx = TARGET_FIELD_INDEX.get(target_name)
            if target_idx is not None:
                target_row[target_idx] = source_data[source_name]
    
    return target_row


def clean_excel_formula(value: str) -> str:
    if not value:
        return ''
    if value.startswith('="') and value.endswith('"'):
        return value[2:-1]
    return value.strip()


def load_phone_lookup(csv_path: str, id_col_name: str = None, phone_col_name: str = None) -> Dict[str, str]:
    global FIELD_MAPPING
    lookup = {}
    if not csv_path or not os.path.exists(csv_path):
        return lookup
    config = load_config()
    phone_config = config.get('phone_lookup_config', {})
    
    if csv_path.endswith('.xlsx'):
        xlsx_config = phone_config.get('xlsx', {})
        if id_col_name is None:
            id_col_name = xlsx_config.get('id_card_column', '*证件号码')
        if phone_col_name is None:
            phone_col_name = xlsx_config.get('phone_column', '*手机号')
        sheet_name = xlsx_config.get('sheet_name', '花名册')
        return load_phone_lookup_from_xlsx(csv_path, id_col_name, phone_col_name, sheet_name)
    
    csv_config = phone_config.get('csv', {})
    if id_col_name is None:
        id_col_name = csv_config.get('id_card_column', '公民身份号码')
    if phone_col_name is None:
        phone_col_name = csv_config.get('phone_column', '联系电话')
    
    try:
        with open(csv_path, 'r', encoding='gb18030') as f:
            reader = csv.reader(f)
            header = next(reader)
            id_col = -1
            phone_col = -1
            for i, col_name in enumerate(header):
                if col_name == id_col_name:
                    id_col = i
                elif col_name == phone_col_name:
                    phone_col = i
                if id_col >= 0 and phone_col >= 0:
                    break
            if id_col < 0 or phone_col < 0:
                return lookup
            for row in reader:
                if len(row) > max(id_col, phone_col):
                    id_card = clean_excel_formula(row[id_col])
                    phone = clean_excel_formula(row[phone_col])
                    if id_card and phone:
                        lookup[id_card] = phone
    except Exception:
        pass
    return lookup


def load_phone_lookup_from_xlsx(xlsx_path: str, id_col_name: str, 
                                 phone_col_name: str, sheet_name: str) -> Dict[str, str]:
    """读取员工导入模板202512.xlsx格式的电话查询表"""
    lookup = {}
    if not xlsx_path or not os.path.exists(xlsx_path):
        return lookup
    
    try:
        wb = openpyxl.load_workbook(xlsx_path)
        ws = wb[sheet_name]
        
        header = [cell.value for cell in ws[1]]
        id_col = header.index(id_col_name) if id_col_name in header else -1
        phone_col = header.index(phone_col_name) if phone_col_name in header else -1
        
        if id_col < 0 or phone_col < 0:
            return lookup
        
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
    """从员工导入模板中读取新增人员信息
    
    优先从"新增人员"sheet读取，如果没有则从"花名册"sheet读取
    返回: {证件号码: {name, phone}}
    """
    lookup = {}
    if not xlsx_path or not os.path.exists(xlsx_path):
        return lookup
    
    try:
        wb = openpyxl.load_workbook(xlsx_path)
        
        sheet_name = None
        if '新增人员' in wb.sheetnames:
            sheet_name = '新增人员'
        elif '花名册' in wb.sheetnames:
            sheet_name = '花名册'
        
        if not sheet_name:
            return lookup
        
        ws = wb[sheet_name]
        header = [cell.value for cell in ws[1]]
        
        name_col = next((i for i, h in enumerate(header) if h and '姓名' in str(h)), 1)
        phone_col = next((i for i, h in enumerate(header) if h and '手机' in str(h)), 2)
        id_col = next((i for i, h in enumerate(header) if '证件号码' in str(h) or h == '证件号码'), None)
        if id_col is None:
            id_col = next((i for i, h in enumerate(header) if h and '证件' in str(h)), 3)
        
        for row in ws.iter_rows(min_row=2, values_only=True):
            if len(row) > max(name_col, phone_col, id_col):
                name = str(row[name_col]).strip() if row[name_col] else ''
                phone = str(row[phone_col]).strip() if row[phone_col] else ''
                id_card = str(row[id_col]).strip() if row[id_col] else ''
                
                if id_card and id_card not in lookup:
                    lookup[id_card] = {'name': name, 'phone': phone}
    except Exception:
        pass
    
    return lookup


def has_value(value: Any) -> bool:
    if value is None:
        return False
    if isinstance(value, (int, float)):
        return True
    if isinstance(value, str):
        return value.strip() != ''
    return True


def detect_dynamic_and_unknown_columns(source_headers: List[Tuple[str, int]], 
                                        data_rows: List[List[Any]],
                                        column_config: Dict) -> Tuple[List[Tuple[str, str]], List[Tuple[str, str]]]:
    fixed_target_columns = set(column_config.get('fixed_columns', []))
    
    column_mapping = column_config.get('column_mapping', DEFAULT_COLUMN_MAPPING)
    source_to_target = {item['source']: item['target'] for item in column_mapping}
    known_sources = set(source_to_target.keys())
    
    fixed_target_set = set(column_config.get('fixed_columns', []))
    fixed_source_columns = set()
    for src, tgt in source_to_target.items():
        if tgt in fixed_target_set:
            fixed_source_columns.add(src)
            fixed_source_columns.add(tgt)
    
    dynamic_cols_with_data = []
    unknown_cols_with_data = []
    
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


def convert_single_file(input_path: str, output_dir: str, phone_lookup: Dict[str, str] = None, 
                       new_employees: Dict[str, Dict] = None) -> Tuple[bool, str, Dict]:
    if phone_lookup is None:
        phone_lookup = {}
    
    config = load_config()
    column_config = config.get('column_config', get_default_config()['column_config'])
    
    filename = os.path.basename(input_path)
    name_without_ext = os.path.splitext(filename)[0]
    
    counter = 1
    while True:
        if counter == 1:
            output_filename = f"{name_without_ext}_converted.xlsx"
        else:
            output_filename = f"{name_without_ext}_converted_{counter}.xlsx"
        
        output_path = os.path.join(output_dir, output_filename)
        
        if not os.path.exists(output_path):
            break
        counter += 1
    
    warnings = []
    try:
        meta, header_rows, data_rows = read_source_file(input_path)
        source_headers = extract_source_headers(header_rows)
        
        dynamic_cols, unknown_cols = detect_dynamic_and_unknown_columns(
            source_headers, data_rows, column_config
        )
        
        fixed_category = {
            '工号': '工号', '姓名': '姓名', '证件号码': '证件号码', '手机号': '手机号', '发薪月份': '发薪月份',
            '基本工资': '总计', '补发工资': '总计', '应发工资': '总计', '实发金额': '总计',
            '社保基数': '扣缴基数', '医保基数': '扣缴基数', '工伤基数': '扣缴基数', '公积金基数': '扣缴基数',
            '扣单位养老保险': '应扣项', '扣个人养老保险': '应扣项', '扣单位失业保险': '应扣项', '扣个人失业保险': '应扣项',
            '扣单位医疗保险': '应扣项', '扣个人医疗保险': '应扣项', '扣单位工伤险': '应扣项', '扣单位住房公积金': '应扣项',
            '扣个人住房公积金': '应扣项', '扣缴税金（扣个人所得税）': '应扣项', '2025年个税补退': '应扣项'
        }
        
        fixed_columns = column_config['fixed_columns']
        
        final_headers_1 = [fixed_category.get(col, col) for col in fixed_columns]
        final_headers_2 = fixed_columns.copy()
        
        for _, target in dynamic_cols:
            final_headers_2.append(target)
            final_headers_1.append(fixed_category.get(target, '其它'))
        
        output_unknown = column_config.get('output_unknown_columns', False)
        if output_unknown:
            for _, target in unknown_cols:
                final_headers_2.append(target)
                final_headers_1.append('其它')
        
        field_index = {name: idx for idx, name in enumerate(final_headers_2)}
        
        workbook = xlsxwriter.Workbook(output_path)
        worksheet = workbook.add_worksheet('Sheet1')
        
        worksheet.set_column('A:A', 12)
        worksheet.set_column('B:B', 10)
        worksheet.set_column('C:C', 22)
        worksheet.set_column('D:D', 12)
        worksheet.set_column('E:AC', 10)
        
        for col_idx, header in enumerate(final_headers_1):
            worksheet.write(0, col_idx, header)
        
        for col_idx, header in enumerate(final_headers_2):
            worksheet.write(1, col_idx, header)
        
        id_card_col_idx = None
        name_col_idx = None
        for col_name, col_idx in source_headers:
            if col_name in ('身份证', '证件号码'):
                id_card_col_idx = col_idx
            if col_name == '姓名':
                name_col_idx = col_idx
        
        current_id_cards = []
        
        fixed_col_set = set(column_config.get('fixed_columns', []))
        column_mapping = column_config.get('column_mapping', DEFAULT_COLUMN_MAPPING)
        source_to_target = {item['source']: item['target'] for item in column_mapping}
        output_unknown = column_config.get('output_unknown_columns', False)
        unknown_col_names = {src for src, tgt in unknown_cols} if output_unknown else set()
        
        for row_idx, data_row in enumerate(data_rows):
            target_row = [''] * len(final_headers_2)
            
            for source_name, target_name in FIELD_MAPPING.items():
                for col_name, col_idx in source_headers:
                    if col_name == source_name and col_idx < len(data_row):
                        target_idx = field_index.get(target_name)
                        if target_idx is not None:
                            target_row[target_idx] = data_row[col_idx]
            
            for col_name, col_idx in source_headers:
                if col_idx >= len(data_row):
                    continue
                value = data_row[col_idx]
                if not has_value(value):
                    continue
                
                if col_name in fixed_col_set:
                    continue
                
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
            intersection = set(current_id_cards) & set(new_employees.keys())
            
            if intersection:
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

class ConfigEditor(tk.Toplevel):
    def __init__(self, parent, config):
        super().__init__(parent)
        self.title("配置编辑器")
        self.geometry("800x600")
        self.resizable(True, True)
        self.transient(parent)
        self.grab_set()
        
        self.config = config.copy()
        self.result = None
        self.mapping_entries = []
        
        column_config = self.config.get('column_config', get_default_config()['column_config'])
        self.fixed_columns = column_config.get('fixed_columns', []).copy()
        self.column_mapping = column_config.get('column_mapping', DEFAULT_COLUMN_MAPPING).copy()
        self.output_unknown_columns = column_config.get('output_unknown_columns', False)
        
        self.setup_ui()
    
    def setup_ui(self):
        notebook = ttk.Notebook(self)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.tab_phone = ttk.Frame(notebook)
        self.tab_source = ttk.Frame(notebook)
        self.tab_fixed = ttk.Frame(notebook)
        self.tab_mapping = ttk.Frame(notebook)
        
        notebook.add(self.tab_phone, text="电话查询配置")
        notebook.add(self.tab_source, text="源表配置")
        notebook.add(self.tab_fixed, text="固定列")
        notebook.add(self.tab_mapping, text="字段映射")
        
        self.setup_phone_tab()
        self.setup_source_tab()
        self.setup_fixed_tab()
        self.setup_mapping_tab()
        
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(btn_frame, text="重置为默认", command=self.reset_to_default).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=self.cancel).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_frame, text="确定", command=self.confirm).pack(side=tk.RIGHT, padx=5)
    
    def setup_mapping_tab(self):
        frame = ttk.Frame(self.tab_mapping, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        
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
        for item in self.mapping_tree.get_children():
            self.mapping_tree.delete(item)
        for item in self.column_mapping:
            self.mapping_tree.insert('', tk.END, values=(item['source'], item['target']))
    
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
        selected = self.mapping_tree.selection()
        if selected:
            item = selected[0]
            values = self.mapping_tree.item(item, 'values')
            self.column_mapping.remove({'source': values[0], 'target': values[1]})
            self.refresh_mapping_list()
    
    def _on_frame_configure(self, event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox('all'))
    
    def _on_canvas_configure(self, event):
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
        frame = ttk.Frame(self.tab_phone, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)
        
        phone_config = self.config.get('phone_lookup_config', {})
        csv_config = phone_config.get('csv', {})
        xlsx_config = phone_config.get('xlsx', {})
        
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
    
    def setup_fixed_tab(self):
        frame = ttk.Frame(self.tab_fixed, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        
        top_frame = ttk.Frame(frame)
        top_frame.pack(fill=tk.X, pady=(0, 5))
        
        fixed_count = len(self.fixed_columns)
        ttk.Label(top_frame, text=f"固定列（{fixed_count}项，按顺序输出，始终显示）").pack(side=tk.LEFT)
        
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
        
        dialog.bind('<Return>', lambda e: save())
        dialog.bind('<Escape>', lambda e: dialog.destroy())
    
    def refresh_fixed_list(self):
        self.fixed_listbox.delete(0, tk.END)
        for col in self.fixed_columns:
            self.fixed_listbox.insert(tk.END, col)
    
    def fixed_move_up(self):
        selected = self.fixed_listbox.curselection()
        if selected and selected[0] > 0:
            idx = selected[0]
            self.fixed_columns[idx-1], self.fixed_columns[idx] = self.fixed_columns[idx], self.fixed_columns[idx-1]
            self.refresh_fixed_list()
            self.fixed_listbox.selection_set(idx-1)
    
    def fixed_move_down(self):
        selected = self.fixed_listbox.curselection()
        if selected and selected[0] < len(self.fixed_columns) - 1:
            idx = selected[0]
            self.fixed_columns[idx], self.fixed_columns[idx+1] = self.fixed_columns[idx+1], self.fixed_columns[idx]
            self.refresh_fixed_list()
            self.fixed_listbox.selection_set(idx+1)
    
    def fixed_delete(self):
        selected = self.fixed_listbox.curselection()
        if selected:
            idx = selected[0]
            del self.fixed_columns[idx]
            self.refresh_fixed_list()
    
    def reset_to_default(self):
        self.config = get_default_config()
        
        self.fixed_columns = self.config.get('column_config', {}).get('fixed_columns', []).copy()
        self.column_mapping = self.config.get('column_config', {}).get('column_mapping', DEFAULT_COLUMN_MAPPING).copy()
        
        self.refresh_fixed_list()
        self.refresh_mapping_list()
        
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
        self.result = None
        self.destroy()
    
    def confirm(self):
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
        self.config['source_config'] = {
            'title_row': int(getattr(self, 'title_row_entry', None).get().strip() or 0) if hasattr(self, 'title_row_entry') else 0,
            'title_col': int(getattr(self, 'title_col_entry', None).get().strip() or 2) if hasattr(self, 'title_col_entry') else 2,
            'header_row_count': int(getattr(self, 'header_count_entry', None).get().strip() or 5) if hasattr(self, 'header_count_entry') else 5,
            'data_start_row': int(getattr(self, 'data_start_entry', None).get().strip() or 5) if hasattr(self, 'data_start_entry') else 5,
            'skip_keyword': getattr(self, 'skip_keyword_entry', None).get().strip() or '合计' if hasattr(self, 'skip_keyword_entry') else '合计'
        }
        
        self.config['column_config'] = {
            'output_unknown_columns': getattr(self, 'output_unknown_var', None).get() if hasattr(self, 'output_unknown_var') else False,
            'fixed_columns': self.fixed_columns,
            'column_mapping': self.column_mapping
        }
        
        self.result = self.config
        self.destroy()


class ExcelConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("东软-薪福通工资表格式转换工具")
        self.root.geometry("900x600")
        self.root.resizable(True, True)
        
        self.file_list = []
        self.selected_index = -1
        self.output_dir = tk.StringVar()
        self.export_csv_path = tk.StringVar()
        self.csv_create_time = tk.StringVar()
        
        self.app_config = load_config_and_apply()
        
        self.setup_ui()
        self.setup_drag_drop()
        self.load_recent_csv()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        title_label = ttk.Label(main_frame, text="东软-薪福通工资表格式转换工具", font=('song', 16, 'bold'))
        title_label.pack(pady=(0, 10))
        
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        left_frame = ttk.LabelFrame(content_frame, text="文件列表", padding="5")
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        list_frame = ttk.Frame(left_frame)
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.file_listbox = tk.Listbox(list_frame, selectmode=tk.EXTENDED, 
                                        yscrollcommand=scrollbar.set,
                                        font=('song', 10))
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.file_listbox.yview)
        
        self.file_listbox.bind('<<ListboxSelect>>', self.on_file_select)
        
        btn_frame = ttk.Frame(left_frame)
        btn_frame.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Button(btn_frame, text="添加文件", command=self.add_files).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="移除选中", command=self.remove_selected).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="清空列表", command=self.clear_list).pack(side=tk.LEFT, padx=2)
        
        drop_hint = ttk.Label(left_frame, text="(或拖放文件到此处)", font=('song', 9), foreground='gray')
        drop_hint.pack(pady=(5, 0))
        
        right_frame = ttk.LabelFrame(content_frame, text="文件预览", padding="10")
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))
        
        info_frame = ttk.Frame(right_frame)
        info_frame.pack(fill=tk.BOTH, expand=True)
        
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
        
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Label(bottom_frame, text="输出目录:").pack(side=tk.LEFT)
        
        output_entry = ttk.Entry(bottom_frame, textvariable=self.output_dir, width=50)
        output_entry.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(bottom_frame, text="浏览...", command=self.browse_output_dir).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(bottom_frame, text="编辑配置", command=self.open_config_editor).pack(side=tk.LEFT, padx=20)
        
        self.convert_btn = ttk.Button(bottom_frame, text="开始转换", command=self.start_convert)
        self.convert_btn.pack(side=tk.RIGHT)
        
        phone_frame = ttk.Frame(main_frame)
        phone_frame.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Label(phone_frame, text="人员信息:").pack(side=tk.LEFT)
        
        self.csv_drop_frame = tk.Frame(phone_frame, bd=2, relief=tk.SUNKEN, bg='white', height=30)
        self.csv_drop_frame.pack(side=tk.LEFT, padx=5)
        self.csv_drop_frame.pack_propagate(False)
        self.csv_drop_frame.configure(width=350)
        
        self.csv_drop_label = tk.Label(self.csv_drop_frame, text="拖拽 CSV/员工模板文件到此处 或 点击选择", 
                                        bg='white', fg='gray', font=('song', 9), anchor='w')
        self.csv_drop_label.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.csv_drop_frame.drop_target_register(DND_FILES)
        self.csv_drop_frame.dnd_bind('<<Drop>>', self.on_csv_drop)
        self.csv_drop_frame.bind('<Button-1>', lambda e: self.on_csv_click())
        
        ttk.Button(phone_frame, text="浏览...", command=self.on_csv_click).pack(side=tk.LEFT, padx=5)
        
        csv_info_frame = ttk.Frame(main_frame)
        csv_info_frame.pack(fill=tk.X, pady=(2, 0))
        
        self.csv_info_label = ttk.Label(csv_info_frame, text="", foreground='gray', font=('song', 8))
        self.csv_info_label.pack(side=tk.LEFT)
        
        self.status_label = ttk.Label(main_frame, text="就绪", foreground='green')
        self.status_label.pack(pady=(5, 0))
        
        log_frame = ttk.LabelFrame(main_frame, text="执行日志", padding="5")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        
        log_scroll_y = ttk.Scrollbar(log_frame)
        log_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.log_text = tk.Text(log_frame, height=8, font=('Consolas', 9), 
                                yscrollcommand=log_scroll_y.set, wrap=tk.WORD)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scroll_y.config(command=self.log_text.yview)
    
    def setup_drag_drop(self):
        self.file_listbox.drop_target_register(DND_FILES)
        self.file_listbox.dnd_bind('<<Drop>>', self.on_excel_drop)
    
    def on_excel_drop(self, event):
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
        filepath = filedialog.askopenfilename(
            title="选择人员信息",
            filetypes=[("支持的文件", "*.csv;*.xlsx"), ("CSV 文件", "*.csv"), 
                       ("Excel 文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if filepath:
            self.set_csv_path(filepath)
    
    def set_csv_path(self, filepath):
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
            self.csv_info_label.config(text="")
    
    def save_recent_csv(self):
        csv_path = self.export_csv_path.get()
        if csv_path and os.path.exists(csv_path):
            create_time = get_file_create_time(csv_path)
            self.app_config['recent_csv'] = {
                'path': csv_path,
                'created_time': create_time or ''
            }
            save_config(self.app_config)
    
    def open_config_editor(self):
        editor = ConfigEditor(self.root, self.app_config)
        self.root.wait_window(editor)
        if editor.result:
            self.app_config = editor.result
            save_config(self.app_config)
            load_config_and_apply()
            messagebox.showinfo("提示", "配置已保存")
    
    def add_files(self):
        files = filedialog.askopenfilenames(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xls *.xlsx"), ("所有文件", "*.*")]
        )
        for f in files:
            if f not in [item['path'] for item in self.file_list]:
                self.add_file_to_list(f)
    
    def add_file_to_list(self, filepath):
        info = get_file_info(filepath)
        if info:
            self.file_list.append(info)
            display_text = f"{info['filename']} ({info['year_month']}) - {info['data_rows']}行"
            self.file_listbox.insert(tk.END, display_text)
    
    def remove_selected(self):
        selected = self.file_listbox.curselection()
        for i in reversed(selected):
            self.file_listbox.delete(i)
            del self.file_list[i]
        self.clear_preview()
    
    def clear_list(self):
        self.file_listbox.delete(0, tk.END)
        self.file_list = []
        self.clear_preview()
    
    def clear_preview(self):
        self.label_filename.config(text="-")
        self.label_title.config(text="-")
        self.label_yearmonth.config(text="-")
        self.label_rows.config(text="-")
        self.label_size.config(text="-")
        self.selected_index = -1
    
    def on_file_select(self, event):
        selection = self.file_listbox.curselection()
        if selection:
            self.selected_index = selection[0]
            info = self.file_list[self.selected_index]
            
            self.label_filename.config(text=info['filename'])
            self.label_title.config(text=info['title'] if info['title'] else "(无标题)")
            self.label_yearmonth.config(text=info['year_month'] if info['year_month'] else "(无法提取)")
            self.label_rows.config(text=str(info['data_rows']))
            
            size = info['file_size']
            if size < 1024:
                size_str = f"{size} B"
            elif size < 1024 * 1024:
                size_str = f"{size / 1024:.1f} KB"
            else:
                size_str = f"{size / (1024 * 1024):.1f} MB"
            self.label_size.config(text=size_str)
    
    def browse_output_dir(self):
        dirpath = filedialog.askdirectory(title="选择输出目录")
        if dirpath:
            self.output_dir.set(dirpath)
    
    def start_convert(self):
        if not self.file_list:
            messagebox.showwarning("提示", "请先添加要转换的文件")
            return
        
        output_dir = self.output_dir.get()
        if not output_dir:
            if self.file_list:
                first_file = self.file_list[0]['path']
                default_dir = os.path.join(os.path.dirname(first_file), 'Converted')
                self.output_dir.set(default_dir)
                output_dir = default_dir
        
        os.makedirs(output_dir, exist_ok=True)
        
        log_dir = os.path.join(output_dir, 'log')
        os.makedirs(log_dir, exist_ok=True)
        log_file = os.path.join(log_dir, f"convert_{datetime.datetime.now().strftime('%Y%m%d')}.log")
        
        export_csv = self.export_csv_path.get()
        phone_lookup = load_phone_lookup(export_csv)
        
        new_employees = None
        if export_csv and os.path.exists(export_csv):
            if export_csv.endswith('.xlsx'):
                new_employees = load_new_employees_from_template(export_csv)
                if new_employees:
                    self.log(f"已加载新增人员数据: {len(new_employees)} 人", "INFO")
        
        self.log_text.delete(1.0, tk.END)
        self.log(f"开始转换，共 {len(self.file_list)} 个文件", "INFO")
        self.log(f"输出目录: {output_dir}", "INFO")
        self.log(f"日志文件: {log_file}", "INFO")
        
        self.convert_btn.config(state=tk.DISABLED)
        success_count = 0
        fail_count = 0
        results = []
        
        for i, file_info in enumerate(self.file_list):
            self.status_label.config(text=f"正在转换 ({i+1}/{len(self.file_list)}): {file_info['filename']}")
            self.root.update()
            
            self.log(f"开始处理: {file_info['filename']}", "INFO")
            
            success, msg, detail = convert_single_file(file_info['path'], output_dir, phone_lookup, new_employees)
            
            if success:
                success_count += 1
                if detail.get('warnings'):
                    for w in detail['warnings']:
                        self.log(f"  ⚠️ 第 {w['row']} 行: 缺失 [{', '.join(w['missing_fields'])}]", "WARNING")
                self.log(f"完成: {msg}", "INFO")
            else:
                fail_count += 1
                self.log(f"失败: {msg}", "ERROR")
            
            results.append((file_info['filename'], success, msg))
        
        log_content = self.get_log_content()
        try:
            with open(log_file, 'a', encoding='utf-8') as f:
                f.write(log_content)
        except Exception as e:
            self.log(f"保存日志文件失败: {e}", "ERROR")
        
        self.convert_btn.config(state=tk.NORMAL)
        
        self.log("-" * 40, "INFO")
        self.log(f"汇总: 成功 {success_count} 个, 失败 {fail_count} 个", "INFO")
        self.log(f"日志已保存: {log_file}", "INFO")
        
        result_msg = f"转换完成！\n\n成功: {success_count} 个文件\n失败: {fail_count} 个文件\n\n日志已保存至:\n{log_file}"
        if fail_count > 0:
            result_msg += "\n\n失败文件:"
            for fname, success, msg in results:
                if not success:
                    result_msg += f"\n- {fname}: {msg}"
        
        self.status_label.config(text=f"完成 - 成功: {success_count}, 失败: {fail_count}")
        messagebox.showinfo("转换结果", result_msg)
    
    def on_closing(self):
        self.save_recent_csv()
        self.root.destroy()
    
    def log(self, message: str, level: str = "INFO"):
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        formatted = f"[{timestamp}] {level} {message}\n"
        
        self.log_text.insert(tk.END, formatted)
        
        if level == "WARNING":
            self.log_text.tag_add("warning", self.log_text.index(tk.END + "-2l"), tk.END)
            self.log_text.tag_config("warning", foreground="orange")
        elif level == "ERROR":
            self.log_text.tag_add("error", self.log_text.index(tk.END + "-2l"), tk.END)
            self.log_text.tag_config("error", foreground="red")
        
        self.log_text.see(tk.END)
    
    def get_log_content(self) -> str:
        return self.log_text.get(1.0, tk.END)


def main():
    root = tkinterdnd2.Tk()
    app = ExcelConverterApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()
