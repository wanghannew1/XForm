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


# ==================== 配置部分 ====================

CONFIG_FILE = 'dongr-xinft_config.json'

DEFAULT_FIELD_MAPPING = {
    '职工号': '工号',
    '姓名': '姓名',
    '身份证': '证件号码',
    '基本工资': '基本工资',
    '应发工资': '应发工资',
    '实发合计': '实发金额',
    '社保基数': '社保基数',
    '医保基数': '医保基数',
    '工伤基数': '工伤基数',
    '公积金基数': '公积金基数',
    '扣款明细_养老_单位': '扣单位养老保险',
    '扣款明细_养老_个人': '扣个人养老保险',
    '扣款明细_失业_单位': '扣单位失业保险',
    '扣款明细_失业_个人': '扣个人失业保险',
    '扣款明细_医疗_单位': '扣单位医疗保险',
    '扣款明细_医疗_个人': '扣个人医疗保险',
    '扣款明细_工伤险_单位': '扣单位工伤险',
    '扣款明细_公积金_单位': '扣单位住房公积金',
    '扣款明细_公积金_个人': '扣个人住房公积金',
    '个人所得税': '扣缴税金（扣个人所得税）',
    '补发': '补发工资',
    '扣款明细_补缴及退款保险差额（单位）': '补扣单位保险',
    '扣款明细_补缴及退款保险差额（个人）': '补扣个人保险',
    '扣款明细_大病险合计': '医保大额',
}

TARGET_HEADER_ROW1 = [
    '工号', '姓名', '证件号码', '手机号', '发薪月份',
    '总计', '总计', '总计', '总计',
    '扣缴基数', '扣缴基数', '扣缴基数', '扣缴基数',
    '应扣项', '应扣项', '应扣项', '应扣项', '应扣项', '应扣项',
    '应扣项', '应扣项', '应扣项', '应扣项', '应扣项', '应扣项',
    '应扣项', '应扣项', '应扣项', '其它'
]

TARGET_HEADER_ROW2 = [
    '工号', '姓名', '证件号码', '手机号', '发薪月份',
    '基本工资', '补发工资', '应发工资', '实发金额',
    '社保基数', '医保基数', '工伤基数', '公积金基数',
    '扣单位养老保险', '扣个人养老保险', '扣单位失业保险', '扣个人失业保险',
    '扣单位医疗保险', '扣个人医疗保险', '扣单位工伤险', '扣单位住房公积金',
    '扣个人住房公积金', '扣单位公务员医疗补助', '医保大额', '补扣单位保险',
    '补扣个人保险', '扣缴税金（扣个人所得税）', '2025年个税补退', '采暖费'
]

TARGET_FIELD_INDEX = {name: idx for idx, name in enumerate(TARGET_HEADER_ROW2)}

FIELD_MAPPING = DEFAULT_FIELD_MAPPING.copy()


# ==================== 配置加载/保存 ====================

def get_default_config() -> Dict:
    return {
        "source_field_mapping": DEFAULT_FIELD_MAPPING.copy(),
        "phone_lookup_config": {
            "id_card_column": "公民身份号码",
            "phone_column": "联系电话"
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
    FIELD_MAPPING = config.get('source_field_mapping', DEFAULT_FIELD_MAPPING.copy())
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
    
    import openpyxl
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


def load_phone_lookup(csv_path: str, id_col_name: str = None, phone_col_name: str = None) -> Dict[str, str]:
    global FIELD_MAPPING
    lookup = {}
    if not csv_path or not os.path.exists(csv_path):
        return lookup
    config = load_config()
    phone_config = config.get('phone_lookup_config', {})
    if id_col_name is None:
        id_col_name = phone_config.get('id_card_column', '公民身份号码')
    if phone_col_name is None:
        phone_col_name = phone_config.get('phone_column', '联系电话')
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
                    id_card = row[id_col].strip()
                    phone = row[phone_col].strip()
                    if id_card and phone:
                        lookup[id_card] = phone
    except Exception:
        pass
    return lookup


def convert_single_file(input_path: str, output_dir: str, phone_lookup: Dict[str, str] = None) -> Tuple[bool, str]:
    if phone_lookup is None:
        phone_lookup = {}
    
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
    
    try:
        meta, header_rows, data_rows = read_source_file(input_path)
        source_headers = extract_source_headers(header_rows)
        
        workbook = xlsxwriter.Workbook(output_path)
        worksheet = workbook.add_worksheet('Sheet1')
        
        for col_idx, header in enumerate(TARGET_HEADER_ROW1):
            worksheet.write(0, col_idx, header)
        
        for col_idx, header in enumerate(TARGET_HEADER_ROW2):
            worksheet.write(1, col_idx, header)
        
        id_card_col_idx = None
        for col_name, col_idx in source_headers:
            if col_name in ('身份证', '证件号码'):
                id_card_col_idx = col_idx
                break
        
        for row_idx, data_row in enumerate(data_rows):
            target_row = map_row_to_target(data_row, source_headers)
            
            for source_name, target_name in FIELD_MAPPING.items():
                for col_name, col_idx in source_headers:
                    if col_name == source_name and col_idx < len(data_row):
                        target_idx = TARGET_FIELD_INDEX.get(target_name)
                        if target_idx is not None:
                            target_row[target_idx] = data_row[col_idx]
            
            id_card = ''
            if id_card_col_idx is not None and id_card_col_idx < len(data_row):
                id_card = str(data_row[id_card_col_idx]).strip()
            phone = phone_lookup.get(id_card, '')
            target_row[3] = phone
            target_row[4] = meta['year_month']
            
            for col_idx, value in enumerate(target_row):
                if value:
                    worksheet.write(row_idx + 2, col_idx, value)
        
        workbook.close()
        return True, f"成功转换 {len(data_rows)} 行数据"
    
    except Exception as e:
        return False, f"转换失败: {str(e)}"


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
        self.geometry("800x550")
        self.resizable(True, True)
        self.transient(parent)
        self.grab_set()
        
        self.config = config.copy()
        self.result = None
        self.mapping_entries = []
        
        self.setup_ui()
    
    def setup_ui(self):
        notebook = ttk.Notebook(self)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.tab_mapping = ttk.Frame(notebook)
        self.tab_phone = ttk.Frame(notebook)
        self.tab_source = ttk.Frame(notebook)
        
        notebook.add(self.tab_mapping, text="字段映射")
        notebook.add(self.tab_phone, text="电话查询配置")
        notebook.add(self.tab_source, text="源表配置")
        
        self.setup_mapping_tab()
        self.setup_phone_tab()
        self.setup_source_tab()
        
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(btn_frame, text="重置为默认", command=self.reset_to_default).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=self.cancel).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_frame, text="确定", command=self.confirm).pack(side=tk.RIGHT, padx=5)
    
    def setup_mapping_tab(self):
        frame = ttk.Frame(self.tab_mapping, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        
        header_frame = ttk.Frame(frame)
        header_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(header_frame, text="源列名（源表）", width=30).pack(side=tk.LEFT, padx=5)
        ttk.Label(header_frame, text="目标列名（输出表）", width=30).pack(side=tk.LEFT, padx=5)
        
        canvas_frame = ttk.Frame(frame)
        canvas_frame.pack(fill=tk.BOTH, expand=True)
        
        self.canvas = tk.Canvas(canvas_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.mapping_frame = ttk.Frame(self.canvas)
        
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.mapping_frame_id = self.canvas.create_window((0, 0), window=self.mapping_frame, anchor='nw')
        self.mapping_frame.bind('<Configure>', self._on_frame_configure)
        self.canvas.bind('<Configure>', self._on_canvas_configure)
        
        for src, tgt in self.config.get('source_field_mapping', {}).items():
            self.add_mapping_row(src, tgt)
        
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(btn_frame, text="添加映射", command=self.add_mapping).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="删除选中", command=self.delete_mapping).pack(side=tk.LEFT, padx=2)
    
    def _on_frame_configure(self, event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox('all'))
    
    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self.mapping_frame_id, width=event.width)
    
    def add_mapping_row(self, src='', tgt=''):
        row_frame = ttk.Frame(self.mapping_frame)
        row_frame.pack(fill=tk.X, pady=2)
        
        src_entry = ttk.Entry(row_frame, width=30)
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
        
        ttk.Label(frame, text="身份证列名:").grid(row=0, column=0, sticky=tk.W, pady=15)
        self.id_card_entry = ttk.Entry(frame, width=30)
        self.id_card_entry.grid(row=0, column=1, sticky=tk.W, padx=10)
        self.id_card_entry.insert(0, phone_config.get('id_card_column', '公民身份号码'))
        ttk.Label(frame, text="← 员工信息表（CSV）中身份证号所在的列名", foreground='gray').grid(row=0, column=2, sticky=tk.W)
        
        ttk.Label(frame, text="联系电话列名:").grid(row=1, column=0, sticky=tk.W, pady=15)
        self.phone_entry = ttk.Entry(frame, width=30)
        self.phone_entry.grid(row=1, column=1, sticky=tk.W, padx=10)
        self.phone_entry.insert(0, phone_config.get('phone_column', '联系电话'))
        ttk.Label(frame, text="← 员工信息表（CSV）中手机号所在的列名", foreground='gray').grid(row=1, column=2, sticky=tk.W)
    
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
    
    def reset_to_default(self):
        self.config = get_default_config()
        for var, src_entry, tgt_entry in self.mapping_entries:
            src_entry.master.destroy()
        self.mapping_entries.clear()
        for src, tgt in DEFAULT_FIELD_MAPPING.items():
            self.add_mapping_row(src, tgt)
        phone_config = self.config.get('phone_lookup_config', {})
        self.id_card_entry.delete(0, tk.END)
        self.id_card_entry.insert(0, phone_config.get('id_card_column', '公民身份号码'))
        self.phone_entry.delete(0, tk.END)
        self.phone_entry.insert(0, phone_config.get('phone_column', '联系电话'))
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
        field_mapping = {}
        for var, src_entry, tgt_entry in self.mapping_entries:
            src = src_entry.get().strip()
            tgt = tgt_entry.get().strip()
            if src and tgt:
                field_mapping[src] = tgt
        
        self.config['source_field_mapping'] = field_mapping
        self.config['phone_lookup_config'] = {
            'id_card_column': self.id_card_entry.get().strip(),
            'phone_column': self.phone_entry.get().strip()
        }
        self.config['source_config'] = {
            'title_row': int(self.title_row_entry.get().strip() or 0),
            'title_col': int(self.title_col_entry.get().strip() or 2),
            'header_row_count': int(self.header_count_entry.get().strip() or 5),
            'data_start_row': int(self.data_start_entry.get().strip() or 5),
            'skip_keyword': self.skip_keyword_entry.get().strip() or '合计'
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
        
        ttk.Label(phone_frame, text="联系电话查询表:").pack(side=tk.LEFT)
        
        self.csv_drop_frame = tk.Frame(phone_frame, bd=2, relief=tk.SUNKEN, bg='white', height=30)
        self.csv_drop_frame.pack(side=tk.LEFT, padx=5)
        self.csv_drop_frame.pack_propagate(False)
        self.csv_drop_frame.configure(width=350)
        
        self.csv_drop_label = tk.Label(self.csv_drop_frame, text="拖拽 CSV 文件到此处 或 点击选择", 
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
            if file_path and file_path.endswith('.csv'):
                self.set_csv_path(file_path)
                break
        return 'action'
    
    def on_csv_click(self):
        filepath = filedialog.askopenfilename(
            title="选择联系电话查询表",
            filetypes=[("CSV 文件", "*.csv"), ("所有文件", "*.*")]
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
            global FIELD_MAPPING
            FIELD_MAPPING = self.app_config.get('source_field_mapping', DEFAULT_FIELD_MAPPING.copy())
            save_config(self.app_config)
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
        
        export_csv = self.export_csv_path.get()
        phone_lookup = load_phone_lookup(export_csv)
        
        self.convert_btn.config(state=tk.DISABLED)
        success_count = 0
        fail_count = 0
        results = []
        
        for i, file_info in enumerate(self.file_list):
            self.status_label.config(text=f"正在转换 ({i+1}/{len(self.file_list)}): {file_info['filename']}")
            self.root.update()
            
            success, msg = convert_single_file(file_info['path'], output_dir, phone_lookup)
            if success:
                success_count += 1
            else:
                fail_count += 1
            results.append((file_info['filename'], success, msg))
        
        self.convert_btn.config(state=tk.NORMAL)
        
        result_msg = f"转换完成！\n\n成功: {success_count} 个文件\n失败: {fail_count} 个文件"
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


def main():
    root = tkinterdnd2.Tk()
    app = ExcelConverterApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()
