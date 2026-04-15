#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
员工信息格式转换工具
将东软导出的 CSV 转换为薪福通 员工导入模板202512.xlsx 格式
支持命令行模式和GUI模式
"""

import os
import sys
import re
import csv
import json
import shutil
import argparse
from typing import Dict, List, Optional, Tuple, Set
from datetime import datetime

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    import tkinterdnd2
    from tkinterdnd2 import DND_FILES
    HAS_TKINTER = True
except ImportError:
    HAS_TKINTER = False

try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False


SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(SCRIPT_DIR, 'employee_info_converter_config.json')
TEMPLATE_FILE = os.path.join(SCRIPT_DIR, '员工导入模板202512.xlsx')
OUTPUT_DIR = os.path.join(SCRIPT_DIR, 'Cleaned')

YELLOW_FILL = PatternFill(start_color='FFFFFF00', end_color='FFFFFF00', fill_type='solid')
GREEN_FILL = PatternFill(start_color='FF00FF00', end_color='FF00FF00', fill_type='solid')
ABNORMAL_SHEET_NAME = '异常数据'
NEW_EMPLOYEES_SHEET_NAME = '新增人员'


def find_latest_cleaned_file(exclude_file: str = None) -> str:
    """查找最近一次生成的转换文件（排除当前文件）"""
    if not os.path.exists(OUTPUT_DIR):
        return None
    
    files = [f for f in os.listdir(OUTPUT_DIR) if f.startswith('员工导入模板_转换后_') and f.endswith('.xlsx')]
    if not files:
        return None
    
    files_with_time = []
    for f in files:
        full_path = os.path.join(OUTPUT_DIR, f)
        if exclude_file and os.path.abspath(full_path) == os.path.abspath(exclude_file):
            continue
        mtime = os.path.getmtime(full_path)
        files_with_time.append((mtime, full_path))
    
    if not files_with_time:
        return None
    
    files_with_time.sort(reverse=True)
    return files_with_time[0][1]


def extract_employees_from_xlsx(file_path: str, config: Dict) -> Dict[str, Dict]:
    """从花名册提取员工数据，返回 {证件号码: {name, phone, row}}"""
    if not os.path.exists(file_path):
        return {}
    
    employees = {}
    try:
        wb = openpyxl.load_workbook(file_path)
        sheet_name = config.get('sheet_name', '花名册')
        if sheet_name not in wb.sheetnames:
            return {}
        
        ws = wb[sheet_name]
        for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if not row or not row[0]:
                continue
            name = str(row[0]).strip() if row[0] else ''
            id_card = str(row[3]).strip() if len(row) > 3 and row[3] else ''
            phone = str(row[1]).strip() if len(row) > 1 and row[1] else ''
            
            if id_card:
                employees[id_card] = {
                    'name': name,
                    'phone': phone,
                    'row': row_idx,
                    'unit': str(row[4]).strip() if len(row) > 4 else '',
                    'settlement': str(row[5]).strip() if len(row) > 5 else '',
                    'employment_type': str(row[6]).strip() if len(row) > 6 else ''
                }
    except Exception:
        pass
    
    return employees


def find_new_employees_from_records(records: List[Dict], prev_file: str, config: Dict) -> Tuple[List[Dict], str]:
    """比对当前记录和上次的文件，返回新增人员列表和参考文件路径"""
    if not prev_file or not os.path.exists(prev_file):
        return [], ''
    
    prev_employees = extract_employees_from_xlsx(prev_file, config)
    
    current_employees = {}
    for rec in records:
        id_card = rec.get('id_card', '')
        if id_card and id_card not in current_employees:
            current_employees[id_card] = {
                'name': rec.get('name', ''),
                'phone': rec.get('phone', ''),
                'unit': rec.get('unit', ''),
                'settlement': rec.get('settlement', ''),
                'employment_type': rec.get('employment_type', '')
            }
    
    new_list = []
    for id_card, emp in current_employees.items():
        if id_card not in prev_employees:
            new_list.append({
                'name': emp['name'],
                'phone': emp['phone'],
                'id_card': id_card,
                'unit': emp.get('unit', ''),
                'settlement': emp.get('settlement', ''),
                'employment_type': emp.get('employment_type', '')
            })
    
    return new_list, prev_file


def find_new_employees(current_file: str, prev_file: str, config: Dict) -> Tuple[List[Dict], str]:
    """比对当前和上次的文件，返回新增人员列表和参考文件路径"""
    if not prev_file or not os.path.exists(prev_file):
        return [], ''
    
    prev_employees = extract_employees_from_xlsx(prev_file, config)
    current_employees = extract_employees_from_xlsx(current_file, config)
    
    new_list = []
    for id_card, emp in current_employees.items():
        if id_card not in prev_employees:
            new_list.append({
                'name': emp['name'],
                'phone': emp['phone'],
                'id_card': id_card,
                'unit': emp.get('unit', ''),
                'settlement': emp.get('settlement', ''),
                'employment_type': emp.get('employment_type', '')
            })
    
    return new_list, prev_file


def create_new_employees_sheet(output_file: str, new_employees: List[Dict], reference_file: str) -> int:
    """在输出文件中创建新增人员Sheet，返回新增人数"""
    if not new_employees:
        return 0
    
    wb = openpyxl.load_workbook(output_file)
    
    if NEW_EMPLOYEES_SHEET_NAME in wb.sheetnames:
        del wb[NEW_EMPLOYEES_SHEET_NAME]
    
    ws = wb.create_sheet(NEW_EMPLOYEES_SHEET_NAME)
    
    headers = ['*姓名', '*手机号', '*证件类型', '*证件号码', '单位名称', '结算单元名称', '用工性质']
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')
    
    for i, emp in enumerate(new_employees, start=2):
        ws.cell(row=i, column=1, value=emp['name'])
        ws.cell(row=i, column=2, value=emp['phone'])
        ws.cell(row=i, column=3, value='居民身份证')
        ws.cell(row=i, column=4, value=emp['id_card'])
        ws.cell(row=i, column=5, value=emp.get('unit', ''))
        ws.cell(row=i, column=6, value=emp.get('settlement', ''))
        ws.cell(row=i, column=7, value=emp.get('employment_type', ''))
        
        for col in range(1, 8):
            cell = ws.cell(row=i, column=col)
            cell.fill = GREEN_FILL
            cell.alignment = Alignment(horizontal='center')
    
    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 13
    ws.column_dimensions['C'].width = 13
    ws.column_dimensions['D'].width = 20
    ws.column_dimensions['E'].width = 20
    ws.column_dimensions['F'].width = 20
    ws.column_dimensions['G'].width = 12
    ws.row_dimensions[1].height = 15
    
    wb.save(output_file)
    return len(new_employees)


def get_default_config() -> Dict:
    return {
        "csv_encoding": "gb18030",
        "id_card_col": 0,
        "name_col": 1,
        "phone_col": 50,
        "date_col": 59,
        "sheet_name": "花名册",
        "recent_csv": "",
        "recent_output_filename": "员工导入模板_转换后"
    }


def load_config() -> Dict:
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return get_default_config()


def save_config(config: Dict) -> None:
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"保存配置失败: {e}")


def get_output_path(filename_prefix: str = '员工导入模板_转换后') -> str:
    """生成输出文件路径，自动添加时间戳，重名时加顺序号"""
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    base_name = f'{filename_prefix}_{timestamp}.xlsx'
    output_path = os.path.join(OUTPUT_DIR, base_name)
    
    counter = 1
    while os.path.exists(output_path):
        base_name = f'{filename_prefix}_{timestamp}_{counter}.xlsx'
        output_path = os.path.join(OUTPUT_DIR, base_name)
        counter += 1
    
    return output_path


def validate_id_card(id_card: str) -> Tuple[bool, str]:
    """验证身份证号有效性
    
    返回: (是否有效, 错误信息)
    """
    if not id_card:
        return False, "身份证号为空"
    
    if len(id_card) != 18:
        return False, f"长度不足18位(当前{len(id_card)}位)"
    
    if not id_card[:17].isdigit():
        return False, "前17位必须为数字"
    
    if id_card[17] not in '0123456789Xx':
        return False, "第18位必须是数字0-9或X"
    
    birth_date = id_card[6:14]
    try:
        datetime.strptime(birth_date, '%Y%m%d')
    except ValueError:
        return False, f"出生日期无效: {birth_date}"
    
    Wi = [7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2]
    CheckCode = ['1', '0', 'X', '9', '8', '7', '6', '5', '4', '3', '2']
    
    sum_val = sum(int(id_card[i]) * Wi[i] for i in range(17))
    expected = CheckCode[sum_val % 11]
    
    if id_card[17].upper() != expected:
        return False, f"校验码错误(实际:{id_card[17]}, 应为:{expected})"
    
    return True, ""


def get_id_card_invalid_reason(phone: str) -> str:
    """获取手机号无效原因"""
    if not phone:
        return "手机号为空"
    
    if len(phone) != 11:
        return f"长度不是11位(当前{len(phone)}位)"
    
    if phone[0] == '0':
        return "以0开头(可能是固话)"
    
    if not phone.isdigit():
        return "包含非数字字符"
    
    if len(set(phone)) == 1:
        return "所有数字相同"
    
    for i in range(len(phone) - 7):
        if len(set(phone[i:i+8])) == 1:
            return "8位以上连续相同"
    
    if re.match(r'^1[3-9]0{9}$', phone):
        return "号段后9位全为0(假号码)"
    
    return "号段不符合规范"


def parse_id_card(value: str) -> str:
    """解析身份证号格式（保留X）"""
    if not value:
        return ''
    value = str(value).strip()
    if value.startswith('="') and value.endswith('"'):
        value = value[2:-1]
    value = value.strip()
    value = re.sub(r'[^0-9Xx]', '', value)
    return value.upper()


def parse_phone(phone: str) -> str:
    """解析各种格式的电话号码"""
    if not phone:
        return ''
    phone = str(phone).strip()
    
    if phone.startswith('="') and phone.endswith('"'):
        phone = phone[2:-1]
    phone = phone.strip()
    
    if 'E' in phone.upper():
        try:
            phone = str(int(float(phone)))
        except:
            pass
    
    if '/' in phone:
        phone = phone.split('/')[0]
    
    phone = re.sub(r'[^\d]', '', phone)
    return phone


def is_valid_phone(phone: str) -> bool:
    """验证手机号是否有效"""
    if not phone:
        return False
    if len(phone) != 11:
        return False
    if phone[0] == '0':
        return False
    if not phone.isdigit():
        return False
    if len(set(phone)) == 1:
        return False
    if not re.match(r'^1[3-9]\d{9}$', phone):
        return False
    for i in range(len(phone) - 7):
        if len(set(phone[i:i+8])) == 1:
            return False
    if re.match(r'^1[3-9]0{9}$', phone):
        return False
    return True


def parse_date(date_str: str) -> Optional[datetime]:
    """解析日期字符串"""
    if not date_str:
        return None
    date_str = str(date_str).strip()
    for fmt in ['%Y/%m/%d', '%Y-%m-%d', '%Y%m%d']:
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    return None


def read_csv_data(csv_path: str, config: Dict) -> Tuple[List[Dict], List[Dict]]:
    """读取CSV文件
    
    返回: (记录列表, 无效身份证号日志列表)
    """
    records = []
    invalid_id_logs = []
    encoding = config.get('csv_encoding', 'gb18030')
    
    with open(csv_path, 'r', encoding=encoding) as f:
        reader = csv.reader(f)
        header = next(reader)
        
        id_col = config.get('id_card_col', 0)
        name_col = config.get('name_col', 1)
        phone_col = config.get('phone_col', 50)
        date_col = config.get('date_col', 59)
        unit_col = config.get('unit_col', 47)
        settlement_col = config.get('settlement_col', 48)
        employment_type_col = config.get('employment_type_col', 49)
        
        for row_num, row in enumerate(reader, start=2):
            if len(row) <= max(id_col, phone_col, date_col):
                continue
            
            id_card_raw = str(row[id_col]).strip()
            id_card = parse_id_card(id_card_raw)
            if not id_card:
                id_card = id_card_raw
            
            name = str(row[name_col]).strip()
            
            phone_raw = str(row[phone_col]) if row[phone_col] else ''
            phone = parse_phone(phone_raw)
            
            date_obj = parse_date(str(row[date_col]))
            
            if not id_card:
                continue
            
            is_valid, error_msg = validate_id_card(id_card)
            if not is_valid:
                invalid_id_logs.append({
                    'row': row_num,
                    'id_card': id_card,
                    'name': name,
                    'phone': phone if phone else '(无电话)',
                    'date': str(row[date_col]),
                    'error': error_msg
                })
            
            if name:
                records.append({
                    'row': row_num,
                    'id_card': id_card,
                    'name': name,
                    'phone': phone,
                    'phone_raw': phone_raw,
                    'date': date_obj,
                    'date_str': str(row[date_col]),
                    'id_valid': is_valid,
                    'id_error': error_msg if not is_valid else '',
                    'unit': str(row[unit_col]).strip() if len(row) > unit_col else '',
                    'settlement': str(row[settlement_col]).strip() if len(row) > settlement_col else '',
                    'employment_type': str(row[employment_type_col]).strip() if len(row) > employment_type_col else ''
                })
    
    return records, invalid_id_logs


def group_by_id_card(records: List[Dict]) -> Tuple[Dict[str, List[Dict]], List[Dict]]:
    """按身份证号分组
    
    返回: (分组字典, 重复详情列表)
    """
    grouped = {}
    for record in records:
        id_card = record['id_card']
        if id_card not in grouped:
            grouped[id_card] = []
        grouped[id_card].append(record)
    
    for id_card in grouped:
        grouped[id_card].sort(key=lambda x: x['date'] or datetime.min, reverse=True)
    
    duplicate_details = []
    for id_card, recs in grouped.items():
        if len(recs) > 1:
            duplicate_details.append({
                'id_card': id_card,
                'name': recs[0]['name'],
                'count': len(recs),
                'records': recs
            })
    
    return grouped, duplicate_details


def get_best_phone(records: List[Dict]) -> Tuple[str, str, str, str]:
    """从记录列表中选择最佳电话号码
    
    返回: (phone, name, date_str, reason)
    """
    for record in records:
        phone = record['phone']
        if is_valid_phone(phone):
            return phone, record['name'], record['date_str'], "电话有效"
    
    if records:
        phone = records[0]['phone']
        reason = get_id_card_invalid_reason(phone) if phone else "无电话"
        return phone, records[0]['name'], records[0]['date_str'], reason
    
    return '', '', '', ''


def convert_to_template(records: List[Dict], config: Dict, output_path: str,
                       invalid_id_logs: List[Dict], duplicate_details: List[Dict],
                       new_employee_ids: Set[str] = None) -> Tuple[int, Dict, List[Dict], List[Dict]]:
    """转换为薪福通模板格式
    
    返回: (输出人数, 统计信息, 无效手机号详情, 异常数据列表)
    """
    grouped, _ = group_by_id_card(records)
    
    results = []
    phone_warnings = []
    abnormal_records = []
    stats = {
        'total_records': len(records),
        'invalid_id_cards': len(invalid_id_logs),
        'duplicate_id_count': len(duplicate_details),
        'duplicate_record_count': sum(d['count'] for d in duplicate_details),
        'skipped_empty_id': 0,
        'valid_records': 0,
        'valid_phones': 0,
        'invalid_phones': 0,
        'skipped_no_phone': 0,
    }
    
    for id_card, recs in grouped.items():
        phone, name, date_str, reason = get_best_phone(recs)
        
        stats['valid_records'] += 1
        
        phone_invalid = False
        id_invalid = False
        
        if is_valid_phone(phone):
            stats['valid_phones'] += 1
        elif phone:
            stats['invalid_phones'] += 1
            phone_invalid = True
            phone_warnings.append({
                'id_card': id_card,
                'name': name,
                'phone': phone,
                'reason': reason
            })
        else:
            stats['skipped_no_phone'] += 1
            phone_invalid = True
        
        for rec in recs:
            if not rec.get('id_valid', True):
                id_invalid = True
                break
        
        for idx, rec in enumerate(recs):
            rec_issues = []
            if not rec.get('id_valid', True):
                rec_issues.append(f"身份证: {rec.get('id_error', '未知')}")
            if not is_valid_phone(rec['phone']):
                rec_issues.append(f"手机号: {get_id_card_invalid_reason(rec['phone'])}")
            
            if rec_issues or len(recs) > 1:
                abnormal_records.append({
                    'name': rec['name'],
                    'phone': rec['phone'],
                    'id_card': rec['id_card'],
                    'issues': '; '.join(rec_issues) if rec_issues else ('(选用数据)' if idx == 0 else '(已剔除)'),
                    'source_row': rec.get('row', ''),
                    'date': rec['date_str'],
                    'is_selected': idx == 0,
                    'has_issue': bool(rec_issues)
                })
        
        results.append({
            'name': name,
            'phone': phone if phone else '',
            'id_type': '居民身份证',
            'id_card': id_card,
            'date': date_str,
            'phone_invalid': phone_invalid,
            'id_invalid': id_invalid,
            'unit': recs[0].get('unit', ''),
            'settlement': recs[0].get('settlement', ''),
            'employment_type': recs[0].get('employment_type', ''),
            'is_new': new_employee_ids is not None and id_card in new_employee_ids,
        })
    
    if not os.path.exists(TEMPLATE_FILE):
        return None, stats, phone_warnings, abnormal_records
    
    shutil.copy(TEMPLATE_FILE, output_path)
    
    wb = openpyxl.load_workbook(output_path)
    sheet_name = config.get('sheet_name', '花名册')
    
    if sheet_name not in wb.sheetnames:
        return None, stats, phone_warnings, abnormal_records
    
    ws = wb[sheet_name]
    
    for i, record in enumerate(results, start=2):
        ws.cell(row=i, column=1, value=record['name'])
        
        phone_cell = ws.cell(row=i, column=2, value=record['phone'])
        if record.get('phone_invalid'):
            phone_cell.fill = YELLOW_FILL
        
        ws.cell(row=i, column=3, value=record['id_type'])
        
        id_cell = ws.cell(row=i, column=4, value=record['id_card'])
        if record.get('id_invalid'):
            id_cell.fill = YELLOW_FILL
        
        ws.cell(row=i, column=5, value=record.get('unit', ''))
        ws.cell(row=i, column=6, value=record.get('settlement', ''))
        ws.cell(row=i, column=7, value=record.get('employment_type', ''))
        
        if record.get('is_new'):
            for col in range(1, 8):
                cell = ws.cell(row=i, column=col)
                if not (cell.fill and cell.fill.fgColor.rgb == 'FFFFFF00'):
                    cell.fill = GREEN_FILL
    
    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 13
    ws.column_dimensions['C'].width = 13
    ws.column_dimensions['D'].width = 20
    ws.column_dimensions['E'].width = 20
    ws.column_dimensions['F'].width = 20
    ws.column_dimensions['G'].width = 12
    
    if ABNORMAL_SHEET_NAME in wb.sheetnames:
        del wb[ABNORMAL_SHEET_NAME]
    ab_ws = wb.create_sheet(ABNORMAL_SHEET_NAME)
    
    headers = ['姓名', '手机号', '证件号码', '问题说明', '来源行号', '经办日期', '是否选用']
    for col, header in enumerate(headers, start=1):
        cell = ab_ws.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True)
    
    for i, rec in enumerate(abnormal_records, start=2):
        ab_ws.cell(row=i, column=1, value=rec['name'])
        ab_ws.cell(row=i, column=2, value=rec['phone'])
        ab_ws.cell(row=i, column=3, value=rec['id_card'])
        ab_ws.cell(row=i, column=4, value=rec['issues'])
        ab_ws.cell(row=i, column=5, value=rec['source_row'])
        ab_ws.cell(row=i, column=6, value=rec['date'])
        ab_ws.cell(row=i, column=7, value='是' if rec['is_selected'] else '否')
        
        if rec['has_issue']:
            for col in range(1, 8):
                ab_ws.cell(row=i, column=col).fill = YELLOW_FILL
    
    ab_ws.column_dimensions['A'].width = 12
    ab_ws.column_dimensions['B'].width = 15
    ab_ws.column_dimensions['C'].width = 20
    ab_ws.column_dimensions['D'].width = 35
    ab_ws.column_dimensions['E'].width = 10
    ab_ws.column_dimensions['F'].width = 12
    ab_ws.column_dimensions['G'].width = 10
    
    wb.save(output_path)
    
    return len(results), stats, phone_warnings, abnormal_records


def log_details(log_fn, stats, invalid_id_logs, duplicate_details, phone_warnings):
    """输出详细处理日志"""
    log_fn("=" * 60)
    
    if invalid_id_logs:
        log_fn("[WARNING] 无效身份证号详情:")
        for item in invalid_id_logs:
            log_fn(f"  行{item['row']}: {item['id_card']} ({item['name']})")
            log_fn(f"         电话: {item['phone']} | 日期: {item['date']}")
            log_fn(f"         原因: {item['error']}")
        log_fn(f"[INFO] 无效身份证号: {len(invalid_id_logs)}条")
        log_fn("")
    else:
        log_fn("[INFO] 身份证号校验: 全部通过")
    
    if duplicate_details:
        log_fn("[INFO] 身份证号重复详情:")
        for item in duplicate_details:
            log_fn(f"  {item['id_card']} ({item['name']}): {item['count']}条记录")
            for rec in item['records']:
                phone_status = "✓" if is_valid_phone(rec['phone']) else "✗"
                log_fn(f"    - 行{rec.get('row', '?')}: 电话={rec['phone']} {phone_status}, 日期={rec['date_str']}")
            phone, name, date_str, reason = get_best_phone(item['records'])
            log_fn(f"    → 选用: {phone} ({reason})")
        log_fn(f"[INFO] 身份证号重复: {len(duplicate_details)}个, 共{stats['duplicate_record_count']}条记录")
        log_fn("")
    else:
        log_fn("[INFO] 身份证号重复: 无")
    
    if phone_warnings:
        log_fn("[WARNING] 无效手机号详情:")
        for item in phone_warnings:
            log_fn(f"  {item['name']}: {item['phone']}")
            log_fn(f"    原因: {item['reason']}")
        log_fn(f"[INFO] 无效手机号: {len(phone_warnings)}个")
        log_fn("")
    else:
        log_fn("[INFO] 手机号校验: 全部通过")
    
    log_fn("=" * 60)


def log_summary(log_fn, output_path, stats):
    """输出统计汇总"""
    log_fn("")
    log_fn("【统计汇总】")
    log_fn(f"  原始记录数: {stats['total_records']}")
    log_fn(f"  无效身份证号: {stats['invalid_id_cards']}")
    log_fn(f"  身份证号重复: {stats['duplicate_id_count']} ({stats['duplicate_record_count']}条记录)")
    log_fn(f"  ─────────────────")
    log_fn(f"  输出人数: {stats['valid_records']}")
    log_fn(f"  有效手机号: {stats['valid_phones']}")
    log_fn(f"  无效手机号: {stats['invalid_phones']}")
    log_fn(f"  无手机号: {stats['skipped_no_phone']}")
    log_fn("=" * 60)
    log_fn("")
    log_fn(f"输出文件: {os.path.abspath(output_path)}")
    log_fn("")


def main_cli():
    """命令行模式"""
    parser = argparse.ArgumentParser(description='员工信息格式转换工具')
    parser.add_argument('--csv', '-i', help='CSV文件路径')
    parser.add_argument('--output', '-o', help='输出文件路径(可选，默认自动生成)')
    parser.add_argument('--compare', '-c', action='store_true', help='转换后自动比对新增人员')
    parser.add_argument('--reference', '-r', help='指定参考文件(上次生成的文件)')
    parser.add_argument('--no-compare', action='store_true', help='跳过比对提示')
    args = parser.parse_args(sys.argv[2:] if '--cli' in sys.argv else sys.argv[1:])
    
    print("=" * 60)
    print("员工信息格式转换工具 (命令行模式)")
    print("=" * 60)
    print()
    
    config = load_config()
    
    csv_file = args.csv or config.get('recent_csv', '')
    
    if not csv_file or not os.path.exists(csv_file):
        csv_file = input("请输入CSV文件路径: ").strip()
    
    if not os.path.exists(csv_file):
        print(f"错误: CSV文件不存在: {csv_file}")
        return
    
    output_file = args.output
    if not output_file:
        output_file = get_output_path(config.get('recent_output_filename', '员工导入模板_转换后'))
    
    print(f"模板文件: {TEMPLATE_FILE}")
    print(f"输出文件: {output_file}")
    print()
    
    print("=" * 60)
    print(f"[INFO] 开始处理CSV文件: {csv_file}")
    
    records, invalid_id_logs = read_csv_data(csv_file, config)
    print(f"[INFO] 读取记录数: {len(records)}")
    
    grouped, duplicate_details = group_by_id_card(records)
    
    reference_file = args.reference
    if not reference_file:
        reference_file = find_latest_cleaned_file(output_file)
    
    new_employee_ids = None
    new_employees_list = []
    
    do_compare = args.compare
    if not do_compare and not args.no_compare and reference_file:
        print()
        ref_name = os.path.basename(reference_file)
        response = input(f"是否比对新增人员？（参考文件: {ref_name}）[Y/n]: ").strip().lower()
        do_compare = response in ('', 'y', 'yes', '是')
    
    if do_compare:
        if reference_file and os.path.exists(reference_file):
            new_employees_list, ref_file = find_new_employees_from_records(records, reference_file, config)
            if new_employees_list:
                new_employee_ids = set(emp['id_card'] for emp in new_employees_list)
                print(f"[INFO] 比对完成，发现 {len(new_employees_list)} 名新增人员")
            else:
                print("[INFO] 没有发现新增人员")
        else:
            print("[WARNING] 未找到参考文件，跳过比对")
    
    print(f"[INFO] 开始转换...")
    count, stats, phone_warnings, abnormal_records = convert_to_template(
        records, config, output_file, invalid_id_logs, duplicate_details, new_employee_ids
    )
    
    if count is None:
        print("[ERROR] 转换失败: 模板文件不存在")
        return
    
    print("")
    log_details(print, stats, invalid_id_logs, duplicate_details, phone_warnings)
    log_summary(print, output_file, stats)
    abnormal_count = sum(1 for r in abnormal_records if r['has_issue'])
    print(f"[INFO] 异常数据Sheet已生成，共{len(abnormal_records)}条记录")
    print(f"       - 有异常: {abnormal_count}条 (已标黄)")
    print(f"       - 重复数据: {len(abnormal_records) - abnormal_count}条 (供参考)")
    
    if new_employees_list:
        ref_file = reference_file
        count_new = create_new_employees_sheet(output_file, new_employees_list, ref_file)
        ref_name = os.path.basename(ref_file) if ref_file else ''
        print(f"[INFO] 新增人员Sheet已创建，共 {count_new} 名新增人员")
        print(f"       参考文件: {ref_name}")
        for emp in new_employees_list:
            unit = emp.get('unit', '')
            print(f"       - {emp['name']}: {emp['id_card']}{f' ({unit})' if unit else ''}")
    
    print("转换完成!")
    
    config['recent_csv'] = csv_file
    save_config(config)


def main_gui():
    """GUI模式"""
    if not HAS_TKINTER:
        print("错误: tkinter未安装，请使用命令行模式 (python script.py --cli)")
        return
    
    if not HAS_OPENPYXL:
        print("错误: openpyxl未安装，请使用命令行模式 (python script.py --cli)")
        return
    
    root = tkinterdnd2.Tk()
    root.title("员工信息格式转换工具")
    root.geometry("700x580")
    root.resizable(True, True)
    
    config = load_config()
    
    csv_path_var = tk.StringVar(value=config.get('recent_csv', ''))
    output_filename_var = tk.StringVar(value=config.get('recent_output_filename', '员工导入模板_转换后'))
    reference_path_var = tk.StringVar()
    compare_var = tk.BooleanVar(value=True)
    
    log_text = None
    stats_label = None
    csv_drop_label = None
    
    def log(message: str):
        if log_text:
            timestamp = datetime.now().strftime("%H:%M:%S")
            log_text.insert(tk.END, f"[{timestamp}] {message}\n")
            log_text.see(tk.END)
            root.update()
    
    def update_stats(stats: Dict = None):
        if stats_label:
            if stats:
                stats_label.config(
                    text=f"总记录数: {stats.get('total_records', 0)} | "
                         f"有效手机号: {stats.get('valid_phones', 0)} | "
                         f"无效手机号: {stats.get('invalid_phones', 0)} | "
                         f"无手机号: {stats.get('skipped_no_phone', 0)}"
                )
            else:
                stats_label.config(text="总记录数: 0 | 有效手机号: 0 | 无效手机号: 0 | 无手机号: 0")
    
    def update_csv_label():
        """更新CSV文件显示"""
        if csv_drop_label:
            csv_path = csv_path_var.get()
            if csv_path:
                csv_drop_label.config(text=csv_path, fg='black')
            else:
                csv_drop_label.config(text="拖拽 CSV 文件到此处 或 点击选择", fg='gray')
    
    def browse_csv():
        filepath = filedialog.askopenfilename(
            title="选择CSV文件",
            filetypes=[("CSV 文件", "*.csv"), ("所有文件", "*.*")]
        )
        if filepath:
            csv_path_var.set(filepath)
            config['recent_csv'] = filepath
            save_config(config)
            update_csv_label()
            log(f"已选择CSV文件: {filepath}")
    
    def on_csv_drop(event):
        files = root.tk.splitlist(event.data)
        for file_path in files:
            file_path = file_path.strip()
            if file_path.startswith('{') and file_path.endswith('}'):
                file_path = file_path[1:-1]
            if file_path and file_path.endswith('.csv'):
                csv_path_var.set(file_path)
                config['recent_csv'] = file_path
                save_config(config)
                update_csv_label()
                log(f"已选择CSV文件: {file_path}")
                break
    
    def start_convert():
        csv_file = csv_path_var.get().strip()
        
        if not csv_file:
            messagebox.showwarning("警告", "请选择CSV文件")
            return
        
        if not os.path.exists(csv_file):
            messagebox.showerror("错误", f"CSV文件不存在: {csv_file}")
            return
        
        if not os.path.exists(TEMPLATE_FILE):
            messagebox.showerror("错误", f"模板文件不存在: {TEMPLATE_FILE}")
            return
        
        filename_prefix = output_filename_var.get().strip() or '员工导入模板_转换后'
        output_file = get_output_path(filename_prefix)
        
        log_text.delete(1.0, tk.END)
        log("=" * 60)
        log("开始转换...")
        log(f"模板文件: {TEMPLATE_FILE}")
        log(f"CSV文件: {csv_file}")
        log("")
        
        try:
            log("读取CSV文件...")
            records, invalid_id_logs = read_csv_data(csv_file, config)
            log(f"读取记录数: {len(records)}")
            log("")
            
            log("分组去重...")
            grouped, duplicate_details = group_by_id_card(records)
            log(f"分组完成，共{len(grouped)}个不同身份证号")
            log("")
            
            new_employee_ids = None
            new_employees_list = []
            
            if compare_var.get():
                reference_file = reference_path_var.get().strip()
                if not reference_file:
                    reference_file = find_latest_cleaned_file(output_file)
                
                if reference_file and os.path.exists(reference_file):
                    new_employees_list, ref_file = find_new_employees_from_records(records, reference_file, config)
                    if new_employees_list:
                        new_employee_ids = set(emp['id_card'] for emp in new_employees_list)
                        log(f"[INFO] 比对完成，发现 {len(new_employees_list)} 名新增人员")
                    else:
                        log("[INFO] 没有发现新增人员")
                else:
                    log("[WARNING] 未找到参考文件，跳过比对")
            
            log("开始转换...")
            count, stats, phone_warnings, abnormal_records = convert_to_template(
                records, config, output_file, invalid_id_logs, duplicate_details, new_employee_ids
            )
            
            if count is None:
                log("[ERROR] 转换失败: 模板文件无效")
                messagebox.showerror("错误", "模板文件无效")
                return
            
            log("")
            log_details(log, stats, invalid_id_logs, duplicate_details, phone_warnings)
            log_summary(log, output_file, stats)
            abnormal_count = sum(1 for r in abnormal_records if r['has_issue'])
            log(f"[INFO] 异常数据Sheet已生成，共{len(abnormal_records)}条记录")
            log(f"       - 有异常: {abnormal_count}条 (已标黄)")
            log(f"       - 重复数据: {len(abnormal_records) - abnormal_count}条 (供参考)")
            
            if new_employees_list:
                ref_file = reference_file
                count_new = create_new_employees_sheet(output_file, new_employees_list, ref_file)
                ref_name = os.path.basename(ref_file) if ref_file else ''
                log(f"[INFO] 新增人员Sheet已创建，共 {count_new} 名新增人员")
                log(f"       参考文件: {ref_name}")
                for emp in new_employees_list:
                    unit = emp.get('unit', '')
                    log(f"       - {emp['name']}: {emp['id_card']}{f' ({unit})' if unit else ''}")
            
            log("转换完成!")
            
            update_stats(stats)
            
            config['recent_output_filename'] = filename_prefix
            config['recent_csv'] = csv_file
            save_config(config)
            
            messagebox.showinfo("完成", f"转换完成!\n\n输出文件:\n{os.path.abspath(output_file)}\n\n转换人数: {count}")
            
        except Exception as e:
            log(f"[ERROR] {e}")
            import traceback
            log(traceback.format_exc())
            messagebox.showerror("错误", str(e))
    
    main_frame = ttk.Frame(root, padding="10")
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    title_label = ttk.Label(main_frame, text="员工信息格式转换工具", 
                           font=('微软雅黑', 16, 'bold'))
    title_label.pack(pady=(0, 15))
    
    csv_frame = ttk.Frame(main_frame)
    csv_frame.pack(fill=tk.X, pady=5)
    ttk.Label(csv_frame, text="源CSV文件:").pack(side=tk.LEFT)
    
    csv_drop_frame = tk.Frame(csv_frame, bd=2, relief=tk.SUNKEN, bg='white', height=35)
    csv_drop_frame.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
    csv_drop_frame.pack_propagate(False)
    
    csv_drop_label = tk.Label(csv_drop_frame, text="拖拽 CSV 文件到此处 或 点击选择", 
                              bg='white', fg='gray', font=('song', 9), anchor='w')
    csv_drop_label.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
    
    csv_drop_frame.drop_target_register(DND_FILES)
    csv_drop_frame.dnd_bind('<<Drop>>', on_csv_drop)
    csv_drop_frame.bind('<Button-1>', lambda e: browse_csv())
    
    ttk.Button(csv_frame, text="浏览...", command=browse_csv).pack(side=tk.LEFT, padx=5)
    
    output_frame = ttk.Frame(main_frame)
    output_frame.pack(fill=tk.X, pady=5)
    ttk.Label(output_frame, text="输出文件名:").pack(side=tk.LEFT)
    ttk.Entry(output_frame, textvariable=output_filename_var, width=35).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
    ttk.Label(output_frame, text=".xlsx", font=('song', 9)).pack(side=tk.LEFT)
    
    ttk.Label(output_frame, text="  输出目录: Cleaned/", font=('song', 9), foreground='gray').pack(side=tk.LEFT, padx=(20, 0))
    
    compare_frame = ttk.Frame(main_frame)
    compare_frame.pack(fill=tk.X, pady=5)
    
    ttk.Checkbutton(compare_frame, text="比对新增人员", variable=compare_var).pack(side=tk.LEFT)
    
    ttk.Label(compare_frame, text="  参考文件:").pack(side=tk.LEFT)
    
    ref_drop_frame = tk.Frame(compare_frame, bd=1, relief=tk.SUNKEN, bg='white', height=25)
    ref_drop_frame.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
    ref_drop_frame.pack_propagate(False)
    
    ref_drop_label = tk.Label(ref_drop_frame, text="默认使用最近一次生成的文件", 
                              bg='white', fg='gray', font=('song', 8), anchor='w')
    ref_drop_label.pack(fill=tk.BOTH, expand=True, padx=3, pady=2)
    
    def browse_reference():
        filepath = filedialog.askopenfilename(
            title="选择参考文件（上次生成的转换文件）",
            initialdir=OUTPUT_DIR,
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if filepath:
            reference_path_var.set(filepath)
            ref_drop_label.config(text=os.path.basename(filepath), fg='black')
            log(f"已选择参考文件: {filepath}")
    
    ref_drop_frame.bind('<Button-1>', lambda e: browse_reference())
    
    ttk.Button(compare_frame, text="浏览...", command=browse_reference, width=6).pack(side=tk.LEFT)
    
    def update_ref_label():
        if ref_drop_label:
            ref_path = reference_path_var.get()
            if ref_path:
                ref_drop_label.config(text=os.path.basename(ref_path), fg='black')
            else:
                ref_drop_label.config(text="默认使用最近一次生成的文件", fg='gray')
    
    stats_frame = ttk.LabelFrame(main_frame, text="转换统计", padding="10")
    stats_frame.pack(fill=tk.X, pady=10)
    
    stats_label = ttk.Label(stats_frame, text="总记录数: 0 | 有效手机号: 0 | 无效手机号: 0 | 无手机号: 0",
                            font=('Consolas', 10))
    stats_label.pack()
    
    log_frame = ttk.LabelFrame(main_frame, text="日志", padding="5")
    log_frame.pack(fill=tk.BOTH, expand=True, pady=10)
    
    log_scroll_y = ttk.Scrollbar(log_frame)
    log_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
    
    log_text = tk.Text(log_frame, height=12, font=('Consolas', 9),
                       yscrollcommand=log_scroll_y.set, wrap=tk.WORD)
    log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    log_scroll_y.config(command=log_text.yview)
    
    button_frame = ttk.Frame(main_frame)
    button_frame.pack(pady=10)
    
    convert_btn = ttk.Button(button_frame, text="开始转换", command=start_convert)
    convert_btn.pack(padx=10)
    
    update_csv_label()
    
    if csv_path_var.get():
        log(f"已加载CSV: {csv_path_var.get()}")
    
    root.mainloop()


if __name__ == '__main__':
    if '--cli' in sys.argv:
        main_cli()
    else:
        main_gui()
