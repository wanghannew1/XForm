import os
import re
import json
import sys
import xlwings as xw
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
from tkinter import ttk
import threading
import pythoncom
import tkinterdnd2
from tkinterdnd2 import DND_FILES

# ---------- 打包后的路径处理 ----------
if getattr(sys, 'frozen', False):
    APPLICATION_PATH = os.path.dirname(sys.executable)
else:
    APPLICATION_PATH = os.path.dirname(os.path.abspath(__file__))

# 如果不是控制台模式，重定向输出到文件（调试用）
if getattr(sys, 'frozen', False) and sys.stderr is None:
    sys.stderr = open(os.path.join(APPLICATION_PATH, 'error.log'), 'w', encoding='utf-8')
    sys.stdout = open(os.path.join(APPLICATION_PATH, 'output.log'), 'w', encoding='utf-8')

CM_TO_POINTS = 28.35
YEAR_MONTH_PATTERN = re.compile(r'(\d{4})年(\d{2})月')
SCRIPT_DIR = APPLICATION_PATH
CONFIG_FILE = os.path.join(SCRIPT_DIR, "xform_config.json")

# ---------- 配置持久化 ----------
HIDE_COLUMNS_CONFIG = {
    "header_row": 3,
    "enabled_modes": ["mode1", "mode2"],
    "mode1_columns": [],
    "mode1_headers": ["实发合计"],
    "mode2_columns": [],
    "mode2_headers": ["部门", "职工号", "交通补贴"],
    "clear_sign_rows": True
}

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return HIDE_COLUMNS_CONFIG.copy()
    return HIDE_COLUMNS_CONFIG.copy()

def save_config(config):
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
    except:
        pass

# ---------- 工具函数 ----------
def sanitize_filename(name):
    """移除文件名中的非法字符"""
    illegal_chars = r'[\\/:\*\?"<>\|]'
    return re.sub(illegal_chars, '_', name)

def ensure_2d_list(value):
    """确保 xlwings range.value 为二维列表"""
    if value is None:
        return []
    if not isinstance(value, list):
        return [[value]]
    if value and not isinstance(value[0], list):
        return [value]
    return value

# ---------- 应用主类 ----------
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("工资表打印格式批量调整工具")
        self.root.geometry("700x800")
        self.file_list = []

        tk.Label(root, text="待处理文件列表：(支持拖拽文件至此)", font=("微软雅黑", 10)).pack(anchor='w', padx=10, pady=5)
        self.listbox = tk.Listbox(root, selectmode=tk.EXTENDED, height=8)
        self.listbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # 拖拽支持
        self.listbox.drop_target_register(DND_FILES)
        self.listbox.dnd_bind('<<Drop>>', self.on_drop)

        btn_frame = tk.Frame(root)
        btn_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Button(btn_frame, text="添加文件", command=self.add_files).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="清空列表", command=self.clear_list).pack(side=tk.LEFT, padx=5)
        self.process_btn = tk.Button(btn_frame, text="开始处理", command=self.start_process)
        self.process_btn.pack(side=tk.RIGHT, padx=5)

        self.progress = ttk.Progressbar(root, orient='horizontal', length=400, mode='determinate')
        self.progress.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(root, text="处理日志：", font=("微软雅黑", 10)).pack(anchor='w', padx=10)
        self.log = scrolledtext.ScrolledText(root, height=12, state='normal')
        self.log.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.log_insert("拖拽功能已启用")

        # 加载配置
        self.config = load_config()

        # ---------- 删除列配置 UI ----------
        hide_frame = tk.LabelFrame(root, text="删除列配置", padx=10, pady=5)
        hide_frame.pack(fill=tk.X, padx=10, pady=5)

        row1_frame = tk.Frame(hide_frame)
        row1_frame.pack(fill=tk.X)
        tk.Label(row1_frame, text="列头行号：").pack(side=tk.LEFT)
        self.header_row_var = tk.StringVar(value=str(self.config["header_row"]))
        tk.Entry(row1_frame, textvariable=self.header_row_var, width=5).pack(side=tk.LEFT, padx=(0, 20))

        self.mode1_enabled = tk.BooleanVar(value="mode1" in self.config.get("enabled_modes", []))
        self.mode2_enabled = tk.BooleanVar(value="mode2" in self.config.get("enabled_modes", []))
        tk.Checkbutton(row1_frame, text="模式1：直接删除", variable=self.mode1_enabled,
                       onvalue=True, offvalue=False).pack(side=tk.LEFT, padx=5)
        tk.Checkbutton(row1_frame, text="模式2：无值则删除", variable=self.mode2_enabled,
                       onvalue=True, offvalue=False).pack(side=tk.LEFT, padx=5)

        mode1_frame = tk.LabelFrame(hide_frame, text="模式1配置", padx=5, pady=5)
        mode1_frame.pack(fill=tk.X, pady=5)
        tk.Label(mode1_frame, text="列地址(逗号分隔)：").pack(side=tk.LEFT)
        self.mode1_columns_var = tk.StringVar(value=",".join(self.config["mode1_columns"]))
        tk.Entry(mode1_frame, textvariable=self.mode1_columns_var, width=20).pack(side=tk.LEFT, padx=5)
        tk.Label(mode1_frame, text="列头名称(逗号分隔)：").pack(side=tk.LEFT)
        self.mode1_headers_var = tk.StringVar(value=",".join(self.config["mode1_headers"]))
        tk.Entry(mode1_frame, textvariable=self.mode1_headers_var, width=25).pack(side=tk.LEFT, padx=5)

        mode2_frame = tk.LabelFrame(hide_frame, text="模式2配置", padx=5, pady=5)
        mode2_frame.pack(fill=tk.X, pady=5)
        tk.Label(mode2_frame, text="列地址(逗号分隔)：").pack(side=tk.LEFT)
        self.mode2_columns_var = tk.StringVar(value=",".join(self.config["mode2_columns"]))
        tk.Entry(mode2_frame, textvariable=self.mode2_columns_var, width=20).pack(side=tk.LEFT, padx=5)
        tk.Label(mode2_frame, text="列头名称(逗号分隔)：").pack(side=tk.LEFT)
        self.mode2_headers_var = tk.StringVar(value=",".join(self.config["mode2_headers"]))
        tk.Entry(mode2_frame, textvariable=self.mode2_headers_var, width=25).pack(side=tk.LEFT, padx=5)

        # ---------- 签名行选项 ----------
        sign_frame = tk.LabelFrame(hide_frame, text="签名行配置", padx=10, pady=5)
        sign_frame.pack(fill=tk.X, pady=5)
        self.clear_sign_rows_var = tk.BooleanVar(value=self.config.get("clear_sign_rows", True))
        tk.Checkbutton(sign_frame, text="清空签名行（合计行与制表人之间的总经理签字等提示词）",
                       variable=self.clear_sign_rows_var, onvalue=True, offvalue=False).pack(anchor=tk.W)

    # ---------- 文件列表操作 ----------
    def on_drop(self, event):
        files = self.root.tk.splitlist(event.data)
        for file_path in files:
            file_path = file_path.strip()
            if file_path.startswith('{') and file_path.endswith('}'):
                file_path = file_path[1:-1]
            if file_path and (file_path.endswith('.xls') or file_path.endswith('.xlsx')):
                if os.path.exists(file_path) and file_path not in self.file_list:
                    self.file_list.append(file_path)
                    self.listbox.insert(tk.END, os.path.basename(file_path))
                    self.log_insert(f"已添加：{os.path.basename(file_path)}")
        return 'action'

    def add_files(self):
        files = filedialog.askopenfilenames(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xls *.xlsx")]
        )
        for f in files:
            if f not in self.file_list:
                self.file_list.append(f)
                self.listbox.insert(tk.END, os.path.basename(f))

    def clear_list(self):
        self.file_list.clear()
        self.listbox.delete(0, tk.END)

    # ---------- 配置读取（主线程） ----------
    def get_hide_config(self):
        config = {
            "header_row": int(self.header_row_var.get()),
            "enabled_modes": [],
            "mode1_columns": [c.strip() for c in self.mode1_columns_var.get().split(",") if c.strip()],
            "mode1_headers": [h.strip() for h in self.mode1_headers_var.get().split(",") if h.strip()],
            "mode2_columns": [c.strip() for c in self.mode2_columns_var.get().split(",") if c.strip()],
            "mode2_headers": [h.strip() for h in self.mode2_headers_var.get().split(",") if h.strip()],
            "clear_sign_rows": self.clear_sign_rows_var.get()
        }
        if self.mode1_enabled.get():
            config["enabled_modes"].append("mode1")
        if self.mode2_enabled.get():
            config["enabled_modes"].append("mode2")
        return config

    # ---------- 列操作工具 ----------
    @staticmethod
    def column_letter_to_index(letter):
        result = 0
        for char in letter.upper():
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result

    @staticmethod
    def index_to_column_letter(idx):
        letters = ''
        while idx > 0:
            idx, remainder = divmod(idx - 1, 26)
            letters = chr(65 + remainder) + letters
        return letters

    def find_column_by_header(self, ws, header_name, header_row, last_col):
        for col in range(1, last_col + 1):
            cell_value = ws.range((header_row, col)).value
            if cell_value is not None and str(cell_value).strip() == header_name:
                return col
        return None

    def is_column_empty(self, ws, col_idx, start_row, end_row):
        """检查列在指定行范围内是否为空（0、None、空白字符串均视为空）"""
        for row in range(start_row, end_row + 1):
            val = ws.range((row, col_idx)).value
            if val is not None:
                if isinstance(val, (int, float)) and val == 0:
                    continue
                if isinstance(val, str) and val.strip() == '':
                    continue
                return False
        return True

    def get_columns_to_delete(self, ws, config, data_end_row, last_col):
        """
        data_end_row: 数据区域结束行（合计行），检查空值时只到此行
        """
        delete_list = []
        enabled_modes = config.get("enabled_modes", [])
        header_row = config["header_row"]

        if "mode1" in enabled_modes:
            for col_letter in config["mode1_columns"]:
                col_idx = self.column_letter_to_index(col_letter)
                if 1 <= col_idx <= last_col:
                    delete_list.append(col_idx)
            for header in config["mode1_headers"]:
                col_idx = self.find_column_by_header(ws, header, header_row, last_col)
                if col_idx and col_idx not in delete_list:
                    delete_list.append(col_idx)

        if "mode2" in enabled_modes:
            for col_letter in config["mode2_columns"]:
                col_idx = self.column_letter_to_index(col_letter)
                if 1 <= col_idx <= last_col and self.is_column_empty(ws, col_idx, header_row + 1, data_end_row):
                    delete_list.append(col_idx)
            for header in config["mode2_headers"]:
                col_idx = self.find_column_by_header(ws, header, header_row, last_col)
                if col_idx and col_idx not in delete_list and self.is_column_empty(ws, col_idx, header_row + 1, data_end_row):
                    delete_list.append(col_idx)

        return list(set(delete_list))

    def merge_with_content(self, rng, sep=' '):
        """合并单元格并保留所有非空内容（用 sep 连接）"""
        values = []
        data = rng.value
        if data is None:
            data = []
        elif not isinstance(data, list):
            data = [[data]]
        elif data and not isinstance(data[0], list):
            data = [data]

        for row in data:
            for v in row:
                if v is not None and str(v).strip() != '':
                    values.append(str(v).strip())
        merged_text = sep.join(values)
        rng.merge()
        rng.value = merged_text
        return merged_text

    # ---------- 批量处理 ----------
    def start_process(self):
        save_config(self.get_hide_config())
        if not self.file_list:
            messagebox.showwarning("提示", "请先添加文件！")
            return

        self.process_btn.config(state='disabled')
        self.progress['value'] = 0
        self.log.delete(1.0, tk.END)

        # 在主线程读取配置，传入工作线程
        config = self.get_hide_config()
        file_list = self.file_list.copy()

        # 使用 daemon=False 确保线程不会被强制终止
        thread = threading.Thread(target=self.process_files, args=(file_list, config))
        thread.daemon = False
        thread.start()

    def process_files(self, file_list, config):
        """后台处理线程，初始化 COM 接口"""
        # 在线程中初始化 COM（打包后必须）
        pythoncom.CoInitialize()

        total = len(file_list)
        app = xw.App(visible=False, add_book=False)
        try:
            for idx, file_path in enumerate(file_list):
                self.log_insert(f"正在处理：{os.path.basename(file_path)}")
                try:
                    self.process_one_file(app, file_path, config)
                    self.log_insert(f"处理成功：{os.path.basename(file_path)}")
                except Exception as e:
                    self.log_insert(f"处理失败：{os.path.basename(file_path)} - {str(e)}")
                progress = int((idx + 1) / total * 100)
                self.root.after(0, lambda p=progress: self.progress.configure(value=p))
        finally:
            if app:
                try:
                    app.quit()
                except:
                    pass
            # 反初始化 COM
            pythoncom.CoUninitialize()
            self.root.after(0, lambda: self.process_btn.config(state='normal'))
            self.root.after(0, lambda: messagebox.showinfo("完成", "全部处理完成！"))

    def log_insert(self, msg):
        self.root.after(0, lambda: self.log.insert(tk.END, msg + "\n"))
        self.root.after(0, lambda: self.log.see(tk.END))

    # ---------- 单文件处理核心 ----------
    def process_one_file(self, app, file_path, config):
        wb = app.books.open(file_path)
        try:
            ws = wb.sheets[0]

            # 0. 取消所有隐藏列，并临时设置列宽
            ws.api.Cells.EntireColumn.Hidden = False
            ws.api.Cells.EntireColumn.ColumnWidth = 9
            ws.used_range
            used = ws.used_range
            last_row = used.last_cell.row
            last_col = used.last_cell.column
            self.log_insert("已取消所有隐藏列，刷新表格范围")

            # 1. 基础信息（新文件名）
            row2_colB = ws.range('B2').value
            name_part = str(row2_colB).replace("名称：", "").strip() if row2_colB else "未知名称"
            name_part = sanitize_filename(name_part)

            row1_colC = ws.range('C1').value
            year_month = ""
            if row1_colC:
                match = YEAR_MONTH_PATTERN.search(str(row1_colC))
                if match:
                    year_month = f"{match.group(1)}年{match.group(2)}月"

            new_filename = f"{name_part}{year_month}工资表.xlsx"
            original_dir = os.path.dirname(file_path)
            new_file_path = os.path.join(original_dir, new_filename)

            header_row = config["header_row"]

            # 2. 先确定数据区域结束行（排除最后三行）
            maker_row = None
            for r in range(last_row, max(1, last_row - 5), -1):
                for c in range(1, last_col + 1):
                    val = ws.range((r, c)).value
                    if val and isinstance(val, str) and "制表人：" in val:
                        maker_row = r
                        break
                if maker_row:
                    break
            if not maker_row:
                raise ValueError("未找到“制表人：”行，无法继续处理。")

            # 数据区域结束行 = 合计行 = 制表人往上3行
            table_end_row = maker_row - 3
            if table_end_row < header_row:
                table_end_row = header_row

            # 3. 清空倒数第2、3行（仅在配置允许时）
            if config.get("clear_sign_rows", True):
                for clear_row in (maker_row - 1, maker_row - 2):
                    if clear_row >= 1:
                        ws.range((clear_row, 1), (clear_row, last_col)).value = None
                        self.log_insert(f"已清空第{clear_row}行")

            # 4. 删除列（判断空值时只到 table_end_row）
            delete_cols = self.get_columns_to_delete(ws, config, table_end_row, last_col)
            for col_idx in sorted(delete_cols, reverse=True):
                header_val = ws.range((header_row, col_idx)).value
                col_name = str(header_val).strip() if header_val else self.index_to_column_letter(col_idx)
                ws.api.Columns(col_idx).Delete()
                self.log_insert(f"已删除列：{col_name}")

            # 删除后刷新区域
            ws.used_range
            last_row = ws.used_range.last_cell.row
            last_col = ws.used_range.last_cell.column

            # 重新确定 maker_row
            maker_row_updated = None
            for r in range(last_row, max(1, last_row - 5), -1):
                for c in range(1, last_col + 1):
                    val = ws.range((r, c)).value
                    if val and isinstance(val, str) and "制表人：" in val:
                        maker_row_updated = r
                        break
                if maker_row_updated:
                    break
            if maker_row_updated:
                maker_row = maker_row_updated

            table_end_row = maker_row - 3
            if table_end_row < header_row:
                table_end_row = header_row

            # 数据右边界
            data_end_col = 1
            for c in range(1, last_col + 1):
                if ws.range((header_row, c)).value is not None:
                    data_end_col = c

            # 二次清空（仅在配置允许时，删除列后行号可能变化）
            if config.get("clear_sign_rows", True):
                for clear_row in (maker_row - 1, maker_row - 2):
                    if clear_row >= 1:
                        ws.range((clear_row, 1), (clear_row, last_col)).value = None

            # 5. 页面设置
            ws.api.PageSetup.LeftMargin = 2 * CM_TO_POINTS
            ws.api.PageSetup.PrintTitleRows = "$1:$5"
            ws.api.PageSetup.Zoom = False
            ws.api.PageSetup.FitToPagesWide = 1
            ws.api.PageSetup.FitToPagesTall = False

            # 6. 第一行合并 A1 到数据列最后一列（空格分隔）
            if data_end_col >= 1:
                merge_rng = ws.range((1, 1), (1, data_end_col))
                self.merge_with_content(merge_rng, sep=' ')
                self.log_insert(f"第一行合并 A1 到 {self.index_to_column_letter(data_end_col)}，保留所有内容")

            # 7. 第二行合并 A~G（无分隔符）
            if last_col >= 7:
                merge_rng = ws.range('A2:G2')
                self.merge_with_content(merge_rng, sep='')
                merge_rng.api.HorizontalAlignment = -4131  # 左对齐
                self.log_insert("第二行合并 A~G，左对齐，内容直接拼接")
            else:
                self.log_insert("提示：列数不足7，无法合并A2:G2")

            # 8. 行高
            ws.api.Rows(1).RowHeight = 60
            ws.api.Rows(2).RowHeight = 20

            # 9. 表格外框
            table_range = ws.range((header_row, 1), (table_end_row, data_end_col))
            table_range.api.BorderAround(Weight=4)
            self.log_insert(f"表格粗外框：行{header_row}~{table_end_row}，列1~{data_end_col}")

            # 10. 列头加粗 (3~5)
            if header_row == 3:
                ws.range((3, 1), (5, data_end_col)).api.Font.Bold = True
                self.log_insert("第3~5行列头已加粗")
            else:
                ws.range((header_row, 1), (header_row, data_end_col)).api.Font.Bold = True

            # 11. 列宽设置
            id_col = self.find_column_by_header(ws, "身份证", header_row, last_col)
            if id_col is None:
                for c in range(1, last_col + 1):
                    val = ws.range((header_row, c)).value
                    if val and "身份证" in str(val):
                        id_col = c
                        break
            for c in range(1, last_col + 1):
                if id_col and c == id_col:
                    ws.range((1, c), (last_row, c)).column_width = 20
                else:
                    ws.range((1, c), (last_row, c)).column_width = 9

            # 12. 数据区域字号10，行高20
            data_rng = ws.range((header_row + 1, 1), (table_end_row, data_end_col))
            data_rng.api.Font.Size = 10
            for r in range(header_row + 1, table_end_row + 1):
                ws.api.Rows(r).RowHeight = 20

            # 13. 处理"制表人："行
            maker_col = None
            maker_val = None
            for c in range(1, last_col + 1):
                val = ws.range((maker_row, c)).value
                if val and isinstance(val, str) and "制表人：" in val:
                    maker_col = c
                    maker_val = val
                    break
            if maker_col:
                for c in range(1, last_col + 1):
                    if c != maker_col:
                        ws.range((maker_row, c)).value = None
                if maker_col > 1:
                    content = str(maker_val).strip()
                    merge_range = ws.range((maker_row, 1), (maker_row, maker_col))
                    merge_range.merge()
                    merge_range.value = content
                    merge_range.api.HorizontalAlignment = -4152  # 右对齐
                    self.log_insert("制表人内容已向左合并至A列，右对齐")
                else:
                    ws.range((maker_row, maker_col)).api.HorizontalAlignment = -4152
                    self.log_insert("制表人列在最左，无需合并")

            # 14. 部长签字替换（兜底）
            footer_data = ensure_2d_list(ws.range((max(1, last_row - 2), 1), (last_row, last_col)).value)
            found_sign = False
            for r_offset, row_data in enumerate(footer_data):
                for c_offset, cell_val in enumerate(row_data):
                    if cell_val and isinstance(cell_val, str) and "部长、分管副总签字：" in cell_val:
                        new_val = cell_val.replace("部长、分管副总签字：", "部长签字：")
                        cell = ws.range((max(1, last_row - 2) + r_offset, 1 + c_offset))
                        cell.value = new_val
                        cell.api.HorizontalAlignment = -4108
                        self.log_insert(f"已修改签字文本：{new_val}")
                        found_sign = True
                        break
                if found_sign:
                    break
            if not found_sign:
                self.log_insert("提示：未找到'部长、分管副总签字：'文本（可能已被清空）")

            # 15. 业务审核删除
            keywords = ["社保基数", "医保基数", "工伤基数", "公积金基数"]
            header_vals = [ws.range((3, c)).value for c in range(1, min(last_col + 1, 51))]
            col_indices = [c + 1 for c, v in enumerate(header_vals)
                           if v and any(kw in str(v) for kw in keywords)]
            if len(col_indices) == 4:
                col_data = [[ws.range((r, c)).value for c in col_indices] for r in range(4, last_row + 1)]
                all_zero = all(
                    cell is None or (isinstance(cell, (int, float)) and cell == 0)
                    for row in col_data for cell in row
                )
                if all_zero:
                    for r_offset, row_data in enumerate(footer_data):
                        for c_offset, cell_val in enumerate(row_data):
                            if cell_val and "业务审核：" in str(cell_val):
                                cell = ws.range((max(1, last_row - 2) + r_offset, 1 + c_offset))
                                cell.value = str(cell_val).replace("业务审核：", "")
                                self.log_insert("已删除'业务审核：'字样")
                                break
                        else:
                            continue
                        break

            # 16. 保存文件
            if os.path.exists(new_file_path):
                base, ext = os.path.splitext(new_file_path)
                counter = 1
                while os.path.exists(f"{base}_{counter}{ext}"):
                    counter += 1
                new_file_path = f"{base}_{counter}{ext}"
                self.log_insert(f"文件已存在，另存为：{os.path.basename(new_file_path)}")
            wb.save(new_file_path)
            self.log_insert(f"处理完成，已保存为：{os.path.basename(new_file_path)}")

        except Exception as e:
            self.log_insert(f"处理失败：{str(e)}")
            raise
        finally:
            if wb:
                try:
                    wb.close()
                except:
                    pass


# ---------- 主入口 ----------
if __name__ == "__main__":
    root = tkinterdnd2.Tk()
    app = App(root)
    root.mainloop()