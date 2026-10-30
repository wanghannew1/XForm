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
    "mode1_headers": [],
    "mode2_columns": [],
    "mode2_headers": ["部门", "职工号", "交通补贴"],
    "clear_sign_rows": False,
    "left_margin_cm": 1.5,
    "right_margin_cm": 1.5,
    "text_replacements": [
        {"find": "部长、分管副总签字：", "replace": "分管领导审核："}
    ],
    "beautify_table": False
}

def load_config():
    config = HIDE_COLUMNS_CONFIG.copy()
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                saved = json.load(f)
                config.update(saved)
        except:
            pass
    return config

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
        self.root.geometry("700x980")
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
        self.clear_sign_rows_var = tk.BooleanVar(value=self.config.get("clear_sign_rows", False))
        tk.Checkbutton(sign_frame, text="清空签名行（合计行与制表人之间的总经理签字等提示词）",
                       variable=self.clear_sign_rows_var, onvalue=True, offvalue=False).pack(anchor=tk.W)

        # ---------- 表格美化选项 ----------
        beauty_frame = tk.LabelFrame(root, text="表格美化", padx=10, pady=5)
        beauty_frame.pack(fill=tk.X, padx=10, pady=5)
        self.beautify_table_var = tk.BooleanVar(value=self.config.get("beautify_table", False))
        tk.Checkbutton(beauty_frame, text="美化表格（边框加粗、列头加粗、数据区域字号与行高设置）",
                       variable=self.beautify_table_var, onvalue=True, offvalue=False).pack(anchor=tk.W)

        # ---------- 页边距配置 ----------
        margin_frame = tk.LabelFrame(root, text="页边距设置", padx=10, pady=5)
        margin_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(margin_frame, text="左边距(cm)：").pack(side=tk.LEFT)
        self.left_margin_var = tk.StringVar(value=str(self.config.get("left_margin_cm", 1.5)))
        tk.Entry(margin_frame, textvariable=self.left_margin_var, width=5).pack(side=tk.LEFT, padx=5)
        tk.Label(margin_frame, text="右边距(cm)：").pack(side=tk.LEFT, padx=(15, 0))
        self.right_margin_var = tk.StringVar(value=str(self.config.get("right_margin_cm", 1.5)))
        tk.Entry(margin_frame, textvariable=self.right_margin_var, width=5).pack(side=tk.LEFT, padx=5)

        # ---------- 文本替换配置 ----------
        replace_frame = tk.LabelFrame(root, text="文本替换配置", padx=10, pady=5)
        replace_frame.pack(fill=tk.X, padx=10, pady=5)

        self.replacements = list(self.config.get("text_replacements", []))

        self.replace_listbox = tk.Listbox(replace_frame, height=4)
        self.replace_listbox.pack(fill=tk.X, pady=5)

        replace_btn_frame = tk.Frame(replace_frame)
        replace_btn_frame.pack(fill=tk.X)
        tk.Button(replace_btn_frame, text="添加", command=self.add_replacement).pack(side=tk.LEFT, padx=2)
        tk.Button(replace_btn_frame, text="编辑", command=self.edit_replacement).pack(side=tk.LEFT, padx=2)
        tk.Button(replace_btn_frame, text="删除", command=self.delete_replacement).pack(side=tk.LEFT, padx=2)

        self.refresh_replace_list()

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
            "clear_sign_rows": self.clear_sign_rows_var.get(),
            "left_margin_cm": float(self.left_margin_var.get()),
            "right_margin_cm": float(self.right_margin_var.get()),
            "text_replacements": self.replacements,
            "beautify_table": self.beautify_table_var.get(),
        }
        if self.mode1_enabled.get():
            config["enabled_modes"].append("mode1")
        if self.mode2_enabled.get():
            config["enabled_modes"].append("mode2")
        return config

    # ---------- 文本替换规则编辑 ----------
    def refresh_replace_list(self):
        self.replace_listbox.delete(0, tk.END)
        for rule in self.replacements:
            self.replace_listbox.insert(tk.END, f"{rule['find']} → {rule['replace']}")

    def add_replacement(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("添加替换规则")
        dialog.geometry("450x150")
        dialog.transient(self.root)
        dialog.grab_set()

        tk.Label(dialog, text="查找文字：").grid(row=0, column=0, sticky=tk.W, padx=10, pady=10)
        find_entry = tk.Entry(dialog, width=35)
        find_entry.grid(row=0, column=1, sticky=tk.W, padx=10)

        tk.Label(dialog, text="替换为：").grid(row=1, column=0, sticky=tk.W, padx=10, pady=10)
        replace_entry = tk.Entry(dialog, width=35)
        replace_entry.grid(row=1, column=1, sticky=tk.W, padx=10)

        def save():
            find_text = find_entry.get().strip()
            replace_text = replace_entry.get().strip()
            if find_text:
                self.replacements.append({"find": find_text, "replace": replace_text})
                self.refresh_replace_list()
            dialog.destroy()

        btn_frame = tk.Frame(dialog)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=10)
        tk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="确定", command=save).pack(side=tk.LEFT, padx=5)

    def edit_replacement(self):
        selection = self.replace_listbox.curselection()
        if not selection:
            return
        idx = selection[0]
        rule = self.replacements[idx]

        dialog = tk.Toplevel(self.root)
        dialog.title("编辑替换规则")
        dialog.geometry("450x150")
        dialog.transient(self.root)
        dialog.grab_set()

        tk.Label(dialog, text="查找文字：").grid(row=0, column=0, sticky=tk.W, padx=10, pady=10)
        find_entry = tk.Entry(dialog, width=35)
        find_entry.insert(0, rule["find"])
        find_entry.grid(row=0, column=1, sticky=tk.W, padx=10)

        tk.Label(dialog, text="替换为：").grid(row=1, column=0, sticky=tk.W, padx=10, pady=10)
        replace_entry = tk.Entry(dialog, width=35)
        replace_entry.insert(0, rule["replace"])
        replace_entry.grid(row=1, column=1, sticky=tk.W, padx=10)

        def save():
            find_text = find_entry.get().strip()
            replace_text = replace_entry.get().strip()
            if find_text:
                self.replacements[idx] = {"find": find_text, "replace": replace_text}
                self.refresh_replace_list()
            dialog.destroy()

        btn_frame = tk.Frame(dialog)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=10)
        tk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="确定", command=save).pack(side=tk.LEFT, padx=5)

    def delete_replacement(self):
        selection = self.replace_listbox.curselection()
        if not selection:
            return
        idx = selection[0]
        del self.replacements[idx]
        self.refresh_replace_list()

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

    def set_merged_cell_value(self, ws, row, col_idx, val):
        """向指定行/列写入值；若该单元格位于合并区内，写入合并区左上角单元格"""
        rng_api = ws.range((row, col_idx)).api
        if rng_api.MergeCells:
            rng_api.MergeArea.Cells(1, 1).Value = val
        else:
            rng_api.Value = val

    def get_columns_to_delete(self, ws, config, last_col):
        """根据处理配置返回需要真实删除的列号列表（不再按空值过滤）"""
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
                if 1 <= col_idx <= last_col:
                    delete_list.append(col_idx)
            for header in config["mode2_headers"]:
                col_idx = self.find_column_by_header(ws, header, header_row, last_col)
                if col_idx and col_idx not in delete_list:
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

            # 3. 清空倒数第2、3行
            clear_enabled = config.get("clear_sign_rows", False)
            self.log_insert(f"[配置] clear_sign_rows={clear_enabled}")
            if clear_enabled:
                for clear_row in (maker_row - 1, maker_row - 2):
                    if clear_row >= 1:
                        ws.range((clear_row, 1), (clear_row, last_col)).value = None
                        self.log_insert(f"已清空第{clear_row}行")

            # 4. 真实删除整列（公式引用随删除自动调整，不会引用错列）
            # 合计行及其以下的内容（如合计行锚点上的合计值、签名行）先快照，
            # 删除后按 target = c - count(deleted < c) 映射写回（即右移到右侧列中保留）。
            # 注意：被删列若是合并区左上角锚点，WPS 引擎原生会丢弃锚点内容，
            # 因此合计行自身的合计值必须纳入快照，否则会随删列丢失。
            delete_cols = self.get_columns_to_delete(ws, config, last_col)
            hidden_cols = set()
            if delete_cols:
                col_to_delete = set(delete_cols)
                sorted_deleted = sorted(col_to_delete)
                preserve_start_row = table_end_row
                last_row_snapshot = ws.used_range.last_cell.row
                snapshots = []  # (row, col, value) 仅记录被删列上的值/合并区左上角锚点值
                for c in sorted_deleted:
                    for r in range(preserve_start_row, last_row_snapshot + 1):
                        rng_api = ws.range((r, c)).api
                        if rng_api.MergeCells and rng_api.MergeArea.Cells(1, 1).Address != rng_api.Address:
                            continue  # 合并区非左上角，其值随合并区整体保留
                        val = ws.range((r, c)).value
                        if val is not None:
                            snapshots.append((r, c, val))
                for col_idx in sorted(col_to_delete, reverse=True):
                    header_val = ws.range((header_row, col_idx)).value
                    col_name = str(header_val).strip() if header_val else self.index_to_column_letter(col_idx)
                    ws.api.Columns(col_idx).Delete()
                    self.log_insert(f"已删除列：{col_name}")
                restored = 0
                for row, col, val in snapshots:
                    target = col - sum(1 for d in sorted_deleted if d < col)
                    self.set_merged_cell_value(ws, row, target, val)
                    restored += 1
                self.log_insert(f"合计行及以下内容已按右移映射恢复 {restored} 处")
                # 删除后刷新使用区域边界
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
            if config.get("clear_sign_rows", False):
                for clear_row in (maker_row - 1, maker_row - 2):
                    if clear_row >= 1:
                        ws.range((clear_row, 1), (clear_row, last_col)).value = None

            # 5. 页面设置
            ws.api.PageSetup.LeftMargin = config.get("left_margin_cm", 1.5) * CM_TO_POINTS
            ws.api.PageSetup.RightMargin = config.get("right_margin_cm", 1.5) * CM_TO_POINTS
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

            beautify = config.get("beautify_table", False)

            # 9. 表格外框（美化选项）
            if beautify:
                table_range = ws.range((header_row, 1), (table_end_row, data_end_col))
                table_range.api.BorderAround(Weight=4)
                self.log_insert(f"表格粗外框：行{header_row}~{table_end_row}，列1~{data_end_col}")
            else:
                self.log_insert("未启用表格美化，跳过外框加粗")

            # 10. 列头加粗（美化选项）
            if beautify:
                if header_row == 3:
                    ws.range((3, 1), (5, data_end_col)).api.Font.Bold = True
                    self.log_insert("第3~5行列头已加粗")
                else:
                    ws.range((header_row, 1), (header_row, data_end_col)).api.Font.Bold = True
            else:
                self.log_insert("未启用表格美化，跳过列头加粗")

            # 11. 列宽设置（跳过已隐藏的列，避免自动解除隐藏）
            id_col = self.find_column_by_header(ws, "身份证", header_row, last_col)
            if id_col is None:
                for c in range(1, last_col + 1):
                    val = ws.range((header_row, c)).value
                    if val and "身份证" in str(val):
                        id_col = c
                        break
            for c in range(1, last_col + 1):
                if c in hidden_cols:
                    continue
                if id_col and c == id_col:
                    ws.range((1, c), (last_row, c)).column_width = 20
                else:
                    ws.range((1, c), (last_row, c)).column_width = 9

            # 12. 数据区域字号10，行高20（美化选项）
            if beautify:
                data_rng = ws.range((header_row + 1, 1), (table_end_row, data_end_col))
                data_rng.api.Font.Size = 10
                for r in range(header_row + 1, table_end_row + 1):
                    ws.api.Rows(r).RowHeight = 20
            else:
                self.log_insert("未启用表格美化，跳过数据区域字号与行高设置")

            # 13. 处理"制表人："行（仅当 clear_sign_rows 为 True 时合并清空，否则保留同行的其他签名文字）
            maker_col = None
            maker_val = None
            for c in range(1, last_col + 1):
                val = ws.range((maker_row, c)).value
                if val and isinstance(val, str) and "制表人：" in val:
                    maker_col = c
                    maker_val = val
                    break
            if maker_col:
                if config.get("clear_sign_rows", False):
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
                else:
                    ws.range((maker_row, maker_col)).api.HorizontalAlignment = -4152
                    self.log_insert("clear_sign_rows=False，保持制表人行原样，制表人单元格右对齐")

            # 14. 文本替换（从配置读取替换规则）
            replacements = config.get("text_replacements", [])
            footer_data = ensure_2d_list(ws.range((max(1, last_row - 2), 1), (last_row, last_col)).value)
            for rule in replacements:
                find_text = rule.get("find", "")
                replace_text = rule.get("replace", "")
                if not find_text:
                    continue
                found = False
                for r_offset, row_data in enumerate(footer_data):
                    for c_offset, cell_val in enumerate(row_data):
                        if cell_val and isinstance(cell_val, str) and find_text in cell_val:
                            new_val = cell_val.replace(find_text, replace_text)
                            cell = ws.range((max(1, last_row - 2) + r_offset, 1 + c_offset))
                            cell.value = new_val
                            cell.api.HorizontalAlignment = -4108
                            self.log_insert(f"文本替换：'{find_text}' → '{replace_text}'")
                            found = True
                            break
                    if found:
                        break
                if not found:
                    self.log_insert(f"提示：未找到'{find_text}'文本（可能已被清空）")

            # 15. 保存文件
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