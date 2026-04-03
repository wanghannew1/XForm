import os
import re
import xlwings as xw
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
from tkinter import ttk
import threading
import tkinterdnd2
from tkinterdnd2 import DND_FILES

CM_TO_POINTS = 28.35
YEAR_MONTH_PATTERN = re.compile(r'(\d{4})年(\d{2})月')

HIDE_COLUMNS_CONFIG = {
    "header_row": 3,
    "mode": 1,
    "mode1_columns": ["C", "D"],
    "mode1_headers": ["部门", "岗位"],
    "mode2_columns": ["G", "H"],
    "mode2_headers": ["员工号", "应纳税所得额"]
}

class App:
    def __init__(self, root):
        self.root = root
        self.tkdnd_available = True
        self.root.title("工资表打印格式批量调整工具")
        self.root.geometry("700x800")
        self.file_list = []

        tk.Label(root, text="待处理文件列表：(支持拖拽文件至此)", font=("微软雅黑", 10)).pack(anchor='w', padx=10,
                                                                                             pady=5)
        self.listbox = tk.Listbox(root, selectmode=tk.EXTENDED, height=8)
        self.listbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.listbox.drop_target_register(DND_FILES)
        self.listbox.dnd_bind('<<Drop>>', self.on_drop)

        btn_frame = tk.Frame(root)
        btn_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Button(btn_frame, text="添加文件", command=self.add_files).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="清空列表", command=self.clear_list).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="开始处理", command=self.start_process).pack(side=tk.RIGHT, padx=5)

        self.progress = ttk.Progressbar(root, orient='horizontal', length=400, mode='determinate')
        self.progress.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(root, text="处理日志：", font=("微软雅黑", 10)).pack(anchor='w', padx=10)
        self.log = scrolledtext.ScrolledText(root, height=12, state='normal')
        self.log.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.log_insert("拖拽功能已启用")

        hide_frame = tk.LabelFrame(root, text="隐藏列配置", padx=10, pady=5)
        hide_frame.pack(fill=tk.X, padx=10, pady=5)

        row1_frame = tk.Frame(hide_frame)
        row1_frame.pack(fill=tk.X)
        tk.Label(row1_frame, text="列头行号：").pack(side=tk.LEFT)
        self.header_row_var = tk.StringVar(value=str(HIDE_COLUMNS_CONFIG["header_row"]))
        tk.Entry(row1_frame, textvariable=self.header_row_var, width=5).pack(side=tk.LEFT, padx=(0, 20))

        self.hide_mode_var = tk.IntVar(value=HIDE_COLUMNS_CONFIG["mode"])
        tk.Radiobutton(row1_frame, text="模式1：直接隐藏", variable=self.hide_mode_var, value=1).pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(row1_frame, text="模式2：无值则隐藏", variable=self.hide_mode_var, value=2).pack(side=tk.LEFT, padx=5)

        mode1_frame = tk.LabelFrame(hide_frame, text="模式1配置", padx=5, pady=5)
        mode1_frame.pack(fill=tk.X, pady=5)
        tk.Label(mode1_frame, text="列地址(逗号分隔)：").pack(side=tk.LEFT)
        self.mode1_columns_var = tk.StringVar(value=",".join(HIDE_COLUMNS_CONFIG["mode1_columns"]))
        tk.Entry(mode1_frame, textvariable=self.mode1_columns_var, width=20).pack(side=tk.LEFT, padx=5)
        tk.Label(mode1_frame, text="列头名称(逗号分隔)：").pack(side=tk.LEFT)
        self.mode1_headers_var = tk.StringVar(value=",".join(HIDE_COLUMNS_CONFIG["mode1_headers"]))
        tk.Entry(mode1_frame, textvariable=self.mode1_headers_var, width=25).pack(side=tk.LEFT, padx=5)

        mode2_frame = tk.LabelFrame(hide_frame, text="模式2配置", padx=5, pady=5)
        mode2_frame.pack(fill=tk.X, pady=5)
        tk.Label(mode2_frame, text="列地址(逗号分隔)：").pack(side=tk.LEFT)
        self.mode2_columns_var = tk.StringVar(value=",".join(HIDE_COLUMNS_CONFIG["mode2_columns"]))
        tk.Entry(mode2_frame, textvariable=self.mode2_columns_var, width=20).pack(side=tk.LEFT, padx=5)
        tk.Label(mode2_frame, text="列头名称(逗号分隔)：").pack(side=tk.LEFT)
        self.mode2_headers_var = tk.StringVar(value=",".join(HIDE_COLUMNS_CONFIG["mode2_headers"]))
        tk.Entry(mode2_frame, textvariable=self.mode2_headers_var, width=25).pack(side=tk.LEFT, padx=5)

    def on_drop(self, event):
        files = self.root.tk.splitlist(event.data)
        for file_path in files:
            file_path = file_path.strip()
            if file_path.startswith('{') and file_path.endswith('}'):
                file_path = file_path[1:-1]
            if file_path and (file_path.endswith('.xls') or file_path.endswith('.xlsx')):
                if os.path.exists(file_path):
                    if file_path not in self.file_list:
                        self.file_list.append(file_path)
                        self.listbox.insert(tk.END, os.path.basename(file_path))
                        self.log_insert(f"已添加：{os.path.basename(file_path)}")
        return 'action'

    def remove_selected(self):
        selection = self.listbox.curselection()
        for i in reversed(selection):
            self.file_list.pop(i)
            self.listbox.delete(i)

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

    def get_hide_config(self):
        config = {
            "header_row": int(self.header_row_var.get()),
            "mode": self.hide_mode_var.get(),
            "mode1_columns": [c.strip() for c in self.mode1_columns_var.get().split(",") if c.strip()],
            "mode1_headers": [h.strip() for h in self.mode1_headers_var.get().split(",") if h.strip()],
            "mode2_columns": [c.strip() for c in self.mode2_columns_var.get().split(",") if c.strip()],
            "mode2_headers": [h.strip() for h in self.mode2_headers_var.get().split(",") if h.strip()]
        }
        return config

    def column_letter_to_index(self, letter):
        result = 0
        for char in letter.upper():
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result

    def find_column_by_header(self, ws, header_name, header_row, last_col):
        for col in range(1, last_col + 1):
            cell_value = ws.range((header_row, col)).value
            if cell_value is not None and str(cell_value).strip() == header_name:
                return col
        return None

    def is_column_empty(self, ws, col_idx, start_row, end_row):
        for row in range(start_row, end_row + 1):
            val = ws.range((row, col_idx)).value
            if val is not None:
                return False
        return True

    def process_hide_columns(self, ws, config, last_row):
        mode = config["mode"]
        header_row = config["header_row"]
        last_col = ws.used_range.last_cell.column

        if mode == 1:
            columns = config["mode1_columns"]
            headers = config["mode1_headers"]
        else:
            columns = config["mode2_columns"]
            headers = config["mode2_headers"]

        col_indices = []
        for col_letter in columns:
            col_idx = self.column_letter_to_index(col_letter)
            if 1 <= col_idx <= last_col:
                col_indices.append(col_idx)

        for header in headers:
            col_idx = self.find_column_by_header(ws, header, header_row, last_col)
            if col_idx and col_idx not in col_indices:
                col_indices.append(col_idx)

        hidden_count = 0
        for col_idx in col_indices:
            if mode == 1:
                ws.api.Columns(col_idx).ColumnWidth = 0
                hidden_count += 1
                self.log_insert(f"已隐藏列（直接）：{col_idx}")
            else:
                if self.is_column_empty(ws, col_idx, header_row + 1, last_row):
                    ws.api.Columns(col_idx).ColumnWidth = 0
                    hidden_count += 1
                    self.log_insert(f"已隐藏列（空值）：{col_idx}")
                else:
                    self.log_insert(f"保留列（有值）：{col_idx}")

        return hidden_count

    def start_process(self):
        if not self.file_list:
            messagebox.showwarning("提示", "请先添加文件！")
            return
        self.disable_buttons()
        self.progress['value'] = 0
        self.log.delete(1.0, tk.END)
        threading.Thread(target=self.process_files, daemon=True).start()

    def disable_buttons(self):
        for child in self.root.winfo_children():
            if isinstance(child, tk.Frame):
                for btn in child.winfo_children():
                    if isinstance(btn, tk.Button):
                        btn.config(state='disabled')

    def enable_buttons(self):
        for child in self.root.winfo_children():
            if isinstance(child, tk.Frame):
                for btn in child.winfo_children():
                    if isinstance(btn, tk.Button):
                        btn.config(state='normal')

    def process_files(self):
        total = len(self.file_list)
        for idx, file_path in enumerate(self.file_list):
            self.log_insert(f"正在处理：{os.path.basename(file_path)}")
            try:
                self.process_one_file(file_path)
                self.log_insert(f"处理成功：{os.path.basename(file_path)}")
            except Exception as e:
                self.log_insert(f"处理失败：{os.path.basename(file_path)} - {str(e)}")
            progress = int((idx + 1) / total * 100)
            self.root.after(0, lambda p=progress: self.progress.configure(value=p))
        self.root.after(0, lambda: messagebox.showinfo("完成", "全部处理完成！"))
        self.root.after(0, self.enable_buttons)

    def log_insert(self, msg):
        self.root.after(0, lambda: self.log.insert(tk.END, msg + "\n"))
        self.root.after(0, lambda: self.log.see(tk.END))

    def process_one_file(self, file_path):
        app = xw.App(visible=False, add_book=False)
        try:
            wb = app.books.open(file_path)
            ws = wb.sheets[0]

            used = ws.used_range
            last_row = used.last_cell.row
            last_col = used.last_cell.column

            row2_colB = ws.range('B2').value
            name_part = str(row2_colB).replace("名称：", "").strip() if row2_colB else "未知名称"

            row1_colC = ws.range('C1').value
            year_month = ""
            if row1_colC:
                match = YEAR_MONTH_PATTERN.search(str(row1_colC))
                if match:
                    year_month = f"{match.group(1)}年{match.group(2)}月"

            new_filename = f"{name_part}{year_month}工资表.xlsx"
            original_dir = os.path.dirname(file_path)
            new_file_path = os.path.join(original_dir, new_filename)

            self.log_insert(f"准备另存为：{new_filename}")

            ws.api.PageSetup.LeftMargin = 2 * CM_TO_POINTS
            ws.api.PageSetup.PrintTitleRows = "$1:$5"
            ws.api.PageSetup.Zoom = False
            ws.api.PageSetup.FitToPagesWide = 1
            ws.api.PageSetup.FitToPagesTall = False

            hide_config = self.get_hide_config()
            self.process_hide_columns(ws, hide_config, last_row)

            search_start = max(1, last_row - 2)
            footer_data = ws.range((search_start, 1), (last_row, last_col)).value

            found = False
            for r_idx, row in enumerate(footer_data):
                for c_idx, val in enumerate(row):
                    if val and isinstance(val, str):
                        if "部长、分管副总签字：" in val:
                            new_val = val.replace("部长、分管副总签字：", "部长签字：")
                            cell = ws.range((search_start + r_idx, 1 + c_idx))
                            cell.value = new_val
                            cell.api.HorizontalAlignment = -4108
                            self.log_insert(f"已修改签字文本：{new_val}，单元格内居中对齐")
                            found = True
                            break
                if found:
                    break

            if not found:
                self.log_insert("提示：未找到'部长、分管副总签字：'文本")

            found_table_maker = False
            for r_idx, row in enumerate(footer_data):
                for c_idx, val in enumerate(row):
                    if val and isinstance(val, str):
                        if "制表人：" in val:
                            ws.range((search_start + r_idx, 1 + c_idx)).api.HorizontalAlignment = -4152
                            self.log_insert(f"已设置'{val}'单元格居右对齐")
                            found_table_maker = True
                            break
                if found_table_maker:
                    break

            if not found_table_maker:
                self.log_insert("提示：未找到'制表人：'文本")

            keywords = ["社保基数", "医保基数", "工伤基数", "公积金基数"]
            header_values = [ws.range((3, c)).value for c in range(1, min(last_col + 1, 51))]
            col_indices = [c + 1 for c, v in enumerate(header_values)
                          if v and any(kw in str(v) for kw in keywords)]

            if len(col_indices) != 4:
                self.log_insert(f"警告：未找到全部四列基数，找到 {len(col_indices)} 列")
            else:
                col_data = [[ws.range((r, c)).value for c in col_indices] for r in range(4, last_row + 1)]
                all_zero = all(
                    cell is None or (isinstance(cell, (int, float)) and cell == 0)
                    for row in col_data for cell in row
                )
                if all_zero:
                    removed = False
                    for r_idx, row in enumerate(footer_data):
                        for c_idx, val in enumerate(row):
                            if val and "业务审核：" in str(val):
                                ws.range((search_start + r_idx, 1 + c_idx)).value = str(val).replace("业务审核：", "")
                                removed = True
                                self.log_insert("已删除'业务审核：'字样（条件满足）")
                                break
                        if removed:
                            break
                    if not removed:
                        self.log_insert("警告：未找到'业务审核：'文本")

            if os.path.exists(new_file_path):
                base, ext = os.path.splitext(new_file_path)
                counter = 1
                while os.path.exists(f"{base}_{counter}{ext}"):
                    counter += 1
                new_file_path = f"{base}_{counter}{ext}"
                self.log_insert(f"文件已存在，另存为：{os.path.basename(new_file_path)}")

            wb.save(new_file_path)
            wb.close()

            self.log_insert(f"处理完成，已保存为：{os.path.basename(new_file_path)}")

        except Exception as e:
            self.log_insert(f"处理失败：{str(e)}")
            raise
        finally:
            app.quit()


if __name__ == "__main__":
    root = tkinterdnd2.Tk()
    app = App(root)
    root.mainloop()
