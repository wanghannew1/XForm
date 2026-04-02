import os
import xlwings as xw
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
from tkinter import ttk
import threading
import tkinterDnD
from tkinterdnd2 import DND_FILES

CM_TO_POINTS = 28.35

class App:
    def __init__(self, root):
        self.root = root
        self.tkdnd_available = True
        self.root.title("工资表打印格式批量调整工具")
        self.root.geometry("700x500")
        self.file_list = []

        # 文件列表区域
        tk.Label(root, text="待处理文件列表：(支持拖拽文件至此)", font=("微软雅黑", 10)).pack(anchor='w', padx=10,
                                                                                             pady=5)
        self.listbox = tk.Listbox(root, selectmode=tk.EXTENDED, height=8)
        self.listbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # 注册拖拽目标
        self.listbox.drop_target_register(DND_FILES)
        self.listbox.dnd_bind('<<Drop>>', self.on_drop)

        # 按钮区域
        btn_frame = tk.Frame(root)
        btn_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Button(btn_frame, text="添加文件", command=self.add_files).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="清空列表", command=self.clear_list).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="开始处理", command=self.start_process).pack(side=tk.RIGHT, padx=5)

        # 进度条
        self.progress = ttk.Progressbar(root, orient='horizontal', length=400, mode='determinate')
        self.progress.pack(fill=tk.X, padx=10, pady=5)

        # 日志区域
        tk.Label(root, text="处理日志：", font=("微软雅黑", 10)).pack(anchor='w', padx=10)
        self.log = scrolledtext.ScrolledText(root, height=12, state='normal')
        self.log.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # 现在可以安全地输出日志
        self.log_insert("拖拽功能已启用")

    def on_drop(self, event):
        """处理拖拽放置事件"""
        # event.data 是一个用空格或换行分隔的文件路径字符串
        files = self.root.tk.splitlist(event.data)
        for file_path in files:
            file_path = file_path.strip()
            # 去除可能的 {} 包裹（路径包含空格时会有）
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
        """删除选中的文件"""
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

    def start_process(self):
        if not self.file_list:
            messagebox.showwarning("提示", "请先添加文件！")
            return
        # 禁用按钮，防止重复点击
        self.disable_buttons()
        self.progress['value'] = 0
        self.log.delete(1.0, tk.END)
        # 启动后台线程处理
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
        """处理单个 Excel 文件"""
        app = xw.App(visible=False, add_book=False)
        try:
            wb = app.books.open(file_path)
            ws = wb.sheets[0]

            # 0. 获取新文件名
            # 取第二行 B 列内容
            row2_colB = ws.range('B2').value
            if row2_colB:
                name_part = str(row2_colB).strip()
                # 过滤掉"名称："字样
                name_part = name_part.replace("名称：", "").strip()
            else:
                name_part = "未知名称"

            # 取第一行 A 列内容中的年月
            row1_colC = ws.range('C1').value
            year_month = ""
            if row1_colC:
                import re
                match = re.search(r'(\d{4})年(\d{2})月', str(row1_colC))
                if match:
                    year_month = f"{match.group(1)}年{match.group(2)}月"

            # 构建新文件名
            new_filename = f"{name_part}{year_month}工资表.xlsx"
            original_dir = os.path.dirname(file_path)
            new_file_path = os.path.join(original_dir, new_filename)

            self.log_insert(f"准备另存为：{new_filename}")

            # 1. 设置左侧页边距 2cm
            ws.api.PageSetup.LeftMargin = 2 * CM_TO_POINTS

            # 2. 打印标题行 1-5
            ws.api.PageSetup.PrintTitleRows = "$1:$5"

            # 3. 缩放：全部列在一页
            ws.api.PageSetup.Zoom = False
            ws.api.PageSetup.FitToPagesWide = 1
            ws.api.PageSetup.FitToPagesTall = False

            # 4. 隐藏 D 列和 E 列
            ws.range('D:D').column_width = 0
            ws.range('E:E').column_width = 0

            # 5. 修改签字文本
            used_range = ws.used_range
            last_row = used_range.last_cell.row
            last_col = used_range.last_cell.column

            found = False
            start_row = max(1, last_row - 2)  # 倒数 3 行
            for row in range(start_row, last_row + 1):
                row_range = ws.range((row, 1), (row, last_col))
                for cell in row_range:
                    if cell.value and isinstance(cell.value, str):
                        if "部长、分管副总签字：" in cell.value:
                            new_val = cell.value.replace("部长、分管副总签字：", "部长签字：")
                            cell.value = new_val
                            cell.api.HorizontalAlignment = -4108  # xlCenter 居中常量
                            self.log_insert(f"已修改签字文本：{new_val}，单元格内居中对齐")
                            found = True
                            break
                    if found:
                        break

            if not found:
                self.log_insert("提示：未找到'部长、分管副总签字：'文本")

            # 新增：查找并设置"制表人："单元格居右对齐
            found_table_maker = False
            for row in range(start_row, last_row + 1):
                row_range = ws.range((row, 1), (row, last_col))
                for cell in row_range:
                    if cell.value and isinstance(cell.value, str):
                        if "制表人：" in cell.value:
                            cell.api.HorizontalAlignment = -4152  # xlRight 居右对齐常量
                            self.log_insert(f"已设置'{cell.value}'单元格居右对齐")
                            found_table_maker = True
                            break
                if found_table_maker:
                    break

            if not found_table_maker:
                self.log_insert("提示：未找到'制表人：'文本")

            # 6. 检查四列基数是否全 0，若是则删除"业务审核："
            header_row = 3
            keywords = ["社保基数", "医保基数", "工伤基数", "公积金基数"]
            col_indices = []
            for col in range(1, 51):
                cell_val = ws.range((header_row, col)).value
                if cell_val:
                    for kw in keywords:
                        if kw in str(cell_val):
                            col_indices.append(col)
                            break
            if len(col_indices) != 4:
                self.log_insert(f"警告：未找到全部四列基数，找到 {len(col_indices)} 列")
            else:
                all_zero = True
                for row in range(header_row + 1, last_row + 1):
                    for col in col_indices:
                        val = ws.range((row, col)).value
                        if val is not None:
                            try:
                                if float(val) != 0:
                                    all_zero = False
                                    break
                            except (ValueError, TypeError):
                                all_zero = False
                                break
                    if not all_zero:
                        break
                if all_zero:
                    removed = False
                    for row in range(last_row, max(1, last_row - 2), -1):  # 倒数 3 行
                        row_range = ws.range((row, 1), (row, last_col))
                        for cell in row_range:
                            if cell.value and "业务审核：" in str(cell.value):
                                new_val = str(cell.value).replace("业务审核：", "")
                                cell.value = new_val
                                removed = True
                                self.log_insert("已删除'业务审核：'字样（条件满足）")
                                break
                        if removed:
                            break
                    if not removed:
                        self.log_insert("警告：未找到'业务审核：'文本")

            # 关键修改：直接另存为新文件
            # 如果新文件已存在，自动添加序号避免覆盖
            if os.path.exists(new_file_path):
                base, ext = os.path.splitext(new_file_path)
                counter = 1
                while os.path.exists(f"{base}_{counter}{ext}"):
                    counter += 1
                new_file_path = f"{base}_{counter}{ext}"
                self.log_insert(f"文件已存在，另存为：{os.path.basename(new_file_path)}")

            wb.save(new_file_path)  # 直接保存到新路径
            wb.close()

            self.log_insert(f"处理完成，已保存为：{os.path.basename(new_file_path)}")

        except Exception as e:
            self.log_insert(f"处理失败：{str(e)}")
            raise
        finally:
            app.quit()


if __name__ == "__main__":
    # root = tk.Tk()
    root = tkinterDnD.Tk()
    app = App(root)
    root.mainloop()
