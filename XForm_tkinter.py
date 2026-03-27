import os
import xlwings as xw
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
from tkinter import ttk
import threading

CM_TO_POINTS = 28.35

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("工资表打印格式批量调整工具")
        self.root.geometry("700x500")
        self.file_list = []

        # 文件列表区域
        tk.Label(root, text="待处理文件列表：", font=("微软雅黑", 10)).pack(anchor='w', padx=10, pady=5)
        self.listbox = tk.Listbox(root, selectmode=tk.EXTENDED, height=8)
        self.listbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

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

    def add_files(self):
        files = filedialog.askopenfilenames(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xls *.xlsx")]
        )
        for f in files:
            if f not in self.file_list:
                self.file_list.append(f)
                self.listbox.insert(tk.END, f)

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
            self.log_insert(f"正在处理: {os.path.basename(file_path)}")
            try:
                self.process_one_file(file_path)
                self.log_insert(f"处理成功: {os.path.basename(file_path)}")
            except Exception as e:
                self.log_insert(f"处理失败: {os.path.basename(file_path)} - {str(e)}")
            progress = int((idx + 1) / total * 100)
            self.root.after(0, lambda p=progress: self.progress.configure(value=p))
        self.root.after(0, lambda: messagebox.showinfo("完成", "全部处理完成！"))
        self.root.after(0, self.enable_buttons)

    def log_insert(self, msg):
        self.root.after(0, lambda: self.log.insert(tk.END, msg + "\n"))
        self.root.after(0, lambda: self.log.see(tk.END))

    def process_one_file(self, file_path):
        """处理单个Excel文件（与之前相同）"""
        app = xw.App(visible=False, add_book=False)
        try:
            wb = app.books.open(file_path)
            ws = wb.sheets[0]

            # 1. 设置左侧页边距1cm
            ws.api.PageSetup.LeftMargin = 2 * CM_TO_POINTS

            # 2. 打印标题行1-5
            ws.api.PageSetup.PrintTitleRows = "$1:$5"

            # 3. 缩放：全部列在一页
            ws.api.PageSetup.Zoom = False
            ws.api.PageSetup.FitToPagesWide = 1
            ws.api.PageSetup.FitToPagesTall = False

            # 4. 隐藏D列和E列
            ws.range('D:D').column_width = 0
            ws.range('E:E').column_width = 0

            # 5. 修改签字文本
            used_range = ws.used_range
            last_row = used_range.last_cell.row
            found = False
            for r in range(max(1, last_row - 20), last_row + 1):
                cell = ws.range((r, 1))
                self.log_insert(f"最后一行内容：{cell.value}")
                print(cell.value)
                if cell.value and "部长、分管副总签字：" in str(cell.value):
                    new_val = str(cell.value).replace("部长、分管副总签字：", "部长签字：")
                    cell.value = new_val
                    found = True
                    break
            if not found:
                self.log_insert("警告：未找到‘部长、分管副总签字：’文本")

            # 6. 检查四列基数是否全0，若是则删除“业务审核：”
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
                    for r in range(last_row, max(1, last_row - 20), -1):
                        cell = ws.range((r, 1))
                        if cell.value and "业务审核：" in str(cell.value):
                            new_val = str(cell.value).replace("业务审核：", "")
                            cell.value = new_val
                            removed = True
                            self.log_insert("已删除‘业务审核：’字样（条件满足）")
                            break
                    if not removed:
                        self.log_insert("警告：未找到‘业务审核：’文本")

            wb.save()
            wb.close()
        finally:
            app.quit()

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()