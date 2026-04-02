import os
import xlwings as xw
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
from tkinter import ttk
import threading

# 单位转换：1 cm = 28.35 磅
CM_TO_POINTS = 28.35


class SalaryPrintFormatter:
    """工资表打印格式调整核心逻辑"""

    def __init__(self, log_callback=None):
        self.log_callback = log_callback  # 用于输出日志的回调函数

    def log(self, msg):
        if self.log_callback:
            self.log_callback(msg)
        else:
            print(msg)

    def process_file(self, file_path):
        """
        处理单个Excel文件
        """
        app = xw.App(visible=False, add_book=False)
        try:
            wb = app.books.open(file_path)
            # 处理第一个工作表（可根据需要修改）
            ws = wb.sheets[0]

            # 1. 设置左侧页边距 2cm（可根据实际需求调整）
            ws.api.PageSetup.LeftMargin = 2 * CM_TO_POINTS
            self.log(f"  已设置左侧页边距: 2cm")

            # 2. 设置打印标题行（第1~5行每页重复）
            ws.api.PageSetup.PrintTitleRows = "$1:$5"
            self.log(f"  已设置打印标题行: 1-5行")

            # 3. 设置缩放：所有列在一页宽
            ws.api.PageSetup.Zoom = False
            ws.api.PageSetup.FitToPagesWide = 1
            ws.api.PageSetup.FitToPagesTall = False
            self.log(f"  已设置缩放: 宽度缩放到1页")

            # 4. 隐藏 D 列和 E 列
            ws.range('D:D').column_width = 0
            ws.range('E:E').column_width = 0
            self.log(f"  已隐藏D列(部门)和E列(岗位)")

            # 获取使用区域的行列信息
            used_range = ws.used_range
            last_row = used_range.last_cell.row
            last_col = used_range.last_cell.column
            self.log(f"  数据范围: 行1-{last_row}, 列1-{last_col}")

            # 5. 修改签字文本：将“部长、分管副总签字：”改为“部长签字：”
            #    搜索范围：最后3行所有列
            found_sign = False
            start_row = max(1, last_row - 2)  # 最后3行起始行
            for row in range(start_row, last_row + 1):
                row_range = ws.range((row, 1), (row, last_col))
                for cell in row_range:
                    if cell.value and isinstance(cell.value, str):
                        if "部长、分管副总签字：" in cell.value:
                            new_val = cell.value.replace("部长、分管副总签字：", "部长签字：")
                            cell.value = new_val
                            self.log(f"  已修改签字文本: {new_val}")
                            found_sign = True
                            break
                if found_sign:
                    break
            if not found_sign:
                self.log("  警告: 未找到‘部长、分管副总签字：’文本")

            # 6. 检查四列基数是否全为0，若是则删除“业务审核：”
            header_row = 3
            keywords = ["社保基数", "医保基数", "工伤基数", "公积金基数"]
            col_indices = []
            # 查找四列所在列号（遍历第3行前50列）
            for col in range(1, 51):
                cell_val = ws.range((header_row, col)).value
                if cell_val and isinstance(cell_val, str):
                    for kw in keywords:
                        if kw in cell_val:
                            col_indices.append(col)
                            break
            if len(col_indices) != 4:
                self.log(f"  警告: 未找到全部四列基数，找到 {len(col_indices)} 列")
            else:
                # 检查从第4行到最后一行，这些列的值是否全为0或空
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
                                # 非数值，视为非零
                                all_zero = False
                                break
                    if not all_zero:
                        break

                if all_zero:
                    self.log("  检测到四列基数全为0，开始删除‘业务审核：’字样")
                    # 删除最后3行中的“业务审核：”
                    removed = False
                    for row in range(last_row, max(1, last_row - 2), -1):
                        row_range = ws.range((row, 1), (row, last_col))
                        for cell in row_range:
                            if cell.value and "业务审核：" in str(cell.value):
                                new_val = str(cell.value).replace("业务审核：", "")
                                cell.value = new_val
                                removed = True
                                self.log("  已删除‘业务审核：’字样（条件满足）")
                                break
                        if removed:
                            break
                    if not removed:
                        self.log("  警告: 未找到‘业务审核：’文本")
                else:
                    self.log("  四列基数不全为0，保留‘业务审核：’字样")

            # 保存文件
            wb.save()
            self.log(f"  处理完成，文件已保存")
            return True

        except Exception as e:
            self.log(f"  处理失败: {str(e)}")
            return False
        finally:
            app.quit()


class App:
    """主界面"""

    def __init__(self, root):
        self.root = root
        self.root.title("工资表打印格式批量调整工具")
        self.root.geometry("750x550")
        self.file_list = []
        self.processor = None

        # 文件列表区域
        tk.Label(root, text="待处理文件列表：", font=("微软雅黑", 10)).pack(anchor='w', padx=10, pady=5)
        self.listbox = tk.Listbox(root, selectmode=tk.EXTENDED, height=8)
        self.listbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # 按钮区域
        btn_frame = tk.Frame(root)
        btn_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Button(btn_frame, text="添加文件", command=self.add_files, width=12).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="清空列表", command=self.clear_list, width=12).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="开始处理", command=self.start_process, width=12).pack(side=tk.RIGHT, padx=5)

        # 进度条
        self.progress = ttk.Progressbar(root, orient='horizontal', length=400, mode='determinate')
        self.progress.pack(fill=tk.X, padx=10, pady=5)

        # 日志区域
        tk.Label(root, text="处理日志：", font=("微软雅黑", 10)).pack(anchor='w', padx=10)
        self.log = scrolledtext.ScrolledText(root, height=12, state='normal', font=("Consolas", 9))
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
        # 禁用按钮
        self.disable_buttons()
        self.progress['value'] = 0
        self.log.delete(1.0, tk.END)
        # 启动后台线程
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
            self.log_insert(f"\n>>> 正在处理: {os.path.basename(file_path)}")
            # 每个文件创建一个处理器实例
            processor = SalaryPrintFormatter(log_callback=self.log_insert)
            success = processor.process_file(file_path)
            if success:
                self.log_insert(f"处理成功: {os.path.basename(file_path)}")
            else:
                self.log_insert(f"处理失败: {os.path.basename(file_path)}")
            # 更新进度
            progress_val = int((idx + 1) / total * 100)
            self.root.after(0, lambda p=progress_val: self.progress.configure(value=p))
        self.root.after(0, lambda: messagebox.showinfo("完成", "全部处理完成！"))
        self.root.after(0, self.enable_buttons)

    def log_insert(self, msg):
        """线程安全地插入日志"""
        self.root.after(0, lambda: self.log.insert(tk.END, msg + "\n"))
        self.root.after(0, lambda: self.log.see(tk.END))


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()