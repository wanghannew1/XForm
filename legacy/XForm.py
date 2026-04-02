import sys
import os
import xlwings as xw
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QPushButton, QListWidget, QTextEdit,
                             QProgressBar, QFileDialog, QMessageBox, QLabel)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont

# 转换单位：厘米 -> 磅 (1 cm = 28.35 pt)
CM_TO_POINTS = 28.35

class ProcessThread(QThread):
    """后台处理线程，避免界面卡顿"""
    log_signal = pyqtSignal(str)          # 日志信息
    progress_signal = pyqtSignal(int)     # 进度值 (0-100)
    finished_signal = pyqtSignal(bool, str)  # 是否成功，错误信息

    def __init__(self, file_list):
        super().__init__()
        self.file_list = file_list

    def run(self):
        total = len(self.file_list)
        for idx, file_path in enumerate(self.file_list):
            self.log_signal.emit(f"正在处理: {os.path.basename(file_path)}")
            try:
                self.process_one_file(file_path)
                self.log_signal.emit(f"处理成功: {os.path.basename(file_path)}")
            except Exception as e:
                self.log_signal.emit(f"处理失败: {os.path.basename(file_path)} - {str(e)}")
            progress = int((idx + 1) / total * 100)
            self.progress_signal.emit(progress)
        self.finished_signal.emit(True, "全部处理完成！")

    def process_one_file(self, file_path):
        """处理单个Excel文件"""
        # 启动Excel应用（后台不可见）
        app = xw.App(visible=False, add_book=False)
        try:
            wb = app.books.open(file_path)
            # 假设处理第一个工作表
            ws = wb.sheets[0]

            # 1. 设置打印页边距（左侧1cm，其他使用默认值）
            ws.api.PageSetup.LeftMargin = 1 * CM_TO_POINTS   # 左侧1cm
            # 其他边距保持原样，也可以显式设置（单位磅）
            # ws.api.PageSetup.RightMargin = ws.api.PageSetup.RightMargin
            # ws.api.PageSetup.TopMargin = ws.api.PageSetup.TopMargin
            # ws.api.PageSetup.BottomMargin = ws.api.PageSetup.BottomMargin

            # 2. 设置打印标题行（第1~5行每页重复）
            ws.api.PageSetup.PrintTitleRows = "$1:$5"

            # 3. 设置缩放比例：全部列在一页
            ws.api.PageSetup.Zoom = False
            ws.api.PageSetup.FitToPagesWide = 1
            ws.api.PageSetup.FitToPagesTall = False

            # 4. 隐藏部门列（D列）和岗位列（E列）
            ws.range('D:D').column_width = 0  # 设置列宽为0实现隐藏
            ws.range('E:E').column_width = 0

            # 5. 修改“部长、分管副总签字：”为“部长签字：”
            #    查找包含该文本的单元格（可能在最后几行）
            used_range = ws.used_range
            last_row = used_range.last_cell.row
            # 从最后一行向上搜索10行（可根据实际调整），找到即替换
            found = False
            for r in range(max(1, last_row - 20), last_row + 1):
                cell = ws.range((r, 1))
                if cell.value and "部长、分管副总签字：" in str(cell.value):
                    new_val = str(cell.value).replace("部长、分管副总签字：", "部长签字：")
                    cell.value = new_val
                    found = True
                    break
            if not found:
                self.log_signal.emit("警告：未找到‘部长、分管副总签字：’文本")

            # 6. 检查社保基数等四列是否全为0，若全为0则删除“业务审核：”
            #   先找到第3行（表头行）中这些列的位置
            header_row = 3
            # 需要查找的列名关键词
            keywords = ["社保基数", "医保基数", "工伤基数", "公积金基数"]
            col_indices = []  # 存储列号（1-based）
            # 遍历第3行的所有列（假设最多50列，可适当扩大）
            for col in range(1, 51):
                cell_val = ws.range((header_row, col)).value
                if cell_val:
                    for kw in keywords:
                        if kw in str(cell_val):
                            col_indices.append(col)
                            break
            if len(col_indices) != 4:
                self.log_signal.emit(f"警告：未找到全部四列基数，找到 {len(col_indices)} 列")
            else:
                # 检查从第4行到最后一行的这些列是否全为0或空
                all_zero = True
                for row in range(header_row + 1, last_row + 1):
                    for col in col_indices:
                        val = ws.range((row, col)).value
                        # 空值视为0，非数值视为非零
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
                    # 删除最后几行中的“业务审核：”字样
                    # 查找包含“业务审核：”的单元格（从最后向上找）
                    removed = False
                    for r in range(last_row, max(1, last_row - 20), -1):
                        cell = ws.range((r, 1))
                        if cell.value and "业务审核：" in str(cell.value):
                            new_val = str(cell.value).replace("业务审核：", "")
                            cell.value = new_val
                            removed = True
                            self.log_signal.emit("已删除‘业务审核：’字样（条件满足）")
                            break
                    if not removed:
                        self.log_signal.emit("警告：未找到‘业务审核：’文本")

            # 保存并关闭
            wb.save()
            wb.close()
        finally:
            app.quit()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("工资表打印格式批量调整工具")
        self.setGeometry(100, 100, 800, 600)
        self.file_list = []   # 存储待处理的文件路径

        # 中央控件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # 文件列表区域
        file_label = QLabel("待处理文件列表：")
        file_label.setFont(QFont("Microsoft YaHei", 10))
        main_layout.addWidget(file_label)

        self.file_list_widget = QListWidget()
        main_layout.addWidget(self.file_list_widget)

        # 按钮区域
        btn_layout = QHBoxLayout()
        self.btn_add_files = QPushButton("添加文件")
        self.btn_clear = QPushButton("清空列表")
        self.btn_process = QPushButton("开始处理")
        self.btn_process.setEnabled(False)
        btn_layout.addWidget(self.btn_add_files)
        btn_layout.addWidget(self.btn_clear)
        btn_layout.addStretch()
        btn_layout.addWidget(self.btn_process)
        main_layout.addLayout(btn_layout)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        main_layout.addWidget(self.progress_bar)

        # 日志区域
        log_label = QLabel("处理日志：")
        log_label.setFont(QFont("Microsoft YaHei", 10))
        main_layout.addWidget(log_label)
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        main_layout.addWidget(self.log_text)

        # 连接信号
        self.btn_add_files.clicked.connect(self.add_files)
        self.btn_clear.clicked.connect(self.clear_list)
        self.btn_process.clicked.connect(self.start_process)

    def add_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择Excel文件", "", "Excel文件 (*.xls *.xlsx)")
        if files:
            for f in files:
                if f not in self.file_list:
                    self.file_list.append(f)
                    self.file_list_widget.addItem(f)
            self.btn_process.setEnabled(len(self.file_list) > 0)

    def clear_list(self):
        self.file_list.clear()
        self.file_list_widget.clear()
        self.btn_process.setEnabled(False)

    def start_process(self):
        if not self.file_list:
            QMessageBox.warning(self, "提示", "请先添加文件！")
            return
        # 禁用按钮，防止重复点击
        self.btn_process.setEnabled(False)
        self.btn_add_files.setEnabled(False)
        self.btn_clear.setEnabled(False)
        self.progress_bar.setValue(0)
        self.log_text.clear()

        # 启动处理线程
        self.thread = ProcessThread(self.file_list.copy())
        self.thread.log_signal.connect(self.append_log)
        self.thread.progress_signal.connect(self.progress_bar.setValue)
        self.thread.finished_signal.connect(self.on_finished)
        self.thread.start()

    def append_log(self, msg):
        self.log_text.append(msg)

    def on_finished(self, success, msg):
        self.btn_process.setEnabled(True)
        self.btn_add_files.setEnabled(True)
        self.btn_clear.setEnabled(True)
        QMessageBox.information(self, "完成", msg)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())