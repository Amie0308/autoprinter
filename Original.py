import sys
import os
import time
import threading
import shutil
import subprocess
import socket
from datetime import datetime
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QLabel, QPushButton, QTextEdit,
                             QSpinBox, QGroupBox, QMessageBox, QComboBox,
                             QFileDialog, QLineEdit, QListWidget, QListWidgetItem)
from PyQt5.QtCore import Qt, QTimer
import win32print
import win32api
import win32ui
import win32con

# 设置环境编码
os.environ['PYTHONIOENCODING'] = 'utf-8'


class PrintMonitorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.print_count = 0
        self.failed_prints = 0
        self.is_auto_printing = False
        self.current_print_job = None
        self.print_timeout = 120  # 2分钟超时
        self.print_start_time = None
        self.test_documents = []
        self.document_folder = os.path.join(os.path.expanduser("~"), "PrintTestDocuments")
        self.tcp_printers = {}  # 初始化TCP打印机配置字典
        self.print_job_retry_count = {}  # 用于跟踪打印作业的重试次数
        self.successful_prints = set()  # 跟踪成功打印的作业

        # 确保文档文件夹存在
        if not os.path.exists(self.document_folder):
            os.makedirs(self.document_folder)

        self.timer = QTimer()
        self.timer.timeout.connect(self.auto_print_test_page)

        self.timeout_timer = QTimer()
        self.timeout_timer.timeout.connect(self.check_print_timeout)

        self.print_thread = None

        self.init_ui()
        self.load_documents()
        self.refresh_printers()
        self.load_tcp_printers_config()  # 加载TCP打印机配置

    def init_ui(self):
        self.setWindowTitle('打印驱动测试工具')
        self.setGeometry(100, 100, 900, 700)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout(central_widget)

        # 状态信息组
        status_group = QGroupBox("打印状态")
        status_layout = QVBoxLayout()

        self.status_label = QLabel("就绪")
        status_layout.addWidget(self.status_label)

        self.print_count_label = QLabel(f"成功打印次数: {self.print_count}")
        status_layout.addWidget(self.print_count_label)

        self.failed_count_label = QLabel(f"失败打印次数: {self.failed_prints}")
        status_layout.addWidget(self.failed_count_label)

        self.current_job_label = QLabel("当前打印作业: 无")
        status_layout.addWidget(self.current_job_label)

        self.document_folder_label = QLabel(f"文档存储位置: {self.document_folder}")
        status_layout.addWidget(self.document_folder_label)

        status_group.setLayout(status_layout)
        layout.addWidget(status_group)

        # 控制组
        control_group = QGroupBox("打印控制")
        control_layout = QVBoxLayout()

        # 打印机选择
        printer_layout = QHBoxLayout()
        printer_layout.addWidget(QLabel("选择打印机:"))

        self.printer_combo = QComboBox()
        printer_layout.addWidget(self.printer_combo)

        refresh_btn = QPushButton("刷新打印机列表")
        refresh_btn.clicked.connect(self.refresh_printers)
        printer_layout.addWidget(refresh_btn)

        control_layout.addLayout(printer_layout)

        # 文档选择
        doc_layout = QHBoxLayout()
        doc_layout.addWidget(QLabel("选择测试文档:"))

        self.doc_combo = QComboBox()
        self.refresh_documents()
        doc_layout.addWidget(self.doc_combo)

        add_doc_btn = QPushButton("添加文档")
        add_doc_btn.clicked.connect(self.add_document)
        doc_layout.addWidget(add_doc_btn)

        refresh_doc_btn = QPushButton("刷新文档列表")
        refresh_doc_btn.clicked.connect(self.refresh_documents)
        doc_layout.addWidget(refresh_doc_btn)

        control_layout.addLayout(doc_layout)

        # 自动打印设置
        auto_print_layout = QHBoxLayout()
        auto_print_layout.addWidget(QLabel("自动打印间隔(分钟):"))

        self.interval_spin = QSpinBox()
        self.interval_spin.setRange(1, 120)
        self.interval_spin.setValue(30)
        auto_print_layout.addWidget(self.interval_spin)

        self.auto_print_btn = QPushButton("开始自动打印")
        self.auto_print_btn.clicked.connect(self.toggle_auto_print)
        auto_print_layout.addWidget(self.auto_print_btn)

        control_layout.addLayout(auto_print_layout)

        # 超时设置
        timeout_layout = QHBoxLayout()
        timeout_layout.addWidget(QLabel("打印超时(秒):"))

        self.timeout_spin = QSpinBox()
        self.timeout_spin.setRange(30, 300)
        self.timeout_spin.setValue(120)
        self.timeout_spin.valueChanged.connect(self.update_timeout)
        timeout_layout.addWidget(self.timeout_spin)

        control_layout.addLayout(timeout_layout)

        # 手动打印按钮
        self.manual_print_btn = QPushButton("立即打印测试文档")
        self.manual_print_btn.clicked.connect(self.manual_print_test_page)
        control_layout.addWidget(self.manual_print_btn)

        control_group.setLayout(control_layout)
        layout.addWidget(control_group)

        # 文档管理组
        doc_manage_group = QGroupBox("文档管理")
        doc_manage_layout = QVBoxLayout()

        self.doc_list = QListWidget()
        self.refresh_doc_list()
        doc_manage_layout.addWidget(self.doc_list)

        # 文档操作按钮
        doc_btn_layout = QHBoxLayout()
        open_folder_btn = QPushButton("打开文档文件夹")
        open_folder_btn.clicked.connect(self.open_document_folder)
        doc_btn_layout.addWidget(open_folder_btn)

        remove_doc_btn = QPushButton("删除选中文档")
        remove_doc_btn.clicked.connect(self.remove_document)
        doc_btn_layout.addWidget(remove_doc_btn)

        doc_manage_layout.addLayout(doc_btn_layout)
        doc_manage_group.setLayout(doc_manage_layout)
        layout.addWidget(doc_manage_group)

        # 日志区域
        log_group = QGroupBox("打印日志")
        log_layout = QVBoxLayout()

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)

        log_group.setLayout(log_layout)
        layout.addWidget(log_group)

        # 添加日志操作按钮
        log_btn_layout = QHBoxLayout()
        clear_btn = QPushButton("清空日志")
        clear_btn.clicked.connect(self.clear_log)
        log_btn_layout.addWidget(clear_btn)

        export_btn = QPushButton("导出日志")
        export_btn.clicked.connect(self.export_log)
        log_btn_layout.addWidget(export_btn)

        layout.addLayout(log_btn_layout)

        self.log_message("应用程序已启动")

    def refresh_printers(self):
        self.printer_combo.clear()
        try:
            # 获取所有类型的打印机
            printer_types = [
                win32print.PRINTER_ENUM_LOCAL,  # 本地打印机
                win32print.PRINTER_ENUM_CONNECTIONS,  # 网络打印机
                win32print.PRINTER_ENUM_NETWORK,  # 网络打印机（另一种方式）
                win32print.PRINTER_ENUM_SHARED,  # 共享打印机
                win32print.PRINTER_ENUM_NAME  # 按名称指定的打印机
            ]

            all_printers = []

            for printer_type in printer_types:
                try:
                    printers = win32print.EnumPrinters(printer_type, None, 1)
                    all_printers.extend(printers)
                except Exception as e:
                    self.log_message(f"获取打印机类型 {printer_type} 时出错: {str(e)}")

            # 去重处理
            unique_printers = {}
            for printer in all_printers:
                printer_name = printer[2]  # 打印机名称在元组的第三个位置
                if printer_name not in unique_printers:
                    unique_printers[printer_name] = printer

            # 添加到下拉框
            for printer_name, printer_info in unique_printers.items():
                self.printer_combo.addItem(printer_name, printer_info)

            # 选择默认打印机
            try:
                default_printer = win32print.GetDefaultPrinter()
                index = self.printer_combo.findText(default_printer)
                if index >= 0:
                    self.printer_combo.setCurrentIndex(index)
                    self.log_message(f"已选择默认打印机: {default_printer}")
                else:
                    self.log_message("未找到默认打印机")
            except Exception as e:
                self.log_message(f"获取默认打印机失败: {str(e)}")

            self.log_message(f"共找到 {len(unique_printers)} 台打印机")

        except Exception as e:
            self.log_message(f"获取打印机列表失败: {str(e)}")
            # 尝试使用备用方法
            self._get_printers_alternative()

    def _get_printers_alternative(self):
        """备用方法获取打印机列表"""
        try:
            # 使用 Windows 命令获取打印机列表
            result = subprocess.run(
                ['wmic', 'printer', 'get', 'name'],
                capture_output=True,
                text=True,
                encoding='gbk'  # 中文系统使用gbk编码
            )

            if result.returncode == 0:
                lines = result.stdout.strip().split('\n')
                for line in lines[1:]:  # 跳过标题行
                    printer_name = line.strip()
                    if printer_name:
                        self.printer_combo.addItem(printer_name)

                self.log_message("使用备用方法获取打印机列表成功")
            else:
                self.log_message("备用方法获取打印机列表失败")

        except Exception as e:
            self.log_message(f"备用方法也失败了: {str(e)}")

    def refresh_documents(self):
        self.doc_combo.clear()
        self.load_documents()
        for doc in self.test_documents:
            self.doc_combo.addItem(os.path.basename(doc), doc)

    def refresh_doc_list(self):
        self.doc_list.clear()
        self.load_documents()
        for doc in self.test_documents:
            item = QListWidgetItem(os.path.basename(doc))
            item.setData(Qt.UserRole, doc)
            self.doc_list.addItem(item)

    def load_documents(self):
        self.test_documents = []
        try:
            for file in os.listdir(self.document_folder):
                if file.lower().endswith(('.pdf', '.txt', '.doc', '.docx', '.xps', '.tcp', '.usb')):
                    self.test_documents.append(os.path.join(self.document_folder, file))
        except Exception as e:
            self.log_message(f"加载文档失败: {str(e)}")

    def add_document(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择测试文档", "",
            "文档文件 (*.pdf *.txt *.doc *.docx *.xps *.tcp *.usb);;所有文件 (*.*)"
        )

        if file_path:
            try:
                # 复制文件到文档文件夹
                file_name = os.path.basename(file_path)
                dest_path = os.path.join(self.document_folder, file_name)
                shutil.copy2(file_path, dest_path)

                self.log_message(f"已添加文档: {file_name}")
                self.refresh_documents()
                self.refresh_doc_list()
            except Exception as e:
                self.log_message(f"添加文档失败: {str(e)}")

    def remove_document(self):
        current_item = self.doc_list.currentItem()
        if current_item:
            file_path = current_item.data(Qt.UserRole)
            try:
                os.remove(file_path)
                self.log_message(f"已删除文档: {os.path.basename(file_path)}")
                self.refresh_documents()
                self.refresh_doc_list()
            except Exception as e:
                self.log_message(f"删除文档失败: {str(e)}")
        else:
            self.log_message("请先选择一个文档")

    def open_document_folder(self):
        try:
            os.startfile(self.document_folder)
        except Exception as e:
            self.log_message(f"打开文件夹失败: {str(e)}")

    def update_timeout(self):
        self.print_timeout = self.timeout_spin.value()

    def toggle_auto_print(self):
        if self.is_auto_printing:
            self.stop_auto_print()
        else:
            self.start_auto_print()

    def start_auto_print(self):
        interval_minutes = self.interval_spin.value()
        interval_ms = interval_minutes * 60 * 1000

        self.timer.start(interval_ms)
        self.is_auto_printing = True
        self.auto_print_btn.setText("停止自动打印")
        self.log_message(f"已启用自动打印，每 {interval_minutes} 分钟打印一次测试文档")

    def stop_auto_print(self):
        self.timer.stop()
        self.is_auto_printing = False
        self.auto_print_btn.setText("开始自动打印")
        self.log_message("已停止自动打印")

    def manual_print_test_page(self):
        if not self.printer_combo.count():
            self.log_message("错误: 没有可用的打印机")
            return

        if not self.test_documents:
            self.log_message("错误: 没有可用的测试文档")
            return

        self.log_message("手动请求打印测试文档")
        self.print_test_page()

    def auto_print_test_page(self):
        if not self.printer_combo.count() or not self.test_documents:
            self.log_message("跳过自动打印：打印机或文档不可用")
            return

        self.log_message("自动打印测试文档")
        self.print_test_page()

    def print_test_page(self):
        # 使用线程打印，避免阻塞UI
        self.print_thread = threading.Thread(target=self._print_test_page_thread)
        self.print_thread.daemon = True
        self.print_thread.start()

    def _print_test_page_thread(self):
        try:
            printer_name = self.printer_combo.currentText()
            doc_index = self.doc_combo.currentIndex()

            if doc_index < 0 or doc_index >= len(self.test_documents):
                self.log_message("错误: 没有可用的测试文档")
                return

            doc_path = self.test_documents[doc_index]
            self.print_start_time = datetime.now()
            self.timeout_timer.start(1000)

            self.log_message(f"开始打印: {os.path.basename(doc_path)} 到 {printer_name}")

            # 根据扩展名选择打印方式
            ext = os.path.splitext(doc_path)[1].lower()

            if ext == ".txt":
                self._print_txt(doc_path, printer_name)
                # 对于文本文件，我们假设打印成功，因为ShellExecute可能不会创建打印作业
                self._handle_print_success(f"{printer_name}_{os.path.basename(doc_path)}")
                return

            elif ext in [".doc", ".docx"]:
                self._print_word(doc_path, printer_name)
                # 对于Word文件，我们假设打印成功
                self._handle_print_success(f"{printer_name}_{os.path.basename(doc_path)}")
                return

            elif ext == ".pdf":
                self._print_pdf(doc_path, printer_name)
                # 对于PDF文件，我们假设打印成功
                self._handle_print_success(f"{printer_name}_{os.path.basename(doc_path)}")
                return

            elif ext == ".usb":
                self._print_raw_usb(doc_path, printer_name)

            elif ext == ".tcp":
                self._print_raw_tcp(doc_path, printer_name)

            else:
                # 默认使用RAW打印
                self._print_raw(doc_path, printer_name)

            self.current_print_job = {
                'printer': printer_name,
                'document': os.path.basename(doc_path),
                'start_time': self.print_start_time,
                'status': 'sent'
            }
            self.update_job_status()

            # 初始化重试计数器
            job_key = f"{printer_name}_{os.path.basename(doc_path)}"
            self.print_job_retry_count[job_key] = 0

            # 立即检查打印状态
            self._check_print_status(printer_name, os.path.basename(doc_path))

        except Exception as e:
            self.failed_prints += 1
            self.log_message(f"打印失败: {e}")
            self.timeout_timer.stop()
            self.update_status_labels()

    def _print_txt(self, doc_path, printer_name):
        """打印文本文件到指定打印机"""
        try:
            # 使用系统命令打印文本文件到指定打印机
            win32api.ShellExecute(
                0,
                "printto",
                doc_path,
                f'"{printer_name}"',
                ".",
                0
            )
            self.log_message(f"文本文件打印任务已发送到打印机: {printer_name}")
        except Exception as e:
            raise Exception(f"文本文件打印失败: {str(e)}")

    def _print_word(self, doc_path, printer_name):
        """打印Word文档到指定打印机"""
        try:
            # 使用系统命令打印Word文档到指定打印机
            win32api.ShellExecute(
                0,
                "printto",
                doc_path,
                f'"{printer_name}"',
                ".",
                0
            )
            self.log_message(f"Word文档打印任务已发送到打印机: {printer_name}")
        except Exception as e:
            raise Exception(f"Word文档打印失败: {str(e)}")

    def _print_pdf(self, doc_path, printer_name):
        """打印PDF文档到指定打印机"""
        try:
            # 使用系统命令打印PDF文档到指定打印机
            win32api.ShellExecute(
                0,
                "printto",
                doc_path,
                f'"{printer_name}"',
                ".",
                0
            )
            self.log_message(f"PDF文档打印任务已发送到打印机: {printer_name}")
        except Exception as e:
            raise Exception(f"PDF文档打印失败: {str(e)}")

    def _print_raw_usb(self, doc_path, printer_name):
        """处理USB打印机原始打印"""
        try:
            # 查找USB打印机对应的物理端口
            hPrinter = win32print.OpenPrinter(printer_name)
            try:
                job_info = ("Python打印任务", None, "RAW")
                job_id = win32print.StartDocPrinter(hPrinter, 1, job_info)
                win32print.StartPagePrinter(hPrinter)

                with open(doc_path, 'rb') as f:
                    data = f.read(1024)
                    while data:
                        win32print.WritePrinter(hPrinter, data)
                        data = f.read(1024)

                win32print.EndPagePrinter(hPrinter)
                win32print.EndDocPrinter(hPrinter)
                self.log_message(f"USB原始数据已发送到打印机: {printer_name}")

            finally:
                win32print.ClosePrinter(hPrinter)

        except Exception as e:
            raise Exception(f"USB打印失败: {str(e)}")

    def _print_raw_tcp(self, doc_path, printer_name):
        """处理TCP/IP打印机原始打印"""
        try:
            # 从配置中获取TCP参数
            if printer_name not in self.tcp_printers:
                raise Exception(f"未配置TCP打印机: {printer_name}")

            ip, port = self.tcp_printers[printer_name]
            self.log_message(f"连接到TCP打印机: {ip}:{port}")

            # 读取文件数据
            with open(doc_path, 'rb') as f:
                data = f.read()

            # 建立TCP连接并发送数据
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.settimeout(10)
                s.connect((ip, port))
                s.sendall(data)
                self.log_message(f"TCP数据已发送 ({len(data)} 字节)")

        except Exception as e:
            raise Exception(f"TCP打印失败: {str(e)}")

    def _print_raw(self, doc_path, printer_name):
        """通用RAW打印方法"""
        try:
            hPrinter = win32print.OpenPrinter(printer_name)
            try:
                job_info = ("Python打印任务", None, "RAW")
                job_id = win32print.StartDocPrinter(hPrinter, 1, job_info)
                win32print.StartPagePrinter(hPrinter)

                with open(doc_path, 'rb') as f:
                    data = f.read(1024)
                    while data:
                        win32print.WritePrinter(hPrinter, data)
                        data = f.read(1024)

                win32print.EndPagePrinter(hPrinter)
                win32print.EndDocPrinter(hPrinter)
                self.log_message(f"RAW打印任务发送到打印机: {printer_name}")

            finally:
                win32print.ClosePrinter(hPrinter)

        except Exception as e:
            raise Exception(f"RAW打印失败: {str(e)}")

    def load_tcp_printers_config(self):
        """加载TCP打印机配置（示例配置）"""
        # 示例配置：打印机名称 -> (IP, 端口)
        self.tcp_printers = {
            "MyTCPPrinter": ("192.168.1.100", 9100)
        }
        self.log_message(f"加载TCP打印机配置: {list(self.tcp_printers.keys())}")

    def _check_print_status(self, printer_name, document_name):
        try:
            # 检查是否已经成功打印
            job_key = f"{printer_name}_{document_name}"
            if job_key in self.successful_prints:
                return

            # 打开打印机
            handle = win32print.OpenPrinter(printer_name)

            # 获取打印作业
            jobs = win32print.EnumJobs(handle, 0, -1, 1)

            # 查找我们的打印作业（更宽松的匹配）
            our_job = None
            for job in jobs:
                job_doc_name = job['pDocument']
                # 使用更宽松的匹配，因为有时文档名可能会被修改
                if document_name in job_doc_name or job_doc_name in document_name:
                    our_job = job
                    break

            if our_job:
                job_status = our_job['Status']
                job_id = our_job['JobId']
                status_text = self._get_job_status_text(job_status)
                self.log_message(f"打印作业 ID {job_id} 状态: {status_text}")

                # 检查各种完成状态
                if (job_status == win32print.JOB_STATUS_COMPLETE or
                        job_status == win32print.JOB_STATUS_PRINTED or
                        job_status == win32print.JOB_STATUS_DELETED or
                        job_status == win32print.JOB_STATUS_DELETING or
                        job_status == win32print.JOB_STATUS_RESTART):
                    self._handle_print_success(job_key)
                elif job_status == win32print.JOB_STATUS_ERROR:
                    self.failed_prints += 1
                    self.update_status_labels()
                    self.log_message("打印作业出错")
                    self.current_print_job = None
                    self.update_job_status()
                else:
                    # 设置定时重新检查
                    QTimer.singleShot(3000, lambda: self._check_print_status(printer_name, document_name))
            else:
                # 可能作业已经完成并被移除，或者尚未进入队列
                job_key = f"{printer_name}_{document_name}"

                # 增加重试计数器
                if job_key not in self.print_job_retry_count:
                    self.print_job_retry_count[job_key] = 0
                self.print_job_retry_count[job_key] += 1

                # 如果重试次数超过3次，认为打印成功
                if self.print_job_retry_count[job_key] > 3:
                    self.log_message("未找到打印作业，但已达到最大重试次数，假设打印成功")
                    self._handle_print_success(job_key)
                else:
                    self.log_message(
                        f"未找到打印作业，可能已完成或尚未进入队列 (重试 {self.print_job_retry_count[job_key]}/3)")
                    # 等待一段时间后再次检查
                    QTimer.singleShot(5000, lambda: self._check_print_status(printer_name, document_name))

            win32print.ClosePrinter(handle)

        except Exception as e:
            self.log_message(f"检查打印状态时出错: {str(e)}")
            # 等待一段时间后重试
            QTimer.singleShot(5000, lambda: self._check_print_status(printer_name, document_name))

    def _get_job_status_text(self, status):
        status_map = {
            win32print.JOB_STATUS_PAUSED: "已暂停",
            win32print.JOB_STATUS_ERROR: "错误",
            win32print.JOB_STATUS_DELETING: "正在删除",
            win32print.JOB_STATUS_SPOOLING: "正在假脱机",
            win32print.JOB_STATUS_PRINTING: "正在打印",
            win32print.JOB_STATUS_OFFLINE: "脱机",
            win32print.JOB_STATUS_PAPEROUT: "缺纸",
            win32print.JOB_STATUS_PRINTED: "已打印",
            win32print.JOB_STATUS_DELETED: "已删除",
            win32print.JOB_STATUS_BLOCKED_DEVQ: "设备队列阻塞",
            win32print.JOB_STATUS_USER_INTERVENTION: "需要用户干预",
            win32print.JOB_STATUS_RESTART: "重新启动",
            win32print.JOB_STATUS_COMPLETE: "完成"
        }

        return status_map.get(status, f"未知状态 ({status})")

    def _handle_print_success(self, job_key):
        # 标记这个作业已经成功打印
        self.successful_prints.add(job_key)

        self.print_count += 1
        self.update_status_labels()
        self.log_message(f"打印成功完成! 总成功次数: {self.print_count}")
        self.current_print_job = None
        self.update_job_status()
        self.timeout_timer.stop()

    def check_print_timeout(self):
        if self.print_start_time and self.current_print_job:
            elapsed = (datetime.now() - self.print_start_time).total_seconds()
            if elapsed > self.print_timeout:
                self.failed_prints += 1
                self.update_status_labels()
                self.log_message(f"打印超时! 已等待 {elapsed:.0f} 秒")
                self.current_print_job = None
                self.update_job_status()
                self.timeout_timer.stop()

    def update_job_status(self):
        if self.current_print_job:
            elapsed = (datetime.now() - self.current_print_job['start_time']).total_seconds()
            self.current_job_label.setText(
                f"当前打印作业: {self.current_print_job['document']} "
                f"到 {self.current_print_job['printer']} ({elapsed:.0f}秒)"
            )
        else:
            self.current_job_label.setText("当前打印作业: 无")

    def update_status_labels(self):
        self.print_count_label.setText(f"成功打印次数: {self.print_count}")
        self.failed_count_label.setText(f"失败打印次数: {self.failed_prints}")

    def log_message(self, message):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"[{timestamp}] {message}"
        self.log_text.append(log_entry)
        # 自动滚动到底部
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )

    def clear_log(self):
        self.log_text.clear()
        self.log_message("日志已清空")

    def export_log(self):
        """导出日志到文件"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "导出日志", "", "文本文件 (*.txt);;所有文件 (*.*)"
        )

        if file_path:
            try:
                # 确保文件扩展名正确
                if not file_path.lower().endswith('.txt'):
                    file_path += '.txt'

                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(self.log_text.toPlainText())
                self.log_message(f"日志已导出到: {file_path}")
            except Exception as e:
                self.log_message(f"导出日志失败: {str(e)}")

    def closeEvent(self, event):
        # 停止定时器
        if self.timer.isActive():
            self.timer.stop()

        if self.timeout_timer.isActive():
            self.timeout_timer.stop()

        event.accept()


def main():
    app = QApplication(sys.argv)
    window = PrintMonitorApp()
    window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()