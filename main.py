import json
import sys
from datetime import datetime
from pathlib import Path
import time
from PyQt5.QtCore import QTimer, pyqtSignal, QObject, QThread
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QComboBox, QPushButton, QTextEdit, QLineEdit, QLabel,
                             QGroupBox, QTableWidget, QTableWidgetItem, QHeaderView,
                             QFileDialog, QAction)
from PyQt5.QtWidgets import QDialog  # 修改导入
from serial import Serial, SerialException
from serial.tools import list_ports


class SerialWorker(QObject):
    # 定义两个信号，一个用于接收数据，一个用于发送错误信息
    received = pyqtSignal(bytes)
    error = pyqtSignal(str)

    # 初始化函数，传入串口号和波特率
    def __init__(self, port=None, baudrate=9600):
        super().__init__()
        self.serial = None  # 串口对象
        self.port = port  # 串口号
        self.baudrate = baudrate  # 波特率
        self.running = False  # 是否正在运行
        self.close_en = 0  # 关闭标志

    # 打开串口
    def open(self):
        try:
            self.serial = Serial(self.port, self.baudrate, timeout=0.1)  # 创建串口对象
            self.running = True  # 设置正在运行标志
        except SerialException as e:
            self.error.emit(f"打开串口失败: {str(e)}")  # 发送错误信息

    # 关闭串口
    def close(self):
        self.close_en = 1  # 设置关闭标志

    # 检测关闭函数
    def detected_close_fun(self):
        if self.close_en == 1 and self.serial and self.serial.is_open:  # 如果关闭标志为1，并且串口对象存在且串口已打开
            self.serial.close()  # 关闭串口
        self.running = False  # 设置正在运行标志为False

    # 读取数据
    def read_data(self):
        while self.running:  # 如果正在运行
            if self.serial and self.serial.in_waiting > 0:  # 如果串口对象存在且串口有数据可读
                data = self.serial.read(self.serial.in_waiting)  # 读取数据
                self.received.emit(data)  # 发送数据信号
            self.detected_close_fun()  # 检测关闭函数
            QThread.msleep(10)  # 线程休眠10毫秒

    # 发送数据
    def serial_send_data(self, data):
        if self.running and self.serial and self.serial.is_open:  # 如果正在运行，并且串口对象存在且串口已打开
            try:
                self.serial.write(data)  # 发送数据
            except SerialException as e:
                self.error.emit(f"发送失败: {str(e)}")  # 发送错误信息


# 获取当前时间
def get_current_time():
    # 返回当前时间的字符串格式，格式为'年-月-日 时:分:秒.毫秒'
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]


class SerialTool(QMainWindow):
    def __init__(self):
        # 初始化父类
        super().__init__()
        # 初始化波特率下拉框
        self.baud_combo = None
        # 初始化打开按钮
        self.open_btn = None
        # 初始化发送文本框
        self.send_text = None
        # 初始化接收文本框
        self.recv_text = None
        # 初始化发送按钮
        self.send_btn = None
        # 初始化十六进制显示
        self.hex_display = None
        # 初始化命令表格
        self.cmd_table = None
        # 初始化添加按钮
        self.add_btn = None
        self.del_btn = None
        # 初始化定时器
        self.timer = None
        self.timer = None
        self.port_combo = None
        self.serial_worker = None
        self.worker_thread = None
        self.commands = []
        self.log_path = Path("logs")
        self.init_ui()
        self.init_menu()
        self.load_commands()
        self.update_ports()

    def init_ui(self):
        # 设置窗口标题
        self.setWindowTitle("PySerial Tool")
        # 设置窗口大小
        self.setGeometry(100, 100, 900, 600)

        # 创建主窗口部件
        main_widget = QWidget()
        # 设置主窗口部件为中央部件
        self.setCentralWidget(main_widget)
        # 创建垂直布局
        layout = QVBoxLayout(main_widget)

        # 串口配置区域
        config_group = QGroupBox("串口配置")
        config_layout = QHBoxLayout(config_group)

        # 创建串口下拉框
        self.port_combo = QComboBox()
        # 创建波特率下拉框
        self.baud_combo = QComboBox()
        # 添加波特率选项
        self.baud_combo.addItems(["9600", "115200", "921600", "1000000"])
        # 创建打开按钮
        self.open_btn = QPushButton("打开")
        # 连接打开按钮的点击事件
        self.open_btn.clicked.connect(self.toggle_serial)

        # 添加控件到布局
        config_layout.addWidget(QLabel("端口:"))
        config_layout.addWidget(self.port_combo)
        config_layout.addWidget(QLabel("波特率:"))
        config_layout.addWidget(self.baud_combo)
        config_layout.addWidget(self.open_btn)

        # 数据收发区域
        data_group = QGroupBox("数据通信")
        data_layout = QVBoxLayout(data_group)

        # 创建接收文本框
        self.recv_text = QTextEdit()
        # 创建发送文本框
        self.send_text = QLineEdit()
        # 创建发送按钮
        self.send_btn = QPushButton("发送")
        # 创建HEX显示下拉框
        self.hex_display = QComboBox()
        # 添加HEX显示选项
        self.hex_display.addItems(["ASCII", "HEX"])

        # 创建控制布局
        ctrl_layout = QHBoxLayout()
        # 添加HEX显示下拉框、发送文本框和发送按钮到控制布局
        ctrl_layout.addWidget(self.hex_display)
        ctrl_layout.addWidget(self.send_text)
        ctrl_layout.addWidget(self.send_btn)

        # 添加接收文本框和控制布局到数据布局
        data_layout.addWidget(self.recv_text)
        data_layout.addLayout(ctrl_layout)

        # 指令存储区域
        cmd_group = QGroupBox("指令管理")
        cmd_layout = QVBoxLayout(cmd_group)

        # 创建指令表格
        self.cmd_table = QTableWidget(0, 4)
        # 设置表格表头
        self.cmd_table.setHorizontalHeaderLabels(["名称", "指令", "类型", "备注"])
        # 设置表格列宽自动调整
        self.cmd_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        # 连接表格双击事件
        self.cmd_table.doubleClicked.connect(self.send_stored_command)

        # 创建按钮布局
        btn_layout = QHBoxLayout()
        # 创建添加按钮
        self.add_btn = QPushButton("添加")
        # 连接添加按钮的点击事件
        self.add_btn.clicked.connect(self.add_command)
        # 创建删除按钮
        self.del_btn = QPushButton("删除")
        # 连接删除按钮的点击事件
        self.del_btn.clicked.connect(self.del_command)

        # 添加添加按钮和删除按钮到按钮布局
        btn_layout.addWidget(self.add_btn)
        btn_layout.addWidget(self.del_btn)

        # 添加指令表格和按钮布局到指令布局
        cmd_layout.addWidget(self.cmd_table)
        cmd_layout.addLayout(btn_layout)

        # 组合布局
        layout.addWidget(config_group)
        layout.addWidget(data_group)
        layout.addWidget(cmd_group)

        # 定时刷新端口
        self.timer = QTimer()
        # 连接定时器超时事件
        self.timer.timeout.connect(self.update_ports)
        # 启动定时器
        self.timer.start(1000)

        # 信号连接
        # 连接发送按钮的点击事件
        self.send_btn.clicked.connect(self.send_data)
        # 连接发送文本框的回车事件
        self.send_text.returnPressed.connect(self.send_data)

    def init_menu(self):
        # 初始化菜单栏
        menu = self.menuBar()
        # 添加文件菜单
        file_menu = menu.addMenu("文件")

        # 添加导出指令菜单项
        export_action = QAction("导出指令", self)
        # 连接导出指令菜单项的触发事件到export_commands方法
        export_action.triggered.connect(self.export_commands)
        # 添加导入指令菜单项
        import_action = QAction("导入指令", self)
        # 连接导入指令菜单项的触发事件到import_commands方法
        import_action.triggered.connect(self.import_commands)
        # 添加设置日志路径菜单项
        log_action = QAction("设置日志路径", self)
        # 连接设置日志路径菜单项的触发事件到set_log_path方法
        log_action.triggered.connect(self.set_log_path)

        # 将导出指令菜单项添加到文件菜单
        file_menu.addAction(export_action)
        # 将导入指令菜单项添加到文件菜单
        file_menu.addAction(import_action)
        # 将设置日志路径菜单项添加到文件菜单
        file_menu.addAction(log_action)

    # 核心功能实现
    def update_ports(self):
        # 更新串口列表
        current = self.port_combo.currentText()  # 获取当前选中的串口
        ports = [p.device for p in list_ports.comports()]  # 获取当前可用的串口列表
        self.port_combo.clear()  # 清空串口列表
        self.port_combo.addItems(ports)  # 将可用的串口添加到串口列表中
        if current in ports:
            self.port_combo.setCurrentText(current)  # 如果当前选中的串口在可用的串口列表中，则将其设置为当前选中的串口

    def toggle_serial(self):
        # 切换串口连接状态
        if self.open_btn.text() == "打开":
            # 如果打开按钮的文本为"打开"，则调用start_serial()方法
            self.start_serial()
        else:
            # 否则，调用stop_serial()方法
            self.stop_serial()

    def start_serial(self):
        # 打开串口
        port = self.port_combo.currentText()
        baud = int(self.baud_combo.currentText())

        # 创建串口工作线程
        self.serial_worker = SerialWorker(port, baud)
        self.worker_thread = QThread()
        self.serial_worker.moveToThread(self.worker_thread)

        # 连接串口工作线程的信号和槽
        self.serial_worker.received.connect(self.handle_received)
        self.serial_worker.error.connect(self.show_error)

        # 连接串口工作线程的启动信号和槽
        self.worker_thread.started.connect(self.serial_worker.open)
        self.worker_thread.started.connect(self.serial_worker.read_data)
        self.worker_thread.start()

        # 更新按钮文本和日志信息
        self.open_btn.setText("关闭")
        self.log_message(f"已连接 {port} @ {baud}")

    def stop_serial(self):
        # 关闭串口
        if self.serial_worker:
            self.serial_worker.close()
            self.worker_thread.quit()
            self.worker_thread.wait()
        self.open_btn.setText("打开")
        self.log_message("串口已关闭")

    def handle_received(self, data):
        # 处理接收到的数据
        display_type = self.hex_display.currentText()
        text = data.hex(' ') if display_type == "HEX" else data.decode('ascii', 'replace')
        self.recv_text.append(f"[RX] {get_current_time()} {text}")
        self.save_log(f"[RX] {get_current_time()} {text}")

    def send_data(self):
        # 发送数据
        text = self.send_text.text()
        if not text:
            return

        display_type = self.hex_display.currentText()
        try:
            data = bytes.fromhex(text) if display_type == "HEX" else text.encode()
            if self.serial_worker:
                self.serial_worker.serial_send_data(data)
            self.recv_text.append(f"[TX] [{get_current_time()}] {text}")
            self.send_text.clear()
            self.save_log(f"[TX] [{get_current_time()}] {text}")
        except ValueError as e:
            self.show_error(f"数据格式错误: {str(e)}")

    # 指令管理功能
    def add_command(self):
        # 添加指令
        dialog = CommandDialog(self)
        if dialog.exec_():
            self.commands.append(dialog.get_command())
            self.update_command_table()

    def del_command(self):
        # 删除指令
        row = self.cmd_table.currentRow()
        if row >= 0:
            del self.commands[row]
            self.update_command_table()

    def send_stored_command(self):
        # 发送存储的指令
        row = self.cmd_table.currentRow()
        cmd = self.commands[row]
        self.send_text.setText(cmd['command'])
        self.hex_display.setCurrentText(cmd['type'])
        self.send_data()

    def update_command_table(self):
        # 更新指令表格
        self.cmd_table.setRowCount(len(self.commands))
        for i, cmd in enumerate(self.commands):
            self.cmd_table.setItem(i, 0, QTableWidgetItem(cmd['name']))
            self.cmd_table.setItem(i, 1, QTableWidgetItem(cmd['command']))
            self.cmd_table.setItem(i, 2, QTableWidgetItem(cmd['type']))
            self.cmd_table.setItem(i, 3, QTableWidgetItem(cmd['note']))

    # 文件操作
    def export_commands(self):
        # 导出指令
        path, _ = QFileDialog.getSaveFileName(self, "导出指令", "", "JSON文件 (*.json)")
        if path:
            with Path(path).open('w') as f:
                json.dump(self.commands, f)

    def import_commands(self):
        # 导入指令
        path, _ = QFileDialog.getOpenFileName(self, "导入指令", "", "JSON文件 (*.json)")
        # 弹出文件选择对话框，选择要导入的JSON文件
        if path and Path(path).exists():
            # 如果文件存在
            with Path(path).open() as f:
                # 打开文件
                self.commands = json.load(f)
                # 将文件内容加载到self.commands中
                self.update_command_table()

    # 日志管理
    def set_log_path(self):
        # 设置日志路径
        path = QFileDialog.getExistingDirectory(self, "选择日志目录")
        if path:
            self.log_path = Path(path)
            self.log_message(f"日志路径已更改为: {self.log_path}")

    def save_log(self, content):
        # 保存日志
        log_file = self.log_path / "serial.log"
        from pathlib import Path

        file_path = log_file
        # 创建父目录（若不存在）
        file_path.parent.mkdir(parents=True, exist_ok=True)
        # 创建文件（若不存在）
        file_path.touch(exist_ok=True)

        with log_file.open('a') as f:
            f.write(f"{content}\n")

    def log_message(self, msg):
        # 记录日志
        self.recv_text.append(f"{msg}")

    # 辅助方法
    def show_error(self, msg):
        # 显示错误信息
        self.recv_text.append(f"[ERROR] {get_current_time()}\n{msg}")
        self.save_log(f"[ERROR] {get_current_time()}\n{msg}")

    def load_commands(self):
        # 加载指令
        default_path = Path("commands.json")
        if default_path.exists():
            try:
                # 打开文件
                with default_path.open() as f:
                    # 加载json文件
                    self.commands = json.load(f)
                    # 更新指令表
                    self.update_command_table()
            except json.JSONDecodeError:
                # 如果json文件格式错误，则跳过
                pass

    def closeEvent(self, event):
        # 关闭窗口时保存指令
        self.stop_serial()
        if self.commands:
            # 如果有指令，则将指令保存到commands.json文件中
            with Path("commands.json").open('w') as f:
                json.dump(self.commands, f)
        event.accept()


class CommandDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("添加指令")
        layout = QVBoxLayout(self)

        # 创建一个文本输入框，用于输入指令名称
        self.name = QLineEdit()
        # 创建一个文本输入框，用于输入指令内容
        self.command = QLineEdit()
        # 创建一个下拉框，用于选择指令类型
        self.type_combo = QComboBox()
        # 向下拉框中添加选项
        self.type_combo.addItems(["ASCII", "HEX"])
        # 创建一个文本输入框，用于输入备注
        self.note = QLineEdit()

        # 创建一个水平布局，用于放置指令名称和指令类型的标签和输入框
        form_layout = QHBoxLayout()
        form_layout.addWidget(QLabel("名称:"))
        form_layout.addWidget(self.name)
        form_layout.addWidget(QLabel("类型:"))
        form_layout.addWidget(self.type_combo)

        # 将水平布局添加到垂直布局中
        layout.addLayout(form_layout)
        # 添加指令内容的标签和输入框
        layout.addWidget(QLabel("指令内容:"))
        layout.addWidget(self.command)
        # 添加备注的标签和输入框
        layout.addWidget(QLabel("备注:"))
        layout.addWidget(self.note)

        # 创建一个确认按钮，并连接到accept方法
        self.confirm_btn = QPushButton("确认")
        self.confirm_btn.clicked.connect(self.accept)
        # 将确认按钮添加到垂直布局中
        layout.addWidget(self.confirm_btn)

    # 获取指令信息
    def get_command(self):
        # 获取命令
        return {
            'name': self.name.text(),  # 获取命令名称
            'command': self.command.text(),  # 获取命令内容
            'type': self.type_combo.currentText(),  # 获取命令类型
            'note': self.note.text()  # 获取命令备注
        }

    # 确认按钮的点击事件
    def accept(self):
        # 如果指令名称和指令内容不为空，则关闭对话框
        if self.name.text() and self.command.text():
            super().close()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SerialTool()
    window.show()
    sys.exit(app.exec_())
