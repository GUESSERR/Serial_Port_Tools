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
    received = pyqtSignal(bytes)
    error = pyqtSignal(str)

    def __init__(self, port=None, baudrate=9600):
        super().__init__()
        self.serial = None
        self.port = port
        self.baudrate = baudrate
        self.running = False
        self.close_en = 0

    def open(self):
        try:
            self.serial = Serial(self.port, self.baudrate, timeout=0.1)
            self.running = True
        except SerialException as e:
            self.error.emit(f"打开串口失败: {str(e)}")

    def close(self):
        self.close_en = 1

    def detected_close_fun(self):
        if self.close_en == 1 and self.serial and self.serial.is_open:
            self.serial.close()
        self.running = False

    def read_data(self):
        while self.running:
            if self.serial and self.serial.in_waiting > 0:
                data = self.serial.read(self.serial.in_waiting)
                self.received.emit(data)
            self.detected_close_fun()
            QThread.msleep(10)

    def serial_send_data(self, data):
        if self.running and self.serial and self.serial.is_open:
            try:
                self.serial.write(data)
            except SerialException as e:
                self.error.emit(f"发送失败: {str(e)}")


def get_current_time():
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]


class SerialTool(QMainWindow):
    def __init__(self):
        super().__init__()
        self.baud_combo = None
        self.open_btn = None
        self.send_text = None
        self.recv_text = None
        self.send_btn = None
        self.hex_display = None
        self.cmd_table = None
        self.add_btn = None
        self.add_btn = None
        self.del_btn = None
        self.del_btn = None
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
        self.setWindowTitle("PySerial Tool")
        self.setGeometry(100, 100, 900, 600)

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)

        # 串口配置区域
        config_group = QGroupBox("串口配置")
        config_layout = QHBoxLayout(config_group)

        self.port_combo = QComboBox()
        self.baud_combo = QComboBox()
        self.baud_combo.addItems(["9600", "115200", "921600", "1000000"])
        self.open_btn = QPushButton("打开")
        self.open_btn.clicked.connect(self.toggle_serial)

        config_layout.addWidget(QLabel("端口:"))
        config_layout.addWidget(self.port_combo)
        config_layout.addWidget(QLabel("波特率:"))
        config_layout.addWidget(self.baud_combo)
        config_layout.addWidget(self.open_btn)

        # 数据收发区域
        data_group = QGroupBox("数据通信")
        data_layout = QVBoxLayout(data_group)

        self.recv_text = QTextEdit()
        self.send_text = QLineEdit()
        self.send_btn = QPushButton("发送")
        self.hex_display = QComboBox()
        self.hex_display.addItems(["ASCII", "HEX"])

        ctrl_layout = QHBoxLayout()
        ctrl_layout.addWidget(self.hex_display)
        ctrl_layout.addWidget(self.send_text)
        ctrl_layout.addWidget(self.send_btn)

        data_layout.addWidget(self.recv_text)
        data_layout.addLayout(ctrl_layout)

        # 指令存储区域
        cmd_group = QGroupBox("指令管理")
        cmd_layout = QVBoxLayout(cmd_group)

        self.cmd_table = QTableWidget(0, 4)
        self.cmd_table.setHorizontalHeaderLabels(["名称", "指令", "类型", "备注"])
        self.cmd_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.cmd_table.doubleClicked.connect(self.send_stored_command)

        btn_layout = QHBoxLayout()
        self.add_btn = QPushButton("添加")
        self.add_btn.clicked.connect(self.add_command)
        self.del_btn = QPushButton("删除")
        self.del_btn.clicked.connect(self.del_command)

        btn_layout.addWidget(self.add_btn)
        btn_layout.addWidget(self.del_btn)

        cmd_layout.addWidget(self.cmd_table)
        cmd_layout.addLayout(btn_layout)

        # 组合布局
        layout.addWidget(config_group)
        layout.addWidget(data_group)
        layout.addWidget(cmd_group)

        # 定时刷新端口
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_ports)
        self.timer.start(1000)

        # 信号连接
        self.send_btn.clicked.connect(self.send_data)
        self.send_text.returnPressed.connect(self.send_data)

    def init_menu(self):
        menu = self.menuBar()
        file_menu = menu.addMenu("文件")

        export_action = QAction("导出指令", self)
        export_action.triggered.connect(self.export_commands)
        import_action = QAction("导入指令", self)
        import_action.triggered.connect(self.import_commands)
        log_action = QAction("设置日志路径", self)
        log_action.triggered.connect(self.set_log_path)

        file_menu.addAction(export_action)
        file_menu.addAction(import_action)
        file_menu.addAction(log_action)

    # 核心功能实现
    def update_ports(self):
        current = self.port_combo.currentText()
        ports = [p.device for p in list_ports.comports()]
        self.port_combo.clear()
        self.port_combo.addItems(ports)
        if current in ports:
            self.port_combo.setCurrentText(current)

    def toggle_serial(self):
        if self.open_btn.text() == "打开":
            self.start_serial()
        else:
            self.stop_serial()

    def start_serial(self):
        port = self.port_combo.currentText()
        baud = int(self.baud_combo.currentText())

        self.serial_worker = SerialWorker(port, baud)
        self.worker_thread = QThread()
        self.serial_worker.moveToThread(self.worker_thread)

        self.serial_worker.received.connect(self.handle_received)
        self.serial_worker.error.connect(self.show_error)

        self.worker_thread.started.connect(self.serial_worker.open)
        self.worker_thread.started.connect(self.serial_worker.read_data)
        self.worker_thread.start()

        self.open_btn.setText("关闭")
        self.log_message(f"已连接 {port} @ {baud}")

    def stop_serial(self):
        if self.serial_worker:
            self.serial_worker.close()
            self.worker_thread.quit()
            self.worker_thread.wait()
        self.open_btn.setText("打开")
        self.log_message("串口已关闭")

    def handle_received(self, data):
        display_type = self.hex_display.currentText()
        text = data.hex(' ') if display_type == "HEX" else data.decode('ascii', 'replace')
        self.recv_text.append(f"[RX] {get_current_time()} {text}")
        self.save_log(f"[RX] {get_current_time()} {text}")

    def send_data(self):
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
        dialog = CommandDialog(self)
        if dialog.exec_():
            self.commands.append(dialog.get_command())
            self.update_command_table()

    def del_command(self):
        row = self.cmd_table.currentRow()
        if row >= 0:
            del self.commands[row]
            self.update_command_table()

    def send_stored_command(self):
        row = self.cmd_table.currentRow()
        cmd = self.commands[row]
        self.send_text.setText(cmd['command'])
        self.hex_display.setCurrentText(cmd['type'])
        self.send_data()

    def update_command_table(self):
        self.cmd_table.setRowCount(len(self.commands))
        for i, cmd in enumerate(self.commands):
            self.cmd_table.setItem(i, 0, QTableWidgetItem(cmd['name']))
            self.cmd_table.setItem(i, 1, QTableWidgetItem(cmd['command']))
            self.cmd_table.setItem(i, 2, QTableWidgetItem(cmd['type']))
            self.cmd_table.setItem(i, 3, QTableWidgetItem(cmd['note']))

    # 文件操作
    def export_commands(self):
        path, _ = QFileDialog.getSaveFileName(self, "导出指令", "", "JSON文件 (*.json)")
        if path:
            with Path(path).open('w') as f:
                json.dump(self.commands, f)

    def import_commands(self):
        path, _ = QFileDialog.getOpenFileName(self, "导入指令", "", "JSON文件 (*.json)")
        if path and Path(path).exists():
            with Path(path).open() as f:
                self.commands = json.load(f)
                self.update_command_table()

    # 日志管理
    def set_log_path(self):
        path = QFileDialog.getExistingDirectory(self, "选择日志目录")
        if path:
            self.log_path = Path(path)
            self.log_message(f"日志路径已更改为: {self.log_path}")

    def save_log(self, content):
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
        self.recv_text.append(f"{msg}")

    # 辅助方法
    def show_error(self, msg):
        self.recv_text.append(f"[ERROR] {get_current_time()}\n{msg}")
        self.save_log(f"[ERROR] {get_current_time()}\n{msg}")

    def load_commands(self):
        default_path = Path("commands.json")
        if default_path.exists():
            try:
                with default_path.open() as f:
                    self.commands = json.load(f)
                    self.update_command_table()
            except json.JSONDecodeError:
                pass

    def closeEvent(self, event):
        self.stop_serial()
        if self.commands:
            with Path("commands.json").open('w') as f:
                json.dump(self.commands, f)
        event.accept()


class CommandDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("添加指令")
        layout = QVBoxLayout(self)

        self.name = QLineEdit()
        self.command = QLineEdit()
        self.type_combo = QComboBox()
        self.type_combo.addItems(["ASCII", "HEX"])
        self.note = QLineEdit()

        form_layout = QHBoxLayout()
        form_layout.addWidget(QLabel("名称:"))
        form_layout.addWidget(self.name)
        form_layout.addWidget(QLabel("类型:"))
        form_layout.addWidget(self.type_combo)

        layout.addLayout(form_layout)
        layout.addWidget(QLabel("指令内容:"))
        layout.addWidget(self.command)
        layout.addWidget(QLabel("备注:"))
        layout.addWidget(self.note)

        self.confirm_btn = QPushButton("确认")
        self.confirm_btn.clicked.connect(self.accept)
        layout.addWidget(self.confirm_btn)

    def get_command(self):
        return {
            'name': self.name.text(),
            'command': self.command.text(),
            'type': self.type_combo.currentText(),
            'note': self.note.text()
        }

    def accept(self):
        if self.name.text() and self.command.text():
            super().close()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SerialTool()
    window.show()
    sys.exit(app.exec_())
