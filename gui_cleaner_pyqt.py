import sys
import os
import pandas as pd
from pypinyin import pinyin, Style
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QLineEdit, QFileDialog, QLabel, QMessageBox
)
from PyQt6.QtCore import QThread, pyqtSignal, QObject


# --- 1. pypinyin 数据文件路径处理 (打包EXE时必需) ---
# 必须放在所有 import 之后，QApplication 启动之前

def get_pypinyin_data_path():
    """获取 pypinyin 数据文件的路径，兼容开发环境和PyInstaller"""
    try:
        # PyInstaller 打包后会创建 _MEIPASS 变量
        base_path = sys._MEIPASS
    except Exception:
        # 在正常 .py 脚本环境中
        import pypinyin as pypinyin_module
        base_path = os.path.dirname(pypinyin_module.__file__)

    data_path = os.path.join(base_path, 'data')

    # 额外检查：确保PyInstaller正确找到了pypinyin/data目录
    # 如果 data_path 不存在，尝试回退到 'pypinyin/data'
    # (这是 --add-data "pypinyin/data:pypinyin/data" 的目标路径)
    if not os.path.exists(data_path) and hasattr(sys, '_MEIPASS'):
        data_path = os.path.join(sys._MEIPASS, 'pypinyin', 'data')

    if not os.path.exists(data_path):
        print(f"警告：无法在 {data_path} 找到 pypinyin 数据。")

    return data_path


# 设置 pypinyin 的数据目录环境变量
os.environ['PYPINYIN_NO_DICT_COPY'] = 'true'
os.environ['PYPINYIN_DATA_DIR'] = get_pypinyin_data_path()


# --- 2. 核心清洗逻辑 (与原脚本相同) ---

def get_correct_pinyin(chinese_name):
    """
    将中文名转换为大写无音调的全拼
    """
    if not isinstance(chinese_name, str) or not chinese_name.strip():
        return ""
    pinyin_list = pinyin(chinese_name, style=Style.NORMAL)
    return ''.join([item[0] for item in pinyin_list]).upper()


# --- 3. PyQt6 多线程工作类 ---
# (将耗时的数据处理任务放入子线程，防止GUI卡死)

class Worker(QObject):
    # 定义信号：完成、错误、进度
    finished = pyqtSignal(str)
    error = pyqtSignal(str)
    progress = pyqtSignal(str)

    def __init__(self, file_path, sheet_name):
        super().__init__()
        self.file_path = file_path
        self.sheet_name = sheet_name

    def run(self):
        """主工作逻辑"""
        try:
            # 1. 校验字段 (您的核心需求)
            self.progress.emit("正在读取文件并校验字段...")
            try:
                # 仅读取表头 (nrows=0) 来快速校验
                df_header = pd.read_excel(self.file_path, sheet_name=self.sheet_name, nrows=0)
            except Exception as e:
                raise Exception(f"读取Sheet失败，请检查Sheet名称是否正确。\n错误: {e}")

            required_cols = {'clientname', 'patientcode'}
            if not required_cols.issubset(df_header.columns):
                missing = required_cols - set(df_header.columns)
                raise Exception(f"错误：字段校验失败！\n未在Sheet中找到以下列: {', '.join(missing)}")

            # 2. 字段校验通过，开始完整处理
            self.progress.emit(f"字段校验通过，正在读取所有数据...")
            df = pd.read_excel(self.file_path, sheet_name=self.sheet_name)
            self.progress.emit(f"读取成功 {len(df)} 条数据，开始清洗...")

            processed_count = 0
            for index, row in df.iterrows():
                client_name = row['clientname']
                patient_code = row['patientcode']

                if pd.notna(client_name) and pd.notna(patient_code):
                    correct_pinyin = get_correct_pinyin(client_name)
                    try:
                        parts = str(patient_code).split('_')
                        if len(parts) >= 2:
                            parts[0] = correct_pinyin
                            new_patient_code = '_'.join(parts)
                            df.at[index, 'patientcode'] = new_patient_code
                            processed_count += 1
                    except Exception as e:
                        print(f"警告：处理第 {index + 2} 行时发生错误。")

            self.progress.emit(f"清洗完成 {processed_count} 条，正在保存...")

            # 3. 保存回原文件
            with pd.ExcelWriter(self.file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=self.sheet_name, index=False)

            self.finished.emit(f"成功！共 {processed_count} 条数据被清洗并保存回原文件。")

        except Exception as e:
            self.error.emit(str(e))


# --- 4. PyQt6 主窗口界面 ---

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel拼音清洗工具 (by Zuo)")
        self.setGeometry(300, 300, 500, 200)

        # 布局
        layout = QVBoxLayout(self)

        # 1. 文件选择行
        file_layout = QHBoxLayout()
        self.file_path_edit = QLineEdit()
        self.file_path_edit.setPlaceholderText("请选择要清洗的Excel文件")
        self.file_path_edit.setReadOnly(True)
        self.browse_button = QPushButton("浏览...")
        self.browse_button.clicked.connect(self.browse_file)
        file_layout.addWidget(self.file_path_edit)
        file_layout.addWidget(self.browse_button)
        layout.addLayout(file_layout)

        # 2. Sheet名称输入行
        sheet_layout = QHBoxLayout()
        sheet_label = QLabel("Sheet 名称:")
        self.sheet_name_edit = QLineEdit()
        # 使用原脚本中的默认值作为提示
        self.sheet_name_edit.setText("dw_eagle_sale2_atm")
        sheet_layout.addWidget(sheet_label)
        sheet_layout.addWidget(self.sheet_name_edit)
        layout.addLayout(sheet_layout)

        # 3. 运行按钮
        self.run_button = QPushButton("开始清洗")
        self.run_button.clicked.connect(self.start_cleaning)
        layout.addWidget(self.run_button)

        # 4. 状态栏
        self.status_label = QLabel("准备就绪...")
        self.status_label.setStyleSheet("color: blue;")
        layout.addWidget(self.status_label)

        # QThread 实例
        self.worker_thread = None
        self.worker = None

    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择Excel文件",
            "",  # 默认目录
            "Excel 文件 (*.xlsx *.xls)"
        )
        if file_path:
            self.file_path_edit.setText(file_path)

    def start_cleaning(self):
        file_path = self.file_path_edit.text()
        sheet_name = self.sheet_name_edit.text().strip()  # 去除前后空格

        if not file_path:
            QMessageBox.warning(self, "警告", "请先选择一个Excel文件。")
            return
        if not sheet_name:
            QMessageBox.warning(self, "警告", "请输入要清洗的Sheet名称。")
            return

        # 禁用按钮，更新状态
        self.run_button.setEnabled(False)
        self.run_button.setText("正在处理...")
        self.status_label.setText("任务开始...")

        # 创建并启动线程
        self.worker_thread = QThread()
        self.worker = Worker(file_path, sheet_name)
        self.worker.moveToThread(self.worker_thread)

        # 连接信号
        self.worker.progress.connect(self.update_status)
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        self.worker_thread.started.connect(self.worker.run)

        self.worker_thread.start()

    def update_status(self, message):
        self.status_label.setText(message)

    def on_finished(self, message):
        self.status_label.setText(message)
        QMessageBox.information(self, "任务完成", message)
        self.cleanup_thread()

    def on_error(self, message):
        self.status_label.setText(f"错误: {message}")
        self.status_label.setStyleSheet("color: red;")
        QMessageBox.critical(self, "发生错误", message)
        self.cleanup_thread()

    def cleanup_thread(self):
        """恢复按钮状态并清理线程"""
        self.run_button.setEnabled(True)
        self.run_button.setText("开始清洗")
        self.status_label.setStyleSheet("color: blue;")

        if self.worker_thread and self.worker_thread.isRunning():
            self.worker_thread.quit()
            self.worker_thread.wait()
        self.worker_thread = None
        self.worker = None


# --- 5. 启动应用 ---
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())