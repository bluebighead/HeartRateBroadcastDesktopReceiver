import sys
import asyncio
import time
import json
import urllib.request
import urllib.error
from datetime import datetime
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QMessageBox, QStackedWidget, QTableWidget,
    QTableWidgetItem, QHeaderView, QAbstractItemView, QGraphicsOpacityEffect,
    QDialog, QCheckBox, QLineEdit, QComboBox, QSpinBox, QFormLayout,
    QMenuBar, QAction, QGroupBox
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer, QPropertyAnimation, QEasingCurve, pyqtProperty
from PyQt5.QtGui import QFont
from bleak import BleakClient, BleakScanner
import pyqtgraph as pg

APP_VERSION = "v1.0.0"
GITHUB_REPO = "bluebighead/HeartRateBroadcastDesktopReceiver"
GITHUB_RELEASES_API = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"

HEART_RATE_SERVICE_UUID = "0000180D-0000-1000-8000-00805F9B34FB"
HEART_RATE_MEASUREMENT_UUID = "00002A37-0000-1000-8000-00805F9B34FB"


class BleakWorker(QThread):
    heart_rate_received = pyqtSignal(int)
    status_changed = pyqtSignal(str)
    error_occurred = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.running = False
        self.client = None

    def run(self):
        self.running = True
        asyncio.run(self._run_ble())

    async def _run_ble(self):
        try:
            self.status_changed.emit("正在扫描心率设备...")
            
            device = await BleakScanner.find_device_by_filter(
                lambda d, ad: HEART_RATE_SERVICE_UUID.lower() in [s.lower() for s in ad.service_uuids]
            )
            
            if not device:
                self.error_occurred.emit("未找到心率设备，请确保设备已开启并靠近电脑")
                return
            
            self.status_changed.emit(f"已找到设备: {device.name or device.address}")
            
            async with BleakClient(device) as client:
                self.client = client
                self.status_changed.emit(f"已连接: {device.name or device.address}")
                
                await client.start_notify(
                    HEART_RATE_MEASUREMENT_UUID,
                    self._heart_rate_handler
                )
                
                while self.running:
                    await asyncio.sleep(0.1)
                
                await client.stop_notify(HEART_RATE_MEASUREMENT_UUID)
                
        except Exception as e:
            self.error_occurred.emit(f"蓝牙错误: {str(e)}")
        finally:
            self.status_changed.emit("已断开连接")

    def _heart_rate_handler(self, sender, data):
        flags = data[0]
        if flags & 0x01:
            heart_rate = int.from_bytes(data[1:3], byteorder='little')
        else:
            heart_rate = data[1]
        self.heart_rate_received.emit(heart_rate)

    def stop(self):
        self.running = False
        self.wait()


class HeartRateRecord:
    def __init__(self, heart_rate, timestamp):
        self.heart_rate = heart_rate
        self.timestamp = timestamp


class CalorieSettingsDialog(QDialog):
    def __init__(self, parent=None, current_settings=None):
        super().__init__(parent)
        self.setWindowTitle("卡路里设置")
        self.setMinimumWidth(350)
        self.settings = current_settings or {
            'enabled': False,
            'weight': 70,
            'age': 30,
            'gender': 'male'
        }
        self.init_ui()
        
    def init_ui(self):
        layout = QVBoxLayout(self)
        
        self.enable_checkbox = QCheckBox("是否开启实时卡路里")
        self.enable_checkbox.setChecked(self.settings['enabled'])
        self.enable_checkbox.stateChanged.connect(self.on_enable_changed)
        self.enable_checkbox.setStyleSheet("font-weight: bold; color: #2C3E50;")
        layout.addWidget(self.enable_checkbox)
        
        layout.addSpacing(10)
        
        form_group = QGroupBox("个人信息")
        form_layout = QFormLayout(form_group)
        
        self.weight_spin = QSpinBox()
        self.weight_spin.setRange(30, 200)
        self.weight_spin.setValue(int(self.settings['weight']))
        self.weight_spin.setSuffix(" kg")
        self.weight_spin.setEnabled(self.settings['enabled'])
        form_layout.addRow("体重:", self.weight_spin)
        
        self.age_spin = QSpinBox()
        self.age_spin.setRange(10, 100)
        self.age_spin.setValue(int(self.settings['age']))
        self.age_spin.setSuffix(" 岁")
        self.age_spin.setEnabled(self.settings['enabled'])
        form_layout.addRow("年龄:", self.age_spin)
        
        self.gender_combo = QComboBox()
        self.gender_combo.addItem("男", "male")
        self.gender_combo.addItem("女", "female")
        if self.settings['gender'] == 'female':
            self.gender_combo.setCurrentIndex(1)
        self.gender_combo.setEnabled(self.settings['enabled'])
        form_layout.addRow("性别:", self.gender_combo)
        
        layout.addWidget(form_group)
        
        layout.addSpacing(20)
        
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        cancel_button = QPushButton("取消")
        cancel_button.setMinimumWidth(80)
        cancel_button.setStyleSheet("""
            QPushButton {
                background-color: #95A5A6;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 8px;
            }
            QPushButton:hover {
                background-color: #7F8C8D;
            }
        """)
        cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(cancel_button)
        
        self.ok_button = QPushButton("确定")
        self.ok_button.setMinimumWidth(80)
        self.ok_button.setStyleSheet("""
            QPushButton {
                background-color: #27AE60;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 8px;
            }
            QPushButton:hover {
                background-color: #2ECC71;
            }
        """)
        self.ok_button.clicked.connect(self.on_ok_clicked)
        button_layout.addWidget(self.ok_button)
        
        layout.addLayout(button_layout)
        
        self.on_enable_changed(self.enable_checkbox.checkState())
    
    def on_enable_changed(self, state):
        enabled = state == Qt.Checked
        self.weight_spin.setEnabled(enabled)
        self.age_spin.setEnabled(enabled)
        self.gender_combo.setEnabled(enabled)
        
        if enabled:
            self.weight_spin.setStyleSheet("")
            self.age_spin.setStyleSheet("")
            self.gender_combo.setStyleSheet("")
        else:
            self.weight_spin.setStyleSheet("color: #BDC3C7;")
            self.age_spin.setStyleSheet("color: #BDC3C7;")
            self.gender_combo.setStyleSheet("color: #BDC3C7;")
    
    def on_ok_clicked(self):
        if self.enable_checkbox.isChecked():
            if self.weight_spin.value() <= 0:
                QMessageBox.warning(self, "提示", "请输入有效的体重")
                return
            if self.age_spin.value() <= 0:
                QMessageBox.warning(self, "提示", "请输入有效的年龄")
                return
        
        self.settings = {
            'enabled': self.enable_checkbox.isChecked(),
            'weight': self.weight_spin.value(),
            'age': self.age_spin.value(),
            'gender': self.gender_combo.currentData()
        }
        self.accept()
    
    def get_settings(self):
        return self.settings


class HeartAnimationLabel(QLabel):
    def __init__(self, parent=None):
        super().__init__("❤", parent)
        self.setFont(QFont("Arial", 48))
        self.setStyleSheet("color: #E74C3C;")
        self.setAlignment(Qt.AlignCenter)
        
        self._scale = 1.0
        self.animation = QPropertyAnimation(self, b"scale")
        self.animation.setDuration(300)
        self.animation.setEasingCurve(QEasingCurve.InOutQuad)
        
        self.beat_timer = QTimer()
        self.beat_timer.timeout.connect(self.beat)
        self.is_beating = False
        
    def get_scale(self):
        return self._scale
    
    def set_scale(self, value):
        self._scale = value
        font = self.font()
        base_size = 48
        new_size = int(base_size * value)
        font.setPointSize(new_size)
        self.setFont(font)
    
    scale = pyqtProperty(float, get_scale, set_scale)
    
    def start_beating(self, heart_rate=72):
        if heart_rate > 0:
            interval = int(60000 / heart_rate)
            self.beat_timer.start(max(400, interval))
        self.is_beating = True
    
    def stop_beating(self):
        self.beat_timer.stop()
        self.is_beating = False
        self.set_scale(1.0)
    
    def beat(self):
        self.animation.stop()
        self.animation.setStartValue(1.0)
        self.animation.setKeyValueAt(0.3, 1.3)
        self.animation.setKeyValueAt(0.6, 1.0)
        self.animation.setEndValue(1.0)
        self.animation.start()
    
    def update_heart_rate(self, heart_rate):
        if self.is_beating and heart_rate > 0:
            interval = int(60000 / heart_rate)
            self.beat_timer.setInterval(max(400, interval))


class HomePage(QWidget):
    heart_rate_recorded = pyqtSignal(int, datetime)
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.ble_worker = None
        self.last_heart_rate = None
        self.start_time = None
        self.current_heart_rate = 0
        self.total_calories = 0.0
        self.calorie_settings = {
            'enabled': False,
            'weight': 70,
            'age': 30,
            'gender': 'male'
        }
        self.init_ui()
        
        self.calorie_timer = QTimer()
        self.calorie_timer.timeout.connect(self.update_calories)

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)
        
        heart_layout = QHBoxLayout()
        heart_layout.setAlignment(Qt.AlignCenter)
        
        self.heart_animation = HeartAnimationLabel()
        heart_layout.addWidget(self.heart_animation)
        
        self.heart_rate_label = QLabel("--")
        self.heart_rate_label.setFont(QFont("Arial", 72, QFont.Bold))
        self.heart_rate_label.setAlignment(Qt.AlignCenter)
        self.heart_rate_label.setStyleSheet("color: #E74C3C;")
        heart_layout.addWidget(self.heart_rate_label)
        
        layout.addLayout(heart_layout)
        
        bpm_label = QLabel("BPM")
        bpm_label.setFont(QFont("Arial", 18))
        bpm_label.setAlignment(Qt.AlignCenter)
        bpm_label.setStyleSheet("color: #7F8C8D;")
        layout.addWidget(bpm_label)
        
        layout.addSpacing(10)
        
        self.calorie_label = QLabel("消耗卡路里: 0.0 kcal")
        self.calorie_label.setFont(QFont("Arial", 14))
        self.calorie_label.setAlignment(Qt.AlignCenter)
        self.calorie_label.setStyleSheet("color: #F39C12; font-weight: bold;")
        self.calorie_label.hide()
        layout.addWidget(self.calorie_label)
        
        self.duration_label = QLabel("运动时长: 00:00:00")
        self.duration_label.setFont(QFont("Arial", 12))
        self.duration_label.setAlignment(Qt.AlignCenter)
        self.duration_label.setStyleSheet("color: #7F8C8D;")
        layout.addWidget(self.duration_label)
        
        layout.addSpacing(15)
        
        self.status_label = QLabel("未连接")
        self.status_label.setFont(QFont("Arial", 12))
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet("color: #95A5A6;")
        layout.addWidget(self.status_label)
        
        layout.addSpacing(30)
        
        button_layout = QHBoxLayout()
        
        self.start_button = QPushButton("开始接收")
        self.start_button.setFont(QFont("Arial", 14))
        self.start_button.setMinimumSize(120, 45)
        self.start_button.setStyleSheet("""
            QPushButton {
                background-color: #27AE60;
                color: white;
                border: none;
                border-radius: 8px;
            }
            QPushButton:hover {
                background-color: #2ECC71;
            }
            QPushButton:disabled {
                background-color: #95A5A6;
            }
        """)
        self.start_button.clicked.connect(self.start_receiving)
        button_layout.addWidget(self.start_button)
        
        self.stop_button = QPushButton("停止接收")
        self.stop_button.setFont(QFont("Arial", 14))
        self.stop_button.setMinimumSize(120, 45)
        self.stop_button.setEnabled(False)
        self.stop_button.setStyleSheet("""
            QPushButton {
                background-color: #E74C3C;
                color: white;
                border: none;
                border-radius: 8px;
            }
            QPushButton:hover {
                background-color: #C0392B;
            }
            QPushButton:disabled {
                background-color: #95A5A6;
            }
        """)
        self.stop_button.clicked.connect(self.stop_receiving)
        button_layout.addWidget(self.stop_button)
        
        layout.addLayout(button_layout)
    
    def set_calorie_settings(self, settings):
        self.calorie_settings = settings
        if settings['enabled']:
            self.calorie_label.show()
        else:
            self.calorie_label.hide()
    
    def get_calorie_settings(self):
        return self.calorie_settings

    def start_receiving(self):
        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)
        self.heart_rate_label.setText("--")
        self.last_heart_rate = None
        self.start_time = time.time()
        self.total_calories = 0.0
        self.current_heart_rate = 0
        self.calorie_label.setText("消耗卡路里: 0.0 kcal")
        self.duration_label.setText("运动时长: 00:00:00")
        
        self.heart_animation.start_beating()
        self.calorie_timer.start(1000)
        
        self.ble_worker = BleakWorker()
        self.ble_worker.heart_rate_received.connect(self.update_heart_rate)
        self.ble_worker.status_changed.connect(self.update_status)
        self.ble_worker.error_occurred.connect(self.show_error)
        self.ble_worker.start()

    def stop_receiving(self):
        if self.ble_worker:
            self.ble_worker.stop()
            self.ble_worker = None
        
        self.heart_animation.stop_beating()
        self.calorie_timer.stop()
        
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self.heart_rate_label.setText("--")
        self.status_label.setText("未连接")
        self.last_heart_rate = None
        self.current_heart_rate = 0

    def update_heart_rate(self, heart_rate):
        self.heart_rate_label.setText(str(heart_rate))
        self.current_heart_rate = heart_rate
        self.heart_animation.update_heart_rate(heart_rate)
        
        if self.last_heart_rate is None or heart_rate != self.last_heart_rate:
            timestamp = datetime.now()
            self.heart_rate_recorded.emit(heart_rate, timestamp)
            self.last_heart_rate = heart_rate

    def update_calories(self):
        if self.start_time and self.current_heart_rate > 0 and self.calorie_settings['enabled']:
            weight = self.calorie_settings['weight']
            age = self.calorie_settings['age']
            gender = self.calorie_settings['gender']
            hr = self.current_heart_rate
            
            if gender == 'male':
                calories_per_minute = ((-55.0969 + 0.6309 * hr + 0.1988 * weight + 0.2017 * age) / 4.184)
            else:
                calories_per_minute = ((-20.4022 + 0.4472 * hr - 0.1263 * weight + 0.074 * age) / 4.184)
            
            calories_per_second = calories_per_minute / 60
            self.total_calories += calories_per_second
            self.calorie_label.setText(f"消耗卡路里: {self.total_calories:.1f} kcal")
        
        if self.start_time:
            elapsed = time.time() - self.start_time
            hours = int(elapsed // 3600)
            minutes = int((elapsed % 3600) // 60)
            seconds = int(elapsed % 60)
            self.duration_label.setText(f"运动时长: {hours:02d}:{minutes:02d}:{seconds:02d}")

    def update_status(self, status):
        self.status_label.setText(status)

    def show_error(self, message):
        QMessageBox.warning(self, "错误", message)
        self.stop_receiving()

    def cleanup(self):
        self.stop_receiving()


class ChartWindow(QWidget):
    def __init__(self, record_page, parent=None):
        super().__init__(parent)
        self.record_page = record_page
        self.setWindowTitle("心率折线图")
        self.setMinimumSize(800, 500)
        self.init_ui()
        
        self.update_timer = QTimer()
        self.update_timer.timeout.connect(self.update_chart)
        self.update_timer.start(500)

    def init_ui(self):
        layout = QVBoxLayout(self)
        
        pg.setConfigOptions(antialias=True)
        
        self.plot_widget = pg.PlotWidget()
        self.plot_widget.setBackground('w')
        self.plot_widget.setTitle("心率变化趋势", color='#2C3E50', size='14pt')
        self.plot_widget.setLabel('left', '心率', 'BPM', color='#2C3E50')
        self.plot_widget.setLabel('bottom', '序号', units='', color='#2C3E50')
        self.plot_widget.showGrid(x=True, y=True, alpha=0.3)
        self.plot_widget.setYRange(40, 200)
        
        self.curve = self.plot_widget.plot(
            pen=pg.mkPen(color='#E74C3C', width=2),
            symbol='o',
            symbolSize=8,
            symbolBrush='#E74C3C'
        )
        
        layout.addWidget(self.plot_widget)
        
        info_layout = QHBoxLayout()
        
        self.current_label = QLabel("当前心率: -- BPM")
        self.current_label.setFont(QFont("Arial", 12))
        self.current_label.setStyleSheet("color: #2C3E50;")
        info_layout.addWidget(self.current_label)
        
        self.count_label = QLabel("数据点: 0")
        self.count_label.setFont(QFont("Arial", 12))
        self.count_label.setStyleSheet("color: #2C3E50;")
        info_layout.addWidget(self.count_label)
        
        info_layout.addStretch()
        layout.addLayout(info_layout)

    def update_chart(self):
        records = self.record_page.records
        if not records:
            self.curve.setData([], [])
            self.current_label.setText("当前心率: -- BPM")
            self.count_label.setText("数据点: 0")
            return
        
        x_data = list(range(1, len(records) + 1))
        y_data = [r.heart_rate for r in records]
        
        self.curve.setData(x_data, y_data)
        
        if y_data:
            self.current_label.setText(f"当前心率: {y_data[-1]} BPM")
        self.count_label.setText(f"数据点: {len(records)}")

    def closeEvent(self, event):
        self.update_timer.stop()
        event.accept()


class RecordPage(QWidget):
    record_added = pyqtSignal()
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.records = []
        self.chart_window = None
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        
        title_label = QLabel("心率变化记录")
        title_label.setFont(QFont("Arial", 18, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("color: #2C3E50; margin: 10px;")
        layout.addWidget(title_label)
        
        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["序号", "心率 (BPM)", "时间"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.setStyleSheet("""
            QTableWidget {
                background-color: white;
                border: 1px solid #BDC3C7;
                border-radius: 5px;
                gridline-color: #ECF0F1;
            }
            QTableWidget::item {
                padding: 8px;
            }
            QTableWidget::item:selected {
                background-color: #3498DB;
                color: white;
            }
            QHeaderView::section {
                background-color: #34495E;
                color: white;
                padding: 8px;
                border: none;
                font-weight: bold;
            }
        """)
        layout.addWidget(self.table)
        
        button_layout = QHBoxLayout()
        
        chart_button = QPushButton("显示折线图")
        chart_button.setFont(QFont("Arial", 12))
        chart_button.setMinimumSize(120, 35)
        chart_button.setStyleSheet("""
            QPushButton {
                background-color: #3498DB;
                color: white;
                border: none;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #2980B9;
            }
        """)
        chart_button.clicked.connect(self.show_chart)
        button_layout.addWidget(chart_button)
        
        button_layout.addStretch()
        
        clear_button = QPushButton("清空记录")
        clear_button.setFont(QFont("Arial", 12))
        clear_button.setMinimumSize(100, 35)
        clear_button.setStyleSheet("""
            QPushButton {
                background-color: #E74C3C;
                color: white;
                border: none;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #C0392B;
            }
        """)
        clear_button.clicked.connect(self.clear_records)
        button_layout.addWidget(clear_button)
        
        layout.addLayout(button_layout)

    def show_chart(self):
        if self.chart_window is None or not self.chart_window.isVisible():
            self.chart_window = ChartWindow(self)
            self.chart_window.setWindowTitle("心率折线图")
            self.chart_window.show()
        else:
            self.chart_window.raise_()
            self.chart_window.activateWindow()

    def add_record(self, heart_rate, timestamp):
        record = HeartRateRecord(heart_rate, timestamp)
        self.records.append(record)
        
        row = self.table.rowCount()
        self.table.insertRow(row)
        
        self.table.setItem(row, 0, QTableWidgetItem(str(row + 1)))
        self.table.setItem(row, 1, QTableWidgetItem(str(heart_rate)))
        self.table.setItem(row, 2, QTableWidgetItem(timestamp.strftime("%Y-%m-%d %H:%M:%S")))
        
        self.table.scrollToBottom()
        self.record_added.emit()

    def clear_records(self):
        self.records.clear()
        self.table.setRowCount(0)


class HeartRateWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle(f"心率广播接收器 {APP_VERSION}")
        self.setMinimumSize(500, 400)
        
        menubar = self.menuBar()
        menubar.setStyleSheet("""
            QMenuBar {
                background-color: #34495E;
                color: white;
                padding: 5px;
            }
            QMenuBar::item {
                background-color: transparent;
                padding: 5px 15px;
                border-radius: 3px;
            }
            QMenuBar::item:selected {
                background-color: #3498DB;
            }
            QMenu {
                background-color: #34495E;
                color: white;
                border: 1px solid #2C3E50;
            }
            QMenu::item {
                padding: 8px 25px;
            }
            QMenu::item:selected {
                background-color: #3498DB;
            }
        """)
        
        settings_menu = menubar.addMenu("设置")
        
        calorie_action = QAction("显示动态实时卡路里", self)
        calorie_action.triggered.connect(self.show_calorie_settings)
        settings_menu.addAction(calorie_action)
        
        settings_menu.addSeparator()
        
        update_action = QAction("检查更新", self)
        update_action.triggered.connect(self.check_update)
        settings_menu.addAction(update_action)
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        nav_bar = QWidget()
        nav_bar.setFixedHeight(50)
        nav_bar.setStyleSheet("background-color: #2C3E50;")
        nav_layout = QHBoxLayout(nav_bar)
        nav_layout.setContentsMargins(20, 0, 20, 0)
        
        self.home_button = QPushButton("首页")
        self.home_button.setFont(QFont("Arial", 12))
        self.home_button.setCheckable(True)
        self.home_button.setChecked(True)
        self.home_button.setFixedHeight(40)
        self.home_button.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 0 20px;
            }
            QPushButton:hover {
                background-color: #34495E;
            }
            QPushButton:checked {
                background-color: #3498DB;
            }
        """)
        self.home_button.clicked.connect(lambda: self.switch_page(0))
        nav_layout.addWidget(self.home_button)
        
        self.record_button = QPushButton("记录")
        self.record_button.setFont(QFont("Arial", 12))
        self.record_button.setCheckable(True)
        self.record_button.setFixedHeight(40)
        self.record_button.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 0 20px;
            }
            QPushButton:hover {
                background-color: #34495E;
            }
            QPushButton:checked {
                background-color: #3498DB;
            }
        """)
        self.record_button.clicked.connect(lambda: self.switch_page(1))
        nav_layout.addWidget(self.record_button)
        
        nav_layout.addStretch()
        main_layout.addWidget(nav_bar)
        
        self.stack = QStackedWidget()
        
        self.home_page = HomePage()
        self.home_page.heart_rate_recorded.connect(self.on_heart_rate_recorded)
        self.stack.addWidget(self.home_page)
        
        self.record_page = RecordPage()
        self.stack.addWidget(self.record_page)
        
        main_layout.addWidget(self.stack)

    def show_calorie_settings(self):
        dialog = CalorieSettingsDialog(self, self.home_page.get_calorie_settings())
        if dialog.exec_() == QDialog.Accepted:
            settings = dialog.get_settings()
            self.home_page.set_calorie_settings(settings)

    def check_update(self):
        try:
            req = urllib.request.Request(GITHUB_RELEASES_API)
            req.add_header('User-Agent', f'HeartRateReceiver/{APP_VERSION}')
            
            with urllib.request.urlopen(req, timeout=10) as response:
                data = json.loads(response.read().decode('utf-8'))
                latest_version = data.get('tag_name', '')
                release_url = data.get('html_url', '')
                release_notes = data.get('body', '')
            
            if not latest_version:
                QMessageBox.warning(self, "检查更新", "无法获取版本信息")
                return
            
            if latest_version == APP_VERSION:
                QMessageBox.information(
                    self, "检查更新",
                    f"当前已是最新版本\n\n当前版本: {APP_VERSION}"
                )
            else:
                msg = f"发现新版本!\n\n当前版本: {APP_VERSION}\n最新版本: {latest_version}"
                if release_notes:
                    msg += f"\n\n更新内容:\n{release_notes[:200]}"
                    if len(release_notes) > 200:
                        msg += "..."
                
                reply = QMessageBox.question(
                    self, "检查更新", msg,
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.Yes
                )
                
                if reply == QMessageBox.Yes and release_url:
                    import webbrowser
                    webbrowser.open(release_url)
                    
        except urllib.error.HTTPError as e:
            if e.code == 404:
                QMessageBox.information(
                    self, "检查更新",
                    f"当前版本: {APP_VERSION}\n\n暂无发布版本\n\n请前往 GitHub 发布第一个版本"
                )
            else:
                QMessageBox.warning(self, "检查更新", f"HTTP错误: {e.code} {e.reason}")
        except urllib.error.URLError as e:
            QMessageBox.warning(self, "检查更新", f"网络错误: {str(e)}\n\n请检查网络连接")
        except Exception as e:
            QMessageBox.warning(self, "检查更新", f"检查更新失败: {str(e)}")

    def switch_page(self, index):
        self.stack.setCurrentIndex(index)
        self.home_button.setChecked(index == 0)
        self.record_button.setChecked(index == 1)

    def on_heart_rate_recorded(self, heart_rate, timestamp):
        self.record_page.add_record(heart_rate, timestamp)

    def closeEvent(self, event):
        self.home_page.cleanup()
        event.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = HeartRateWindow()
    window.show()
    sys.exit(app.exec_())
