#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
汽水音乐自动看广告脚本 - GUI 版本 (PyQt5)
可视化配置所有参数、实时日志、开始/暂停/停止控制

使用方法：
  双击 启动GUI.bat 或运行: python qishui_gui.py
"""

import ctypes
import os
import sys
import time
import json
import logging
import threading
from pathlib import Path
from datetime import datetime

# ============================================================
#  DPI 感知
# ============================================================
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PROCESS_PER_MONITOR_DPI_AWARE
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

# 导入 PyQt5
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QPushButton, QLabel, QLineEdit, QSpinBox, QDoubleSpinBox, QCheckBox,
    QTextEdit, QGroupBox, QTabWidget, QSplitter, QScrollArea, QMessageBox,
    QFrame, QSizePolicy
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer, QMutex
from PyQt5.QtGui import QFont, QTextCursor, QColor

# 导入核心模块
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import qishui_ad_watcher as core


# ============================================================
#  配置文件路径
# ============================================================
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "gui_config.json")


def save_config_to_file(config_dict: dict):
    """保存配置到 JSON 文件"""
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config_dict, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"保存配置失败: {e}")


def load_config_from_file() -> dict:
    """从 JSON 文件加载配置"""
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception as e:
        print(f"加载配置失败: {e}")
    return {}


# ============================================================
#  Qt 日志 Handler —— 将日志输出到 QTextEdit
# ============================================================
class QtLogHandler(logging.Handler):
    """将日志通过信号发送到主线程的 QTextEdit"""

    def __init__(self, callback):
        super().__init__()
        self.callback = callback

    def emit(self, record):
        try:
            msg = self.format(record)
            level = record.levelname
            self.callback(msg, level)
        except Exception:
            pass


# ============================================================
#  运行线程 —— 在独立线程中运行 AdWatcher
# ============================================================
class WatcherThread(QThread):
    """在独立线程中运行 AdWatcher.run()"""
    log_signal = pyqtSignal(str, str)   # (message, level)
    finished_signal = pyqtSignal(dict)  # stats

    def __init__(self, config):
        super().__init__()
        self.config = config
        self.watcher = None

    def run(self):
        try:
            self.watcher = core.AdWatcher(self.config)
            # 安装日志 handler
            handler = QtLogHandler(self._emit_log)
            handler.setLevel(logging.INFO)
            handler.setFormatter(logging.Formatter(
                '%(asctime)s [%(levelname)s] %(message)s', datefmt='%H:%M:%S'))
            core.logger.addHandler(handler)
            try:
                self.watcher.run()
            finally:
                core.logger.removeHandler(handler)
            self.finished_signal.emit(self.watcher.stats)
        except Exception as e:
            self.log_signal.emit(f"运行异常: {e}", "ERROR")
            self.finished_signal.emit({})

    def _emit_log(self, msg, level):
        self.log_signal.emit(msg, level)

    def stop(self):
        if self.watcher:
            self.watcher.stop()

    def pause(self):
        if self.watcher:
            self.watcher.pause()

    def resume(self):
        if self.watcher:
            self.watcher.resume()

    def get_stats(self):
        if self.watcher:
            return self.watcher.stats
        return {}


# ============================================================
#  测试线程
# ============================================================
class TestThread(QThread):
    log_signal = pyqtSignal(str, str)
    finished_signal = pyqtSignal()

    def __init__(self, config, mode="test"):
        super().__init__()
        self.config = config
        self.mode = mode

    def run(self):
        handler = QtLogHandler(self._emit_log)
        handler.setLevel(logging.INFO)
        handler.setFormatter(logging.Formatter(
            '%(asctime)s [%(levelname)s] %(message)s', datefmt='%H:%M:%S'))
        core.logger.addHandler(handler)
        try:
            if self.mode == "test":
                core.run_test_mode(self.config)
            elif self.mode == "list":
                core.run_list_windows_mode()
        except Exception as e:
            self.log_signal.emit(f"异常: {e}", "ERROR")
        finally:
            core.logger.removeHandler(handler)
            self.finished_signal.emit()

    def _emit_log(self, msg, level):
        self.log_signal.emit(msg, level)


# ============================================================
#  输入框子类 —— 带 hint 提示
# ============================================================
class ConfigRow(QWidget):
    """一行配置：标签 + 输入框 + 提示"""

    def __init__(self, label_text, var_setter, var_getter, hint="", input_type="str"):
        super().__init__()
        self.var_setter = var_setter
        self.var_getter = var_getter
        self.input_type = input_type

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        self.label = QLabel(label_text)
        self.label.setFixedWidth(160)
        self.label.setToolTip(hint)

        if input_type == "int":
            self.input = QSpinBox()
            self.input.setRange(-99999, 999999)
        elif input_type == "double":
            self.input = QDoubleSpinBox()
            self.input.setRange(-9999.0, 9999.0)
            self.input.setDecimals(3)
            self.input.setSingleStep(0.1)
        else:
            self.input = QLineEdit()

        self.input.setValue(var_getter()) if input_type in ("int", "double") else self.input.setText(str(var_getter()))
        self.input.setToolTip(hint)

        layout.addWidget(self.label)
        layout.addWidget(self.input, 1)

        if hint:
            self.hint_label = QLabel(hint)
            self.hint_label.setStyleSheet("color: gray; font-size: 9pt;")
            self.hint_label.setWordWrap(True)
            self.hint_label.setMaximumWidth(250)

    def apply_to_config(self, config, attr_name):
        """把输入框的值写入 config 对象"""
        if self.input_type == "int":
            setattr(config, attr_name, int(self.input.value()))
        elif self.input_type == "double":
            setattr(config, attr_name, float(self.input.value()))
        else:
            setattr(config, attr_name, self.input.text().strip())

    def get_value(self):
        if self.input_type == "int":
            return int(self.input.value())
        elif self.input_type == "double":
            return float(self.input.value())
        else:
            return self.input.text().strip()

    def set_value(self, val):
        if self.input_type == "int":
            self.input.setValue(int(val))
        elif self.input_type == "double":
            self.input.setValue(float(val))
        else:
            self.input.setText(str(val))


# ============================================================
#  主窗口
# ============================================================
class QishuiGUI(QMainWindow):
    """汽水音乐广告自动化 GUI 主窗口"""

    def __init__(self):
        super().__init__()
        self.watcher_thread = None
        self.test_thread = None
        self.config_rows = {}
        self.checkboxes = {}
        self.saved_config = load_config_from_file()

        self._init_ui()
        self._load_config_to_ui()
        self._setup_logging()

        # 统计刷新定时器
        self.stats_timer = QTimer(self)
        self.stats_timer.timeout.connect(self._update_stats)
        self.stats_timer.start(1000)

    # ----------------------------------------------------------
    #  UI 构建
    # ----------------------------------------------------------
    def _init_ui(self):
        self.setWindowTitle("汽水音乐自动看广告 - 控制面板")
        self.resize(1000, 780)
        self.setMinimumSize(850, 650)

        # 中央 widget
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(8, 8, 8, 8)

        # === 顶部控制栏 ===
        ctrl_layout = QHBoxLayout()
        ctrl_layout.setSpacing(6)

        self.btn_start = QPushButton("开始运行")
        self.btn_start.setFixedWidth(100)
        self.btn_start.setStyleSheet("QPushButton { background-color: #4CAF50; color: white; font-weight: bold; padding: 6px; }"
                                     "QPushButton:hover { background-color: #45a049; }"
                                     "QPushButton:disabled { background-color: #a5d6a7; }")
        self.btn_start.clicked.connect(self.on_start)

        self.btn_pause = QPushButton("暂停")
        self.btn_pause.setFixedWidth(70)
        self.btn_pause.setEnabled(False)
        self.btn_pause.clicked.connect(self.on_pause)

        self.btn_stop = QPushButton("停止")
        self.btn_stop.setFixedWidth(70)
        self.btn_stop.setStyleSheet("QPushButton { background-color: #f44336; color: white; font-weight: bold; padding: 6px; }"
                                    "QPushButton:hover { background-color: #d32f2f; }"
                                    "QPushButton:disabled { background-color: #ef9a9a; }")
        self.btn_stop.setEnabled(False)
        self.btn_stop.clicked.connect(self.on_stop)

        self.btn_save = QPushButton("保存配置")
        self.btn_save.setFixedWidth(80)
        self.btn_save.clicked.connect(self.on_save_config)

        self.btn_test = QPushButton("测试截图+OCR")
        self.btn_test.setFixedWidth(120)
        self.btn_test.clicked.connect(self.on_test)

        self.btn_list = QPushButton("列出窗口")
        self.btn_list.setFixedWidth(80)
        self.btn_list.clicked.connect(self.on_list_windows)

        self.lbl_status = QLabel("状态: 未运行")
        self.lbl_status.setStyleSheet("font-weight: bold; color: #555;")

        ctrl_layout.addWidget(self.btn_start)
        ctrl_layout.addWidget(self.btn_pause)
        ctrl_layout.addWidget(self.btn_stop)
        ctrl_layout.addSpacing(10)
        ctrl_layout.addWidget(self.btn_save)
        ctrl_layout.addWidget(self.btn_test)
        ctrl_layout.addWidget(self.btn_list)
        ctrl_layout.addStretch()
        ctrl_layout.addWidget(self.lbl_status)

        main_layout.addLayout(ctrl_layout)

        # === 主体：左侧配置 + 右侧日志 ===
        splitter = QSplitter(Qt.Horizontal)

        # 左侧：配置面板（带滚动）
        config_widget = QWidget()
        config_layout = QVBoxLayout(config_widget)
        config_layout.setContentsMargins(0, 0, 0, 0)

        # 创建带滚动的配置区
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_content = QWidget()
        self._build_config_panel(scroll_content)
        scroll.setWidget(scroll_content)
        config_layout.addWidget(scroll)

        splitter.addWidget(config_widget)

        # 右侧：日志区
        log_widget = QWidget()
        log_layout = QVBoxLayout(log_widget)
        log_layout.setContentsMargins(0, 0, 0, 0)

        log_label = QLabel("运行日志")
        log_label.setStyleSheet("font-weight: bold; padding: 2px;")
        log_layout.addWidget(log_label)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont("Consolas", 9))
        self.log_text.setStyleSheet("""
            QTextEdit {
                background-color: #1e1e1e;
                color: #cccccc;
                border: 1px solid #3c3c3c;
            }
        """)
        log_layout.addWidget(self.log_text)

        # 清空日志按钮
        btn_clear_log = QPushButton("清空日志")
        btn_clear_log.setFixedWidth(80)
        btn_clear_log.clicked.connect(self.log_text.clear)
        log_layout.addWidget(btn_clear_log, alignment=Qt.AlignRight)

        splitter.addWidget(log_widget)
        splitter.setSizes([400, 600])

        main_layout.addWidget(splitter, 1)

        # === 底部统计 ===
        stats_group = QGroupBox("统计信息")
        stats_layout = QHBoxLayout(stats_group)
        self.lbl_stats = QLabel("广告: 0 | 取消直播: 0 | 退出直播间: 0 | 错误: 0 | 运行: 0秒")
        self.lbl_stats.setFont(QFont("Consolas", 10))
        self.lbl_stats.setStyleSheet("padding: 4px;")
        stats_layout.addWidget(self.lbl_stats)
        main_layout.addWidget(stats_group)

    def _build_config_panel(self, parent):
        """构建参数配置面板"""
        layout = QVBoxLayout(parent)
        layout.setSpacing(6)
        c = core.Config()

        def add_group(title):
            group = QGroupBox(title)
            g_layout = QVBoxLayout(group)
            g_layout.setSpacing(4)
            layout.addWidget(group)
            return g_layout

        def add_row(g_layout, key, label_text, hint="", input_type="str"):
            row = ConfigRow(label_text, None, lambda: getattr(c, key), hint, input_type)
            g_layout.addWidget(row)
            self.config_rows[key] = (row, input_type)

        # --- 基本设置 ---
        g = add_group("基本设置")
        add_row(g, "window_title", "窗口标题关键词", "投屏窗口标题（如：投屏、手机投屏），勿填'汽水'否则会匹配到本程序窗口", "str")
        add_row(g, "scan_interval", "扫描间隔(秒)", "越小越灵敏，直播倒计时需快", "double")
        add_row(g, "max_ads", "广告数量上限", "0=无限，达到后自动暂停", "int")
        add_row(g, "ad_timeout", "广告超时(秒)", "超过此时间未完成则放弃", "int")
        add_row(g, "ad_load_wait", "广告加载等待(秒)", "点击领时长后等待时间", "double")
        add_row(g, "reward_timeout", "奖励弹窗超时(秒)", "等待领取奖励的最大时间", "int")
        add_row(g, "click_delay", "点击后延迟(秒)", "每次点击后的等待", "double")

        # --- 视频广告点击位置 ---
        g = add_group("视频广告点击位置")
        add_row(g, "x_button_offset", "领取成功X按钮偏移", "点击'领取成功'右侧偏移像素，OCR常连读，负值向左", "int")

        # --- 直播广告点击位置 ---
        g = add_group("直播广告点击位置")
        add_row(g, "live_cancel_offset", "取消按钮偏移", "OCR把'4s后进入直播间取消'识别为整体时，点右边缘的偏移", "int")

        # --- 直播间退出点击位置 ---
        g = add_group("直播间退出点击位置")
        add_row(g, "live_x_viewer_offset", "观众数右侧偏移", "X按钮在观众数右侧的像素偏移", "int")
        add_row(g, "live_x_follow_offset", "关注右侧偏移", "X按钮在'关注'右侧的像素偏移", "int")
        add_row(g, "live_x_fixed_x_ratio", "固定位置X比例", "兜底：窗口宽度的比例(0-1)", "double")
        add_row(g, "live_x_fixed_y_ratio", "固定位置Y比例", "兜底：窗口高度的比例(0-1)", "double")
        add_row(g, "live_room_timeout", "直播间退出超时(秒)", "误入直播间后多久没退出就报错", "int")

        # --- 其他选项 ---
        g = add_group("其他选项")
        cb_debug = QCheckBox("保存调试截图")
        cb_debug.setChecked(c.debug)
        cb_debug.setToolTip("在 debug_screenshots 目录保存标注截图和点击位置")
        g.addWidget(cb_debug)
        self.checkboxes['debug'] = cb_debug

        cb_dry_run = QCheckBox("试运行模式（只检测不点击）")
        cb_dry_run.setChecked(c.dry_run)
        cb_dry_run.setToolTip("只做截图和OCR识别，不执行任何点击操作")
        g.addWidget(cb_dry_run)
        self.checkboxes['dry_run'] = cb_dry_run

        layout.addStretch()

    # ----------------------------------------------------------
    #  加载/保存配置
    # ----------------------------------------------------------
    def _load_config_to_ui(self):
        """把保存的配置加载到 UI"""
        saved = self.saved_config
        if not saved:
            return
        for key, (row, input_type) in self.config_rows.items():
            if key in saved:
                try:
                    row.set_value(saved[key])
                except (ValueError, TypeError):
                    pass
        if 'debug' in saved:
            self.checkboxes['debug'].setChecked(bool(saved['debug']))
        if 'dry_run' in saved:
            self.checkboxes['dry_run'].setChecked(bool(saved['dry_run']))

    def _save_current_config(self):
        """保存当前配置到文件"""
        config_dict = {}
        for key, (row, input_type) in self.config_rows.items():
            config_dict[key] = row.get_value()
        config_dict['debug'] = self.checkboxes['debug'].isChecked()
        config_dict['dry_run'] = self.checkboxes['dry_run'].isChecked()
        save_config_to_file(config_dict)

    def on_save_config(self):
        self._save_current_config()
        QMessageBox.information(self, "保存配置", "配置已保存到 gui_config.json")

    # ----------------------------------------------------------
    #  从 UI 构建 Config
    # ----------------------------------------------------------
    def _build_config(self) -> core.Config:
        config = core.Config()
        for key, (row, input_type) in self.config_rows.items():
            row.apply_to_config(config, key)
        config.debug = self.checkboxes['debug'].isChecked()
        config.dry_run = self.checkboxes['dry_run'].isChecked()
        return config

    # ----------------------------------------------------------
    #  日志系统
    # ----------------------------------------------------------
    def _setup_logging(self):
        """配置核心模块的日志"""
        core.logger.setLevel(logging.INFO)
        core.logger.handlers.clear()

        # 文件 handler
        file_handler = logging.FileHandler('qishui_ad_watcher.log', encoding='utf-8')
        file_handler.setLevel(logging.INFO)
        file_handler.setFormatter(logging.Formatter(
            '%(asctime)s [%(levelname)s] %(message)s', datefmt='%H:%M:%S'))
        core.logger.addHandler(file_handler)

    def _append_log(self, msg, level):
        """向日志区追加一条消息（线程安全，通过信号调用）"""
        color_map = {
            "INFO": "#cccccc",
            "WARNING": "#ffcc00",
            "ERROR": "#ff6b6b",
            "DEBUG": "#888888",
        }
        color = color_map.get(level, "#cccccc")
        # 保留最多 1000 行
        if self.log_text.document().blockCount() > 1000:
            cursor = self.log_text.textCursor()
            cursor.movePosition(QTextCursor.Start)
            cursor.movePosition(QTextCursor.Down, QTextCursor.KeepAnchor, 200)
            cursor.removeSelectedText()
        self.log_text.append(f'<span style="color:{color};">{msg}</span>')
        self.log_text.moveCursor(QTextCursor.End)

    # ----------------------------------------------------------
    #  按钮回调
    # ----------------------------------------------------------
    def on_start(self):
        """开始运行"""
        if self.watcher_thread and self.watcher_thread.isRunning():
            QMessageBox.information(self, "提示", "脚本已在运行中")
            return

        config = self._build_config()
        self._save_current_config()

        self.watcher_thread = WatcherThread(config)
        self.watcher_thread.log_signal.connect(self._append_log)
        self.watcher_thread.finished_signal.connect(self._on_watcher_finished)
        self.watcher_thread.start()

        self.btn_start.setEnabled(False)
        self.btn_pause.setEnabled(True)
        self.btn_pause.setText("暂停")
        self.btn_stop.setEnabled(True)
        self.lbl_status.setText("状态: 运行中")
        self.lbl_status.setStyleSheet("font-weight: bold; color: #4CAF50;")

    def _on_watcher_finished(self, stats):
        """watcher 运行结束"""
        self.btn_start.setEnabled(True)
        self.btn_pause.setEnabled(False)
        self.btn_pause.setText("暂停")
        self.btn_stop.setEnabled(False)
        self.lbl_status.setText("状态: 已停止")
        self.lbl_status.setStyleSheet("font-weight: bold; color: #555;")

    def on_pause(self):
        """暂停/恢复"""
        if not self.watcher_thread or not self.watcher_thread.isRunning():
            return
        if self.btn_pause.text() == "暂停":
            self.watcher_thread.pause()
            self.btn_pause.setText("恢复")
            self.lbl_status.setText("状态: 已暂停")
            self.lbl_status.setStyleSheet("font-weight: bold; color: #ff9800;")
        else:
            self.watcher_thread.resume()
            self.btn_pause.setText("暂停")
            self.lbl_status.setText("状态: 运行中")
            self.lbl_status.setStyleSheet("font-weight: bold; color: #4CAF50;")

    def on_stop(self):
        """停止"""
        if self.watcher_thread and self.watcher_thread.isRunning():
            self.watcher_thread.stop()
            self.lbl_status.setText("状态: 正在停止...")
            self.lbl_status.setStyleSheet("font-weight: bold; color: #f44336;")

    def on_test(self):
        """测试截图 + OCR"""
        if self.test_thread and self.test_thread.isRunning():
            return
        config = self._build_config()
        self.test_thread = TestThread(config, "test")
        self.test_thread.log_signal.connect(self._append_log)
        self.btn_test.setEnabled(False)
        self.test_thread.finished_signal.connect(lambda: self.btn_test.setEnabled(True))
        self.test_thread.start()

    def on_list_windows(self):
        """列出窗口"""
        if self.test_thread and self.test_thread.isRunning():
            return
        self.test_thread = TestThread(None, "list")
        self.test_thread.log_signal.connect(self._append_log)
        self.btn_list.setEnabled(False)
        self.test_thread.finished_signal.connect(lambda: self.btn_list.setEnabled(True))
        self.test_thread.start()

    # ----------------------------------------------------------
    #  统计刷新
    # ----------------------------------------------------------
    def _update_stats(self):
        """定时刷新统计信息"""
        if self.watcher_thread and self.watcher_thread.isRunning():
            s = self.watcher_thread.get_stats()
            elapsed = time.time() - s.get('start_time', time.time()) if s.get('start_time') else 0
            self.lbl_stats.setText(
                f"广告: {s.get('ads_watched', 0)} | "
                f"取消直播: {s.get('live_ads_cancelled', 0)} | "
                f"退出直播间: {s.get('live_rooms_exited', 0)} | "
                f"错误: {s.get('errors', 0)} | "
                f"运行: {elapsed:.0f}秒"
            )

    # ----------------------------------------------------------
    #  窗口关闭
    # ----------------------------------------------------------
    def closeEvent(self, event):
        if self.watcher_thread and self.watcher_thread.isRunning():
            reply = QMessageBox.question(
                self, "退出确认",
                "脚本正在运行，确定要退出吗？",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                self.watcher_thread.stop()
                self.watcher_thread.wait(3000)
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()


# ============================================================
#  入口
# ============================================================
def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = QishuiGUI()
    window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
