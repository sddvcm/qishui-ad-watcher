#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
汽水音乐自动看广告脚本
适用于 vivo办公套件投屏到电脑的场景

工作原理：
  截图投屏窗口 → OCR识别文字 → 状态机判断 → 自动点击

整体流程：
  1. 在福利中心点击"领时长"，打开广告
  2. 广告播放中：
     - 直播广告：检测到"取消"按钮(5秒倒计时)，立即点击取消
     - 视频广告：等待右上角倒计时结束
  3. 检测到"领取成功"，点击右侧X按钮
  4. 弹窗中点击"领取奖励"，自动进入下一个广告
  5. 循环执行

使用方法：
  1. pip install rapidocr-onnxruntime pywin32 keyboard mss Pillow numpy
  2. 打开 vivo办公套件，投屏手机，打开汽水音乐福利中心
  3. 运行: python qishui_ad_watcher.py
  4. 按 ESC 停止

命令行参数：
  --window <关键词>    指定窗口标题关键词（默认 投屏，匹配"手机投屏"窗口）
  --interval <秒>      扫描间隔（默认 0.5）
  --max-ads <数量>     最大广告数（默认 0=无限）
  --debug              保存调试截图
  --no-debug           关闭调试截图
  --list-windows       列出所有窗口
  --test               测试模式：截图+OCR，不点击
  --dry-run            检测但不点击
"""

import ctypes
import io
import logging
import os
import sys
import time
import threading
from datetime import datetime
from pathlib import Path
from dataclasses import dataclass, field
from typing import Optional, List, Tuple

# ============================================================
#  控制台编码：bat 里用 chcp 65001 设了 UTF-8，Python 同步用 UTF-8
# ============================================================
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

# ============================================================
#  DPI 感知：确保在高DPI屏幕上坐标准确
# ============================================================
ctypes.windll.user32.SetProcessDPIAware()

import win32gui
import win32con
import win32api
import win32ui
import mss
from PIL import Image, ImageDraw
import numpy as np

# PrintWindow 标志：捕获 DirectX/DirectComposition 渲染的内容
PW_RENDERFULLCONTENT = 2

# ============================================================
#  OCR 引擎
# ============================================================
try:
    from rapidocr_onnxruntime import RapidOCR
    HAS_RAPIDOCR = True
except ImportError:
    HAS_RAPIDOCR = False

# ============================================================
#  热键
# ============================================================
try:
    import keyboard
    HAS_KEYBOARD = True
except ImportError:
    HAS_KEYBOARD = False


# ============================================================
#  配置
# ============================================================
@dataclass
class Config:
    # 窗口标题关键词（不区分大小写）—— 默认匹配"手机投屏"窗口
    window_title: str = "投屏"

    # 扫描间隔（秒）—— 直播广告只有5秒倒计时，必须快
    scan_interval: float = 0.3

    # 广告超时时间（秒）—— 超过此时间未检测到完成则返回空闲
    ad_timeout: int = 120

    # 点击后延迟（秒）
    click_delay: float = 0.3

    # 广告加载等待时间（秒）
    ad_load_wait: float = 3.0

    # 奖励弹窗超时（秒）
    reward_timeout: int = 15

    # === 点击位置偏移参数（像素）===

    # 视频广告：点击"领取成功"右侧X按钮的偏移
    # OCR 常把"领取成功×"识别为一个整体，X就在文本右边缘，偏移设小值即可
    x_button_offset: int = -5

    # 直播广告：取消按钮偏移（当OCR把"4s后进入直播间取消"识别为整体时，点右边缘的偏移）
    live_cancel_offset: int = -10

    # 直播间X按钮：通过观众数定位时，X在观众数右侧的偏移
    live_x_viewer_offset: int = 40

    # 直播间X按钮：通过"关注"按钮定位时，X在关注右侧的偏移
    live_x_follow_offset: int = 200

    # 直播间X按钮：固定位置兜底，窗口宽度的比例
    live_x_fixed_x_ratio: float = 0.90
    live_x_fixed_y_ratio: float = 0.10

    # OCR最大图片宽度（超过此宽度会缩放，以提高速度）
    max_image_width: int = 1000

    # 区域过滤比率 [x_start, y_start, x_end, y_end]（相对于窗口截图）
    # 右上角区域：检测"领取成功"、倒计时
    region_top_right: tuple = (0.55, 0.0, 1.0, 0.18)
    # 中心区域：检测"5秒后进入直播间"、"取消"
    region_center: tuple = (0.1, 0.25, 0.9, 0.75)
    # 直播间X按钮区域：右上角，检测并点击退出直播间
    region_live_close: tuple = (0.75, 0.03, 1.0, 0.12)

    # 直播间退出超时（秒）—— 误入直播间后多久没退出就报错
    live_room_timeout: int = 20

    # 调试模式
    debug: bool = True
    debug_dir: str = "debug_screenshots"

    # 最大广告数（0=无限）
    max_ads: int = 0

    # 试运行模式（检测但不点击）
    dry_run: bool = False


# ============================================================
#  日志
# ============================================================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%H:%M:%S',
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler('qishui_ad_watcher.log', encoding='utf-8'),
    ]
)
logger = logging.getLogger("qishui")


# ============================================================
#  鼠标点击工具
# ============================================================
MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004


def mouse_click(x: int, y: int):
    """在屏幕绝对坐标处模拟鼠标左键点击"""
    ctypes.windll.user32.SetCursorPos(x, y)
    time.sleep(0.02)
    ctypes.windll.user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(0.05)
    ctypes.windll.user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)


def mouse_click_postmessage(hwnd: int, x: int, y: int):
    """通过 PostMessage 向窗口发送点击消息（后台点击，不移动鼠标）"""
    lparam = win32api.MAKELONG(int(x), int(y))
    win32api.PostMessageW(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
    time.sleep(0.05)
    win32api.PostMessageW(hwnd, win32con.WM_LBUTTONUP, 0, lparam)


# ============================================================
#  OCR 引擎封装
# ============================================================
class OCREngine:
    """RapidOCR 封装，提供文字识别和查找功能"""

    def __init__(self, config: Config):
        self.config = config
        self.engine = None

        if not HAS_RAPIDOCR:
            raise RuntimeError(
                "未安装 rapidocr-onnxruntime\n"
                "请运行: pip install rapidocr-onnxruntime"
            )

        logger.info("正在加载 OCR 模型（首次运行需下载，请耐心等待）...")
        self.engine = RapidOCR()
        logger.info("OCR 模型加载完成")

    def detect(self, image: Image.Image) -> List[dict]:
        """
        对图片进行 OCR 识别

        返回结果列表，每项包含:
          text: 识别到的文字
          x, y, w, h: 边界框（相对于原始图片）
          center: 中心点 (x, y)
          right: 右边缘中点 (x, y)
          score: 置信度 0~1
        """
        # 如果图片过大，缩放以提高速度
        orig_width = image.width
        if orig_width > self.config.max_image_width:
            scale = self.config.max_image_width / orig_width
            new_size = (self.config.max_image_width, int(image.height * scale))
            resized = image.resize(new_size, Image.LANCZOS)
        else:
            scale = 1.0
            resized = image

        img_array = np.array(resized)
        result, elapse = self.engine(img_array)

        if not result:
            return []

        results = []
        for box, text, score in result:
            text = text.strip() if text else ""
            if not text:
                continue

            # box 是4个角点: [[x1,y1], [x2,y2], [x3,y3], [x4,y4]]
            xs = [p[0] for p in box]
            ys = [p[1] for p in box]
            x_min = int(min(xs) / scale)
            y_min = int(min(ys) / scale)
            x_max = int(max(xs) / scale)
            y_max = int(max(ys) / scale)

            results.append({
                'text': text,
                'x': x_min,
                'y': y_min,
                'w': x_max - x_min,
                'h': y_max - y_min,
                'center': ((x_min + x_max) // 2, (y_min + y_max) // 2),
                'right': (x_max, (y_min + y_max) // 2),
                'score': float(score),
            })

        return results

    @staticmethod
    def find_text(results: List[dict], target: str) -> Optional[dict]:
        """在全部 OCR 结果中查找目标文字（模糊包含匹配）"""
        for r in results:
            if target in r['text'] or r['text'] in target:
                return r
        return None

    @staticmethod
    def find_text_in_region(
        results: List[dict],
        target: str,
        region: tuple,
        img_size: tuple
    ) -> Optional[dict]:
        """
        在指定区域内查找目标文字
        region: (x_start, y_start, x_end, y_end) 比率
        img_size: (width, height)
        """
        w, h = img_size
        x_start = int(w * region[0])
        y_start = int(h * region[1])
        x_end = int(w * region[2])
        y_end = int(h * region[3])

        for r in results:
            cx, cy = r['center']
            if x_start <= cx <= x_end and y_start <= cy <= y_end:
                if target in r['text'] or r['text'] in target:
                    return r
        return None


# ============================================================
#  窗口管理
# ============================================================
def find_windows(keyword: str) -> List[Tuple[int, str]]:
    """查找标题包含关键词的可见窗口"""
    found = []

    def callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if title and keyword.lower() in title.lower():
                found.append((hwnd, title))

    win32gui.EnumWindows(callback, None)
    return found


def list_all_windows() -> List[Tuple[int, str]]:
    """列出所有可见窗口"""
    windows = []

    def callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if title:
                windows.append((hwnd, title))

    win32gui.EnumWindows(callback, None)
    return windows


def get_window_rect(hwnd) -> Tuple[int, int, int, int]:
    """获取窗口位置和大小 (left, top, width, height)"""
    rect = win32gui.GetWindowRect(hwnd)
    left, top, right, bottom = rect
    return (left, top, right - left, bottom - top)


def capture_window_printwindow(hwnd) -> Tuple[Image.Image, Tuple[int, int, int, int]]:
    """
    使用 PrintWindow API 截取窗口内容
    优势：即使窗口被其他窗口遮挡也能截取到内容
    """
    left, top, width, height = get_window_rect(hwnd)

    # 获取窗口设备上下文
    hwnd_dc = win32gui.GetWindowDC(hwnd)
    mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
    save_dc = mfc_dc.CreateCompatibleDC()

    # 创建位图
    bitmap = win32ui.CreateBitmap()
    bitmap.CreateCompatibleBitmap(mfc_dc, width, height)
    save_dc.SelectObject(bitmap)

    # PrintWindow: PW_RENDERFULLCONTENT=2 可捕获 DirectX 渲染内容
    ctypes.windll.user32.PrintWindow(hwnd, save_dc.GetSafeHdc(), PW_RENDERFULLCONTENT)

    # 转换为 PIL Image
    bmpinfo = bitmap.GetInfo()
    bmpstr = bitmap.GetBitmapBits(True)
    image = Image.frombuffer(
        'RGB',
        (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
        bmpstr, 'raw', 'BGRX', 0, 1
    )

    # 清理资源
    win32gui.DeleteObject(bitmap.GetHandle())
    save_dc.DeleteDC()
    mfc_dc.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwnd_dc)

    return image, (left, top, width, height)


def capture_window_mss(hwnd) -> Tuple[Image.Image, Tuple[int, int, int, int]]:
    """使用 mss 截取屏幕区域（窗口必须在最前面可见）"""
    left, top, width, height = get_window_rect(hwnd)
    sct = mss.mss()
    try:
        monitor = {
            "top": top, "left": left,
            "width": width, "height": height,
        }
        shot = sct.grab(monitor)
        image = Image.frombytes("RGB", shot.size, shot.bgra, "raw", "BGRX")
    finally:
        sct.close()
    return image, (left, top, width, height)


def capture_window(hwnd) -> Tuple[Image.Image, Tuple[int, int, int, int]]:
    """
    截取窗口画面
    vivo投屏窗口的内容是视频流渲染的，PrintWindow 抓不到，必须用 mss 前台截图
    """
    # 先确保窗口在最前面
    try:
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        pass
    time.sleep(0.05)
    return capture_window_mss(hwnd)


# ============================================================
#  状态机
# ============================================================
class State:
    IDLE = "空闲"              # 等待点击"领时长"
    AD_PLAYING = "广告播放中"   # 广告播放中，等待完成或取消直播
    LIVE_ROOM = "直播间内"      # 已误入直播间，需要点X退出
    REWARD_POPUP = "奖励弹窗"   # 奖励弹窗，点击"领取奖励"


class AdWatcher:
    """汽水音乐广告自动观看器"""

    def __init__(self, config: Config):
        self.config = config
        self.ocr = OCREngine(config)
        self.state = State.IDLE
        self.running = False
        self.paused = False
        self.hwnd = None
        self.window_pos = (0, 0, 0, 0)
        self.state_start_time = time.time()
        self.last_screenshot = None
        # 标记当前广告是否已点过取消直播，避免重复点击
        self._live_cancelled_this_ad = False
        # GUI 日志回调函数（由 GUI 设置）
        self.log_callback = None
        self.stats = {
            'ads_watched': 0,
            'live_ads_cancelled': 0,
            'live_rooms_exited': 0,
            'errors': 0,
            'start_time': None,
        }

        # 创建调试目录
        if config.debug:
            Path(config.debug_dir).mkdir(exist_ok=True)

    # ----------------------------------------------------------
    #  窗口查找
    # ----------------------------------------------------------
    # 需要排除的窗口标题关键词（脚本自身窗口、命令行窗口等）
    EXCLUDE_TITLE_KEYWORDS = ("控制面板", "命令提示符", "Windows PowerShell", "cmd.exe", "Terminal")

    def find_target_window(self) -> bool:
        """查找投屏窗口，自动排除脚本自身的 GUI/终端窗口"""
        results = find_windows(self.config.window_title)
        if not results:
            logger.error(f"未找到包含 '{self.config.window_title}' 的窗口")
            logger.info("当前可见窗口列表：")
            for hwnd, title in list_all_windows():
                logger.info(f"  [{hwnd}] {title}")
            logger.info("提示：使用 --window <关键词> 指定窗口标题")
            return False

        # 过滤掉脚本自身的窗口（GUI 控制面板、cmd/PowerShell 终端等）
        filtered = []
        excluded = []
        for hwnd, title in results:
            title_lower = title.lower()
            is_self = any(kw.lower() in title_lower for kw in self.EXCLUDE_TITLE_KEYWORDS)
            if is_self:
                excluded.append((hwnd, title))
                logger.info(f"排除自身窗口: [{hwnd}] {title}")
            else:
                filtered.append((hwnd, title))

        if not filtered:
            logger.error(
                f"匹配到 {len(results)} 个窗口，但全部被排除为脚本自身窗口：\n"
                + "\n".join(f"  [{h}] {t}" for h, t in excluded)
                + "\n请在 GUI 配置里把'窗口标题关键词'改成更精确的投屏窗口标题（如'手机投屏'），避免匹配到本程序窗口。"
            )
            return False

        # 选择面积最大的窗口（投屏窗口通常较大）
        best_hwnd = None
        best_title = ""
        best_area = 0
        for hwnd, title in filtered:
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            area = (right - left) * (bottom - top)
            if area > best_area:
                best_area = area
                best_hwnd = hwnd
                best_title = title

        self.hwnd = best_hwnd
        logger.info(f"已锁定窗口: {best_title}")
        return True

    # ----------------------------------------------------------
    #  截图 + OCR
    # ----------------------------------------------------------
    def capture_and_ocr(self) -> Tuple[List[dict], Image.Image]:
        """截取窗口并执行 OCR"""
        screenshot, win_pos = capture_window(self.hwnd)
        self.window_pos = win_pos
        self.last_screenshot = screenshot
        results = self.ocr.detect(screenshot)
        return results, screenshot

    # ----------------------------------------------------------
    #  调试
    # ----------------------------------------------------------
    def save_debug(self, image: Image.Image, results: List[dict], tag: str = ""):
        """保存调试截图和 OCR 结果"""
        if not self.config.debug:
            return

        ts = datetime.now().strftime('%H%M%S_%f')[:-3]
        prefix = f"{ts}_{self.state}_{tag}"

        # 在截图上标注 OCR 结果
        annotated = image.copy()
        draw = ImageDraw.Draw(annotated)
        for r in results:
            x1, y1 = r['x'], r['y']
            x2, y2 = r['x'] + r['w'], r['y'] + r['h']
            draw.rectangle([x1, y1, x2, y2], outline="red", width=2)
            draw.text((x1, y1 - 12), f"{r['text'][:15]}({r['score']:.1f})", fill="red")

        annotated.save(Path(self.config.debug_dir) / f"{prefix}.png")

        with open(Path(self.config.debug_dir) / f"{prefix}.txt", 'w', encoding='utf-8') as f:
            f.write(f"状态: {self.state}\n")
            f.write(f"窗口位置: {self.window_pos}\n")
            f.write(f"图片尺寸: {image.size}\n")
            f.write(f"OCR结果 ({len(results)} 项):\n")
            for r in results:
                f.write(f"  '{r['text']}' score={r['score']:.2f} "
                        f"pos=({r['x']},{r['y']},{r['w']},{r['h']})\n")

    # ----------------------------------------------------------
    #  点击操作
    # ----------------------------------------------------------
    def click_at(self, x: int, y: int, desc: str = ""):
        """在屏幕绝对坐标点击，并在调试截图上标记点击位置"""
        if self.config.dry_run:
            logger.info(f"[试运行] 点击: {desc} ({x}, {y})")
            return

        logger.info(f"点击: {desc} ({x}, {y})")
        mouse_click(x, y)

        # 在调试截图上标记点击位置并保存
        if self.config.debug and self.last_screenshot is not None:
            left, top, _, _ = self.window_pos
            local_x = x - left
            local_y = y - top
            self._save_click_marker(local_x, local_y, desc)

        time.sleep(self.config.click_delay)

    def _save_click_marker(self, x: int, y: int, desc: str):
        """保存带点击标记的截图"""
        ts = datetime.now().strftime('%H%M%S_%f')[:-3]
        prefix = f"{ts}_CLICK_{self.state}"

        marked = self.last_screenshot.copy()
        draw = ImageDraw.Draw(marked)
        # 画十字准星
        size = 25
        draw.line([(x - size, y), (x + size, y)], fill="lime", width=3)
        draw.line([(x, y - size), (x, y + size)], fill="lime", width=3)
        # 画圆圈
        r = 15
        draw.ellipse([x - r, y - r, x + r, y + r], outline="lime", width=3)
        # 标注文字
        label = desc[:30] if desc else "click"
        draw.text((x + 20, y - 20), f"CLICK: {label} ({x},{y})", fill="lime")

        marked.save(Path(self.config.debug_dir) / f"{prefix}.png")

        # 同时写入点击日志
        with open(Path(self.config.debug_dir) / f"{prefix}.txt", 'w', encoding='utf-8') as f:
            f.write(f"点击时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"状态: {self.state}\n")
            f.write(f"描述: {desc}\n")
            f.write(f"窗口位置: {self.window_pos}\n")
            f.write(f"点击坐标(窗口内): ({x}, {y})\n")

    def click_ocr_center(self, result: dict, desc: str = ""):
        """点击 OCR 识别到的文字中心"""
        left, top, _, _ = self.window_pos
        cx, cy = result['center']
        self.click_at(left + cx, top + cy, desc)

    def click_ocr_right(self, result: dict, offset: int, desc: str = ""):
        """点击 OCR 识别到的文字右侧（偏移 offset 像素）"""
        left, top, _, _ = self.window_pos
        rx, ry = result['right']
        self.click_at(left + rx + offset, top + ry, desc)

    # ----------------------------------------------------------
    #  主循环
    # ----------------------------------------------------------
    def run(self):
        """主循环"""
        self.running = True
        self.stats['start_time'] = time.time()

        if not self.find_target_window():
            return

        # 将窗口置前
        try:
            win32gui.ShowWindow(self.hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(self.hwnd)
        except Exception:
            pass
        time.sleep(1)

        logger.info("=" * 55)
        logger.info("  汽水音乐自动看广告脚本 已启动")
        if HAS_KEYBOARD:
            logger.info("  按 ESC 键停止脚本")
        else:
            logger.info("  按 Ctrl+C 停止脚本")
        if self.config.dry_run:
            logger.info("  [试运行模式] 只检测不点击")
        logger.info("=" * 55)

        while self.running:
            try:
                if self.paused:
                    time.sleep(0.5)
                    continue
                self.tick()
            except Exception as e:
                logger.error(f"运行异常: {e}", exc_info=True)
                self.stats['errors'] += 1
                time.sleep(2)

        self.print_stats()

    def print_stats(self):
        """打印统计信息"""
        elapsed = time.time() - self.stats['start_time']
        logger.info("=" * 55)
        logger.info(f"  脚本已停止")
        logger.info(f"  观看广告: {self.stats['ads_watched']} 个")
        logger.info(f"  取消直播: {self.stats['live_ads_cancelled']} 次")
        logger.info(f"  退出直播间: {self.stats['live_rooms_exited']} 次")
        logger.info(f"  运行时长: {elapsed:.0f} 秒")
        logger.info(f"  错误次数: {self.stats['errors']}")
        logger.info("=" * 55)

    # ----------------------------------------------------------
    #  状态机 - 单次循环
    # ----------------------------------------------------------
    def _handle_continue_watching_popup(self, results: List[dict]) -> bool:
        """
        检测"坚持退出"免费解锁弹窗（如"第X日免费听"），命中则点击主按钮继续广告流程

        弹窗特征：带"坚持退出"灰色链接，主按钮可能有两种：
          - "继续观看"（如"再看08秒继续解锁"）
          - "领取奖励"（如"再看3个视频提前得"）
        用"坚持退出"作为唯一特征触发，避免与其他界面的"领取奖励"混淆。
        优先找"继续观看"，找不到再找"领取奖励"。
        """
        # 唯一特征"坚持退出"确认是本类解锁弹窗
        exit_link = OCREngine.find_text(results, "坚持退出")
        if not exit_link:
            return False

        # 优先查找"继续观看"主按钮，其次"领取奖励"按钮
        continue_btn = OCREngine.find_text(results, "继续观看")
        if not continue_btn:
            continue_btn = OCREngine.find_text(results, "领取奖励")
        if not continue_btn:
            logger.warning("检测到'坚持退出'解锁弹窗，但未识别到'继续观看'或'领取奖励'按钮，忽略")
            return False

        btn_text = continue_btn['text']
        logger.info(f"检测到'坚持退出'免费解锁弹窗，点击'{btn_text}'按钮...")
        self.click_ocr_center(continue_btn, f"解锁弹窗({btn_text})")
        return True

    def tick(self):
        """状态机单次执行"""

        # 检查广告数量上限
        if self.config.max_ads > 0 and self.stats['ads_watched'] >= self.config.max_ads:
            logger.info(f"已达到最大广告数 {self.config.max_ads}，停止")
            self.running = False
            return

        # 截图 + OCR
        results, screenshot = self.capture_and_ocr()
        img_size = screenshot.size

        # 保存调试信息
        self.save_debug(screenshot, results)

        # === 特别弹窗检测：识别"继续观看"免费解锁弹窗 ===
        # 界面特征"第18日免费听/再看06秒继续解锁"，有绿色"继续观看"主按钮和灰色"坚持退出"链接
        # 用"坚持退出"作为唯一特征触发（其他界面无此词），识别到就点"继续观看"恢复广告流程
        if self._handle_continue_watching_popup(results):
            self.state = State.AD_PLAYING
            self.state_start_time = time.time()
            self._live_cancelled_this_ad = False
            logger.info("点击'继续观看'后进入广告播放状态")
            time.sleep(self.config.ad_load_wait)
            return

        # 根据当前状态处理
        if self.state == State.IDLE:
            self._handle_idle(results, img_size)

        elif self.state == State.AD_PLAYING:
            self._handle_ad_playing(results, img_size)

        elif self.state == State.LIVE_ROOM:
            self._handle_live_room(results, img_size)

        elif self.state == State.REWARD_POPUP:
            self._handle_reward_popup(results, img_size)

    # ----------------------------------------------------------
    #  状态处理：空闲
    # ----------------------------------------------------------
    def _handle_idle(self, results: List[dict], img_size: tuple):
        """空闲状态：寻找并点击"领时长"/"继续领"按钮"""

        # 1. 寻找入口按钮（兼容"领时长"和"继续领"两种文字）
        target = OCREngine.find_text(results, "领时长")
        if not target:
            target = OCREngine.find_text(results, "继续领")
        if target:
            logger.info(f"找到入口按钮'{target['text']}'，点击开始观看广告...")
            self.click_ocr_center(target, f"入口按钮({target['text']})")
            self.state = State.AD_PLAYING
            self.state_start_time = time.time()
            self._live_cancelled_this_ad = False
            logger.info(f"等待广告加载({self.config.ad_load_wait}s)...")
            time.sleep(self.config.ad_load_wait)
            return

        # 2. 可能已在广告中（从异常中恢复）
        if OCREngine.find_text(results, "领取成功"):
            logger.info("检测到'领取成功'，广告已播放完毕，切换到广告播放状态")
            self.state = State.AD_PLAYING
            self._live_cancelled_this_ad = False
            return

        if OCREngine.find_text(results, "领取奖励"):
            logger.info("检测到'领取奖励'弹窗，切换到奖励弹窗状态")
            self.state = State.REWARD_POPUP
            return

        # 3. 检测是否在广告播放中（有倒计时文字）
        if (OCREngine.find_text(results, "可领取") or
            OCREngine.find_text(results, "秒后")):
            logger.info("检测到广告正在播放（有倒计时文字），切换到广告播放状态")
            self.state = State.AD_PLAYING
            self._live_cancelled_this_ad = False
            return

        # 4. 检测是否误入直播间
        if (OCREngine.find_text(results, "说点什么") or
            OCREngine.find_text(results, "本场点赞")):
            logger.warning("检测到已进入直播间，切换到直播间退出流程")
            self.state = State.LIVE_ROOM
            self.state_start_time = time.time()
            return

        logger.info("空闲中... 等待'领时长'按钮出现")
        time.sleep(self.config.scan_interval)

    # ----------------------------------------------------------
    #  状态处理：广告播放中
    # ----------------------------------------------------------
    def _handle_ad_playing(self, results: List[dict], img_size: tuple):
        """广告播放状态：检测直播取消 + 领取成功"""

        # 超时检查
        elapsed = time.time() - self.state_start_time
        if elapsed > self.config.ad_timeout:
            logger.warning(f"广告超时({self.config.ad_timeout}s)，返回空闲状态")
            self.state = State.IDLE
            return

        # === 优先级0：检测是否已误入直播间 ===
        # 直播间特征：有"关注"、"说点什么"、观众数等元素，但没有"领取成功"
        in_live_room = (
            OCREngine.find_text(results, "说点什么") or
            OCREngine.find_text(results, "关注") and
            (OCREngine.find_text(results, "本场点赞") or
             OCREngine.find_text(results, "直播中"))
        )
        has_success = OCREngine.find_text(results, "领取成功")
        if in_live_room and not has_success:
            logger.warning(">>> 未及时点取消，已进入直播间！切换到直播间退出流程 <<<")
            self.state = State.LIVE_ROOM
            self.state_start_time = time.time()
            return

        # === 优先级1：直播广告取消（5秒倒计时，必须快速！） ===
        # 只有同时检测到"取消"按钮时才处理直播广告
        # 没有"取消"按钮说明要么不是直播广告，要么倒计时已过，不要误点
        if not self._live_cancelled_this_ad:
            cancel_result = OCREngine.find_text_in_region(
                results, "取消", self.config.region_center, img_size
            )
            if cancel_result:
                # 确认旁边有"进入直播间"相关文字，才是直播倒计时弹窗
                live_result = OCREngine.find_text_in_region(
                    results, "进入直播间", self.config.region_center, img_size
                )
                if live_result:
                    logger.warning(">>> 检测到直播广告倒计时（5秒内）！立即点击'取消' <<<")
                    # OCR 常把"4s后进入直播间取消"识别为一整条文字，
                    # 如果点中心会点到"进入直播间"上，必须点右边缘才是"取消"
                    cancel_text = cancel_result['text']
                    if "进入直播间" in cancel_text or "后进入" in cancel_text:
                        logger.info(f"OCR将倒计时和取消识别为整体: '{cancel_text}'，点击右边缘")
                        self.click_ocr_right(cancel_result, self.config.live_cancel_offset, "取消(直播广告-右边缘)")
                    else:
                        self.click_ocr_center(cancel_result, "取消(直播广告)")
                    self.stats['live_ads_cancelled'] += 1
                    # 标记本次广告已取消直播，避免重复点击
                    self._live_cancelled_this_ad = True
                    logger.info("已点击取消，继续等待广告倒计时结束...")
                    time.sleep(1.0)
                    return
                else:
                    # 有"取消"但没有"进入直播间"，可能是其他弹窗，跳过
                    logger.debug("检测到'取消'文字但无'进入直播间'，非直播倒计时，忽略")

        # === 优先级2：广告完成，检测"领取成功" ===
        # 扩大搜索范围：先在右上角找，找不到就全屏找
        success_result = OCREngine.find_text_in_region(
            results, "领取成功", self.config.region_top_right, img_size
        )
        if not success_result:
            success_result = OCREngine.find_text(results, "领取成功")
        if success_result:
            logger.info("广告观看完成！检测到'领取成功'，准备点击右侧X按钮关闭...")
            # 点击"领取成功"右侧的X按钮
            self.click_ocr_right(
                success_result,
                self.config.x_button_offset,
                "X按钮(关闭广告)"
            )
            self.state = State.REWARD_POPUP
            self.state_start_time = time.time()
            time.sleep(1.5)
            return

        # === 优先级3：检测倒计时，输出等待日志 ===
        countdown = OCREngine.find_text(results, "秒后可领取")
        if countdown:
            logger.info(f"广告播放中... 倒计时: {countdown['text']}，等待结束")
        elif not self._live_cancelled_this_ad:
            logger.info("广告播放中... 正在检测直播倒计时和'领取成功'")
        else:
            logger.info("广告播放中... 已取消直播，等待右上角倒计时结束")

        # 广告播放中，不做健全性检查（投屏截图OCR识别率低）
        # 只靠超时和"领取成功"来驱动状态流转
        time.sleep(self.config.scan_interval)

    # ----------------------------------------------------------
    #  状态处理：直播间内（误入后退出）
    # ----------------------------------------------------------
    def _handle_live_room(self, results: List[dict], img_size: tuple):
        """直播间状态：点击右上角X退出直播间"""

        # 超时检查
        elapsed = time.time() - self.state_start_time
        if elapsed > self.config.live_room_timeout:
            logger.warning(f"直播间退出超时({self.config.live_room_timeout}s)，返回空闲状态")
            self.state = State.IDLE
            return

        # 检测是否还在直播间
        still_in_live = (
            OCREngine.find_text(results, "说点什么") or
            OCREngine.find_text(results, "本场点赞")
        )

        if not still_in_live:
            # 已退出直播间，可能弹出了推荐弹窗或回到了广告
            logger.info("已退出直播间，检查当前画面...")

            # 优先检查"退出不看了"推荐弹窗（退出直播间后常弹出"为你精选以下直播"）
            exit_btn = OCREngine.find_text(results, "退出不看了")
            if not exit_btn:
                exit_btn = OCREngine.find_text(results, "不看了")
            if exit_btn:
                logger.info("检测到'退出不看了'推荐弹窗，点击退出...")
                self.click_ocr_center(exit_btn, "退出不看了")
                time.sleep(1.5)
                return  # 留在 LIVE_ROOM 状态，下一轮重新检测

            # 检查是否有广告完成标志（兼容项目1"领取成功"和项目2"继续观看"）
            if OCREngine.find_text(results, "领取成功") or OCREngine.find_text(results, "继续观看"):
                logger.info("退出直播间后检测到广告完成，切换到广告播放状态")
                self.state = State.AD_PLAYING
                self.state_start_time = time.time()
                return
            # 否则回到空闲
            logger.info("退出直播间，回到空闲状态")
            self.state = State.IDLE
            return

        # 还在直播间，尝试点击右上角X退出
        logger.info("正在直播间内，尝试点击右上角X退出...")

        left, top, _, _ = self.window_pos
        clicked = False

        # 方法1：通过观众数定位 —— X按钮在观众数右侧约50像素
        # 观众数通常是4位数字，如"2350"
        viewer_result = None
        for r in results:
            text = r['text']
            # 找纯数字、3-5位、在顶部区域
            if text.isdigit() and 3 <= len(text) <= 5:
                cx, cy = r['center']
                if cy < img_size[1] * 0.15:  # 在顶部15%区域
                    viewer_result = r
                    break

        if viewer_result:
            rx, ry = viewer_result['right']
            click_x = left + rx + self.config.live_x_viewer_offset
            click_y = top + ry
            logger.info(f"通过观众数'{viewer_result['text']}'定位X按钮，点击({click_x},{click_y})")
            self.click_at(click_x, click_y, "直播间X(观众数右侧)")
            clicked = True

        # 方法2：通过"关注"按钮定位 —— X在关注右侧较远处
        if not clicked:
            follow_btn = OCREngine.find_text(results, "关注")
            if follow_btn:
                rx, ry = follow_btn['right']
                click_x = left + rx + self.config.live_x_follow_offset
                click_y = top + ry
                logger.info(f"通过'关注'按钮定位X按钮，点击({click_x},{click_y})")
                self.click_at(click_x, click_y, "直播间X(关注右侧)")
                clicked = True

        # 方法3：固定位置兜底 —— 根据实际截图，X在窗口宽度的~90%，高度的~10%
        if not clicked:
            w, h = img_size
            click_x = left + int(w * self.config.live_x_fixed_x_ratio)
            click_y = top + int(h * self.config.live_x_fixed_y_ratio)
            logger.info(f"使用固定位置定位X按钮，点击({click_x},{click_y})")
            self.click_at(click_x, click_y, "直播间X(固定位置)")
            clicked = True

        self.stats['live_rooms_exited'] += 1
        time.sleep(2.0)  # 等待退出动画

    # ----------------------------------------------------------
    #  状态处理：奖励弹窗
    # ----------------------------------------------------------
    def _handle_reward_popup(self, results: List[dict], img_size: tuple):
        """奖励弹窗状态：寻找并点击"领取奖励"按钮"""

        target = OCREngine.find_text(results, "领取奖励")
        if target:
            logger.info("找到'领取奖励'按钮，点击领取，准备进入下一个广告...")
            self.click_ocr_center(target, "领取奖励")
            self.stats['ads_watched'] += 1
            logger.info(f">>> 已完成第 {self.stats['ads_watched']} 个广告，等待下一个广告加载 <<<")

            # 下一个广告会自动播放
            self.state = State.AD_PLAYING
            self.state_start_time = time.time()
            self._live_cancelled_this_ad = False
            time.sleep(self.config.ad_load_wait)
            return

        # 超时检查
        elapsed = time.time() - self.state_start_time
        if elapsed > self.config.reward_timeout:
            logger.warning(f"奖励弹窗超时({self.config.reward_timeout}s)，返回空闲状态")
            self.state = State.IDLE
            return

        logger.info(f"奖励弹窗中... 等待'领取奖励'按钮出现 (已等待{elapsed:.0f}s)")
        time.sleep(self.config.scan_interval)

    # ----------------------------------------------------------
    #  停止
    # ----------------------------------------------------------
    def stop(self):
        """停止脚本"""
        self.running = False
        self.paused = False
        logger.info("正在停止...")

    def pause(self):
        """暂停脚本"""
        self.paused = True
        logger.info("脚本已暂停")

    def resume(self):
        """恢复脚本"""
        self.paused = False
        logger.info("脚本已恢复运行")


# ============================================================
#  测试模式
# ============================================================
def run_test_mode(config: Config):
    """测试模式：截图 + OCR，不执行任何点击"""
    logger.info("=" * 55)
    logger.info("  测试模式：截图 + OCR 识别")
    logger.info("=" * 55)

    watcher = AdWatcher(config)
    if not watcher.find_target_window():
        return

    # 将窗口置前，确保截取到正确内容
    try:
        win32gui.ShowWindow(watcher.hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(watcher.hwnd)
    except Exception:
        pass
    time.sleep(1)

    # 截图 + OCR
    results, screenshot = watcher.capture_and_ocr()

    logger.info(f"窗口位置: {watcher.window_pos}")
    logger.info(f"图片尺寸: {screenshot.size}")
    logger.info(f"OCR 识别到 {len(results)} 项文字：")
    logger.info("-" * 55)

    for i, r in enumerate(results, 1):
        logger.info(
            f"  {i:2d}. '{r['text']}'  "
            f"置信度={r['score']:.2f}  "
            f"位置=({r['x']},{r['y']}) "
            f"尺寸={r['w']}x{r['h']}"
        )

    # 保存标注截图
    debug_dir = Path(config.debug_dir)
    debug_dir.mkdir(exist_ok=True)

    annotated = screenshot.copy()
    draw = ImageDraw.Draw(annotated)
    for r in results:
        x1, y1 = r['x'], r['y']
        x2, y2 = r['x'] + r['w'], r['y'] + r['h']
        draw.rectangle([x1, y1, x2, y2], outline="red", width=2)
        draw.text((x1, max(0, y1 - 12)), r['text'][:20], fill="red")

    test_path = debug_dir / "test_result.png"
    annotated.save(test_path)
    logger.info("-" * 55)
    logger.info(f"标注截图已保存: {test_path}")
    logger.info("请检查截图确认 OCR 是否正确识别了关键文字")


# ============================================================
#  窗口列表模式
# ============================================================
def run_list_windows_mode():
    """列出所有可见窗口"""
    logger.info("=" * 55)
    logger.info("  当前可见窗口列表")
    logger.info("=" * 55)

    windows = list_all_windows()
    for hwnd, title in windows:
        rect = win32gui.GetWindowRect(hwnd)
        w = rect[2] - rect[0]
        h = rect[3] - rect[1]
        logger.info(f"  [{hwnd:10d}] {w:5d}x{h:5d}  {title}")

    logger.info("-" * 55)
    logger.info(f"共 {len(windows)} 个窗口")
    logger.info("使用 --window <关键词> 指定要匹配的窗口标题")


# ============================================================
#  命令行参数解析
# ============================================================
def parse_args() -> Config:
    """解析命令行参数"""
    config = Config()

    args = sys.argv[1:]

    # 先解析所有配置参数
    if "--window" in args:
        idx = args.index("--window")
        if idx + 1 < len(args):
            config.window_title = args[idx + 1]

    if "--interval" in args:
        idx = args.index("--interval")
        if idx + 1 < len(args):
            config.scan_interval = float(args[idx + 1])

    if "--max-ads" in args:
        idx = args.index("--max-ads")
        if idx + 1 < len(args):
            config.max_ads = int(args[idx + 1])

    if "--debug" in args:
        config.debug = True

    if "--no-debug" in args:
        config.debug = False

    if "--dry-run" in args:
        config.dry_run = True

    # 再处理模式标志（在配置解析之后）
    if "--list-windows" in args:
        run_list_windows_mode()
        sys.exit(0)

    if "--test" in args:
        run_test_mode(config)
        sys.exit(0)

    return config


# ============================================================
#  主函数
# ============================================================
def main():
    config = parse_args()

    # 创建 watcher
    watcher = AdWatcher(config)

    # 设置 ESC 热键
    if HAS_KEYBOARD:
        keyboard.add_hotkey('esc', lambda: watcher.stop())
    else:
        logger.warning("未安装 keyboard 模块，使用 Ctrl+C 停止")

    # 运行
    try:
        watcher.run()
    except KeyboardInterrupt:
        watcher.stop()


if __name__ == '__main__':
    main()
