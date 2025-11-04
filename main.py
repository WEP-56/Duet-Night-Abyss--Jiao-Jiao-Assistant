import os
import sys
import time
import json
import threading
import random
import ctypes
from datetime import datetime

import cv2
import numpy as np
import win32con
import win32gui
import win32ui
import win32api

import tkinter as tk
from tkinter import ttk, messagebox
import importlib


# ------------------------------
# Background screenshot (from test.py idea)
# ------------------------------
class BackgroundScreenshot:
    def __init__(self, hwnd=None):
        self.hwnd = hwnd

    def set_hwnd(self, hwnd):
        self.hwnd = hwnd

    def capture_background(self):
        if not self.hwnd or not win32gui.IsWindow(self.hwnd):
            return None
        try:
            left, top, right, bottom = win32gui.GetWindowRect(self.hwnd)
            width = max(1, right - left)
            height = max(1, bottom - top)

            hwndDC = win32gui.GetWindowDC(self.hwnd)
            mfcDC = win32ui.CreateDCFromHandle(hwndDC)
            saveDC = mfcDC.CreateCompatibleDC()
            saveBitMap = win32ui.CreateBitmap()
            saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
            saveDC.SelectObject(saveBitMap)

            res = ctypes.windll.user32.PrintWindow(self.hwnd, saveDC.GetSafeHdc(), 2)
            if res != 1:
                saveDC.BitBlt((0, 0), (width, height), mfcDC, (0, 0), win32con.SRCCOPY)

            bmpinfo = saveBitMap.GetInfo()
            bmpstr = saveBitMap.GetBitmapBits(True)
            img = np.frombuffer(bmpstr, dtype=np.uint8)
            img.shape = (bmpinfo['bmHeight'], bmpinfo['bmWidth'], 4)
            img = img[:, :, :3]

            win32gui.DeleteObject(saveBitMap.GetHandle())
            saveDC.DeleteDC()
            mfcDC.DeleteDC()
            win32gui.ReleaseDC(self.hwnd, hwndDC)

            return img
        except Exception:
            try:
                win32gui.DeleteObject(saveBitMap.GetHandle())
                saveDC.DeleteDC()
                mfcDC.DeleteDC()
                win32gui.ReleaseDC(self.hwnd, hwndDC)
            except Exception:
                pass
            return None


# ------------------------------
# SendMessage input helpers (from test2.py idea)
# ------------------------------
def _pack_lparam(x, y):
    return (y << 16) | (x & 0xFFFF)


def get_client_center(hwnd):
    left, top, right, bottom = win32gui.GetClientRect(hwnd)
    width = max(0, right - left)
    height = max(0, bottom - top)
    return width // 2, height // 2


def child_from_client_point(hwnd_parent, x, y):
    try:
        child = win32gui.ChildWindowFromPoint(hwnd_parent, (x, y))
        if child and win32gui.IsWindow(child):
            return child
    except Exception:
        pass
    return hwnd_parent


def map_point_parent_to_child(hwnd_parent, hwnd_child, x, y):
    if hwnd_parent == hwnd_child:
        return x, y
    try:
        sx, sy = win32gui.ClientToScreen(hwnd_parent, (x, y))
        cx, cy = win32gui.ScreenToClient(hwnd_child, (sx, sy))
        return cx, cy
    except Exception:
        return x, y


def ensure_restored(hwnd):
    try:
        placement = win32gui.GetWindowPlacement(hwnd)
        show_cmd = placement[1]
        if show_cmd == win32con.SW_SHOWMINIMIZED:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            time.sleep(0.2)
    except Exception:
        pass


def send_mouse_move(hwnd, x, y, wparam=0):
    lparam = _pack_lparam(x, y)
    win32gui.SendMessage(hwnd, win32con.WM_MOUSEMOVE, wparam, lparam)


def send_left_click(hwnd, x, y):
    lparam = _pack_lparam(x, y)
    send_mouse_move(hwnd, x, y)
    win32gui.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
    win32gui.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)


def _vk_from_key_name(name: str) -> int:
    name = name.strip().lower()
    if len(name) == 1:
        return ord(name.upper())
    mapping = {
        'space': win32con.VK_SPACE,
        'shift': win32con.VK_SHIFT,
        'ctrl': win32con.VK_CONTROL,
        'control': win32con.VK_CONTROL,
        'alt': win32con.VK_MENU,
        'tab': win32con.VK_TAB,
        'esc': win32con.VK_ESCAPE,
        'escape': win32con.VK_ESCAPE,
    }
    if name in mapping:
        return mapping[name]
    if name.isdigit():
        return ord(name)
    raise ValueError(f"Unsupported key: {name}")


def _make_key_lparam(vk: int, is_keyup: bool) -> int:
    scan = win32api.MapVirtualKey(vk, 0) & 0xFF
    lparam = (1) | (scan << 16)
    if is_keyup:
        lparam |= (1 << 30) | (1 << 31)
    return lparam


def send_key_down(hwnd, key_name: str):
    vk = _vk_from_key_name(key_name)
    lparam = _make_key_lparam(vk, is_keyup=False)
    win32gui.SendMessage(hwnd, win32con.WM_KEYDOWN, vk, lparam)


def send_key_up(hwnd, key_name: str):
    vk = _vk_from_key_name(key_name)
    lparam = _make_key_lparam(vk, is_keyup=True)
    win32gui.SendMessage(hwnd, win32con.WM_KEYUP, vk, lparam)


def send_key_press(hwnd, key_name: str, hold: float = 0.0):
    send_key_down(hwnd, key_name)
    if hold > 0:
        time.sleep(hold)
    send_key_up(hwnd, key_name)


# ------------------------------
# Template matching helpers
# ------------------------------
def match_template(bgr_img, template_path, threshold=0.85):
    if bgr_img is None:
        return None
    if not os.path.isfile(template_path):
        return None
    try:
        tpl = cv2.imread(template_path, cv2.IMREAD_COLOR)
        if tpl is None:
            return None
        res = cv2.matchTemplate(bgr_img, tpl, cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
        if max_val >= threshold:
            h, w = tpl.shape[:2]
            x, y = max_loc
            center = (x + w // 2, y + h // 2)
            return {
                'score': float(max_val),
                'rect': (x, y, w, h),
                'center': center,
            }
        return None
    except Exception:
        return None


def enumerate_windows():
    wins = []
    def _enum_cb(h, _):
        if win32gui.IsWindowVisible(h) and win32gui.GetWindowText(h):
            wins.append((h, win32gui.GetClassName(h), win32gui.GetWindowText(h)))
    win32gui.EnumWindows(_enum_cb, None)
    return wins


# ------------------------------
# Edge-based map matching helpers
# ------------------------------
def _edges1ch(img_bgr):
    try:
        gray = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)
    except Exception:
        gray = img_bgr
    edges = cv2.Canny(gray, 50, 150)
    kernel = np.ones((3, 3), np.uint8)
    edges = cv2.dilate(edges, kernel, iterations=1)
    return edges


def _load_template_edge_and_mask(path):
    # Read with alpha if present
    tpl = cv2.imread(path, cv2.IMREAD_UNCHANGED)
    if tpl is None:
        return None, None
    if tpl.ndim == 3 and tpl.shape[2] == 4:
        bgr = tpl[:, :, :3]
        alpha = tpl[:, :, 3]
    else:
        bgr = tpl if tpl.ndim == 3 else cv2.cvtColor(tpl, cv2.COLOR_GRAY2BGR)
        alpha = None
    e = _edges1ch(bgr)
    edge_mask = (e > 0).astype(np.uint8) * 255
    if alpha is not None:
        # use alpha>0 as content mask; dilate a bit to improve tolerance on very细的线条
        a_mask = (alpha > 0).astype(np.uint8) * 255
        kernel = np.ones((3, 3), np.uint8)
        a_mask = cv2.dilate(a_mask, kernel, iterations=1)
        mask = cv2.bitwise_and(edge_mask, a_mask)
        # if intersection becomes (almost) empty due to thin edges, fallback to alpha-only mask
        if np.count_nonzero(mask) < 50:
            mask = a_mask
    else:
        mask = edge_mask
    return e, mask


# ------------------------------
# Player for JSON actions
# ------------------------------
def play_actions(hwnd, steps, logfn, stop_event=None):
    ensure_restored(hwnd)
    cx, cy = get_client_center(hwnd)
    target = child_from_client_point(hwnd, cx, cy)
    # Activate with a left click first if the first actionable step is a key, to ensure background input
    try:
        first = next((s for s in steps if s.get('type') in ('key', 'mouse')), None)
        if first and first.get('type') == 'key':
            tx, ty = map_point_parent_to_child(hwnd, target, cx, cy)
            send_left_click(target, tx, ty)
            if stop_event and stop_event.is_set():
                return
            time.sleep(0.05)
    except Exception:
        pass
    for i, st in enumerate(steps):
        if stop_event and stop_event.is_set():
            return
        t = st.get('type')
        delay = float(st.get('delay', 0))
        try:
            if t == 'key':
                key = st['key']
                hold = float(st.get('hold', 0))
                logfn(f"动作{i+1}: key {key} hold={hold}s delay={delay}s")
                # deliver to both child and top-level for compatibility
                send_key_press(target, key, hold)
                send_key_press(hwnd, key, 0)
            elif t == 'mouse':
                btn = st.get('button', 'left').lower()
                hold = float(st.get('hold', 0))
                logfn(f"动作{i+1}: mouse {btn} hold={hold}s delay={delay}s")
                if btn == 'left':
                    tx, ty = map_point_parent_to_child(hwnd, target, cx, cy)
                    send_left_click(target, tx, ty)
                    if hold > 0:
                        # interruptible hold
                        waited = 0.0
                        step = 0.05
                        while waited < hold and not (stop_event and stop_event.is_set()):
                            time.sleep(step)
                            waited += step
                elif btn == 'right':
                    # emulate hold by down/up
                    tx, ty = map_point_parent_to_child(hwnd, target, cx, cy)
                    lparam = _pack_lparam(tx, ty)
                    send_mouse_move(target, tx, ty)
                    win32gui.SendMessage(target, win32con.WM_RBUTTONDOWN, win32con.MK_RBUTTON, lparam)
                    if hold > 0:
                        waited = 0.0
                        step = 0.05
                        while waited < hold and not (stop_event and stop_event.is_set()):
                            time.sleep(step)
                            waited += step
                    win32gui.SendMessage(target, win32con.WM_RBUTTONUP, 0, lparam)
            else:
                logfn(f"未知动作类型: {t}")
        except Exception as e:
            logfn(f"执行动作错误: {e}")
        if delay > 0:
            waited = 0.0
            step = 0.05
            while waited < delay and not (stop_event and stop_event.is_set()):
                time.sleep(step)
                waited += step


# ------------------------------
# Main App GUI
# ------------------------------
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("免费脚本，仅供学习-by WEP")
        self.running = False
        self.worker = None
        self.threshold = 0.85
        # Map recognition parameters (old-script style)
        self.feature_presence_thr = 0.78  # a feature is considered present if score >= this
        self.quick_accept_thr = 0.90      # if any feature score >= this, quick accept top candidate
        self.min_hits = 2                 # minimal number of present features to accept when not quick-accept
        self.retry_interval = 1.0
        self.timeout_seconds = 300.0
        self.post_click_wait = 1.5

        self.base_dir = self._compute_base_dir()
        self.control_dir = os.path.join(self.base_dir, 'control')
        # roots; actual working dirs depend on selected mode
        self.map_root = os.path.join(self.base_dir, 'map')
        self.json_root = os.path.join(self.base_dir, 'json')
        # defaults before a mode is chosen
        self.map_dir = self.map_root
        self.json_dir = self.json_root
        self.config_path = os.path.join(self.base_dir, 'config.json')
        self.log_file_path = os.path.join(self.base_dir, 'app.log')

        self.auto_keyword = tk.StringVar(value='二重螺旋')
        self.selected_hwnd = None
        self.capturer = BackgroundScreenshot()
        self._tpl_edge_cache = {}
        self.stop_event = threading.Event()

        # Settings (GUI-configurable)
        self.fail_fallback_random = False   # 识图失败时随机脚本（默认关闭）
        self.post_likai_delay = 1.3         # 进入地图后延迟秒
        self.max_loops = 0                  # 循环次数（0=不限）
        self.auto_stop_seconds = 0          # 定时关闭（秒，0=禁用）
        self.started_at = None
        self.loops_done = 0
        self.auto_stop_timer = None

        # Try load persisted config
        self._load_config()

        # expose action player to logic modules
        self.play_actions = play_actions

        # mode display mapping
        self.mode_name_map = {
            '55mod': '夜航55',
            'juesemihan': '驱离角色密函',
            'wuqimihan': '驱离武器密函',
        }
        self.mode_display_to_key = {}
        self.mode_key_to_display = {}

        self._build_ui()
        self._bind_hotkeys()
        self.refresh_windows()

    # UI
    def _build_ui(self):
        wrapper = ttk.Frame(self.root, padding=10)
        wrapper.pack(fill='both', expand=True)

        title = ttk.Label(wrapper, text='牛马皎皎后台挂机1.0', font=('Microsoft YaHei', 16, 'bold'))
        title.pack(anchor='center', pady=(0, 4))

        tip = ttk.Label(
            wrapper,
            text='进入对应副本列表，点击前往进入确认选择界面再开启脚本\n若脚本完全无法运行请尝试刷新窗口，或记录终端信息至邮箱1484413790@qq.com\n 此脚本完全免费，如果您是购买获得请申请退款！！',
            font=('Microsoft YaHei', 9)
        )
        tip.pack(anchor='center', pady=(0, 10))

        # Controls row (imitate download layout visually simple)
        controls_frame = ttk.Frame(wrapper)
        controls_frame.pack(fill='x', pady=(0, 8))

        self.btn_start = ttk.Button(controls_frame, text='开始 (F10)', command=self.on_start)
        self.btn_start.pack(side='left', padx=5)

        self.btn_stop = ttk.Button(controls_frame, text='停止 (F12)', command=self.on_stop)
        self.btn_stop.pack(side='left', padx=5)

        self.btn_refresh = ttk.Button(controls_frame, text='刷新窗口 (F9)', command=self.refresh_windows)
        self.btn_refresh.pack(side='left', padx=5)

        self.btn_settings = ttk.Button(controls_frame, text='设置', command=self._open_settings)
        self.btn_settings.pack(side='left', padx=5)

        # Mode selection
        row_mode = ttk.Frame(wrapper)
        row_mode.pack(fill='x', pady=4)
        ttk.Label(row_mode, text='模式:').pack(side='left')
        self.mode_var = tk.StringVar(value='夜航55')
        self.combo_mode = ttk.Combobox(row_mode, state='readonly', width=30, textvariable=self.mode_var)
        self.combo_mode.pack(side='left', padx=6)
        ttk.Button(row_mode, text='刷新模式', command=self._refresh_modes).pack(side='left', padx=4)

        # 武器密函选择
        row_mihan = ttk.Frame(wrapper)
        row_mihan.pack(fill='x', pady=4)
        ttk.Label(row_mihan, text='武器密函选择:').pack(side='left')
        self.wuqi_mihan_var = tk.StringVar(value='')
        self.combo_wuqi_mihan = ttk.Combobox(row_mihan, state='readonly', width=30, textvariable=self.wuqi_mihan_var)
        self.combo_wuqi_mihan.pack(side='left', padx=6)
        ttk.Button(row_mihan, text='刷新密函', command=self._refresh_wuqi_mihan).pack(side='left', padx=4)

        # Window selection
        row2 = ttk.Frame(wrapper)
        row2.pack(fill='x', pady=4)

        ttk.Label(row2, text='自动关键字:').pack(side='left')
        ent = ttk.Entry(row2, textvariable=self.auto_keyword, width=20)
        ent.pack(side='left', padx=4)

        ttk.Label(row2, text='窗口:').pack(side='left')
        self.combo = ttk.Combobox(row2, state='readonly', width=50)
        self.combo.pack(side='left', padx=4, fill='x', expand=True)
        self.combo.bind('<<ComboboxSelected>>', self.on_combo_select)

        # Log box
        self.log = tk.Text(wrapper, height=18)
        self.log.pack(fill='both', expand=True)

        # Status bar
        status_frame = ttk.Frame(wrapper)
        status_frame.pack(fill='x')
        self.status_var = tk.StringVar(value='状态: 就绪')
        ttk.Label(status_frame, textvariable=self.status_var, anchor='w').pack(side='left', fill='x', expand=True)

        # init modes
        self._refresh_modes()
        self._refresh_wuqi_mihan()

    def _bind_hotkeys(self):
        # Local Tk bindings (also support global with keyboard if installed)
        try:
            import keyboard  # type: ignore
            keyboard.add_hotkey('f10', self.on_start)
            keyboard.add_hotkey('f12', self.on_stop)
            keyboard.add_hotkey('f9', self.refresh_windows)
        except Exception:
            pass

    def _refresh_modes(self):
        try:
            logic_dir = os.path.join(self.base_dir, 'logic')
            modes = []
            if os.path.isdir(logic_dir):
                for fn in os.listdir(logic_dir):
                    if not fn.lower().endswith('.py'):
                        continue
                    name = os.path.splitext(fn)[0]
                    if name in ('__init__', '') or name.startswith('_'):
                        continue
                    modes.append(name)
            modes.sort()
            # build display mapping
            labels = []
            self.mode_display_to_key = {}
            self.mode_key_to_display = {}
            for k in modes:
                lbl = f"{self.mode_name_map.get(k, k)}"
                labels.append(lbl)
                self.mode_display_to_key[lbl] = k
                self.mode_key_to_display[k] = lbl
            self.combo_mode['values'] = labels
            # default to 55mod label if present, else first
            if labels:
                prefer_key = '55mod'
                prefer_label = self.mode_key_to_display.get(prefer_key)
                current = self.mode_var.get()
                if prefer_label and prefer_label in labels:
                    self.mode_var.set(prefer_label)
                elif current not in labels:
                    self.mode_var.set(labels[0])
            else:
                self.mode_var.set('')
            self._log(f"已刷新模式列表: {', '.join(labels) if labels else '无'}")
        except Exception as e:
            self._log(f"刷新模式失败: {e}")

    def _refresh_wuqi_mihan(self):
        try:
            folder = os.path.join(self.control_dir, '武器密函png')
            names = []
            if os.path.isdir(folder):
                for fn in os.listdir(folder):
                    if fn.lower().endswith('.png'):
                        names.append(os.path.splitext(fn)[0])
            names.sort()
            self.combo_wuqi_mihan['values'] = names
            if names:
                cur = self.wuqi_mihan_var.get()
                if cur not in names:
                    self.wuqi_mihan_var.set(names[0])
            else:
                self.wuqi_mihan_var.set('')
            self._log(f"已刷新武器密函列表: {', '.join(names) if names else '无'}")
        except Exception as e:
            self._log(f"刷新武器密函失败: {e}")

    def get_selected_wuqi_mihan_path(self):
        name = (self.wuqi_mihan_var.get() or '').strip()
        if not name:
            return None
        return os.path.join(self.control_dir, '武器密函png', f"{name}.png")

    # Helpers for logic modules
    def detect_template_abs(self, template_abs_path, threshold=None):
        thr = self.threshold if threshold is None else float(threshold)
        img = self.capturer.capture_background()
        if img is None:
            return None
        if not os.path.isfile(template_abs_path):
            self._log(f"模板不存在: {template_abs_path}")
            return None
        try:
            res = match_template(img, template_abs_path, thr)
            if res is None:
                # simple max-score probe for diagnostics
                tpl = cv2.imread(template_abs_path, cv2.IMREAD_COLOR)
                if tpl is not None:
                    r = cv2.matchTemplate(img, tpl, cv2.TM_CCOEFF_NORMED)
                    _min, _max, _minl, _maxl = cv2.minMaxLoc(r)
                    self._log(f"匹配阈值未达标: path={os.path.basename(template_abs_path)} max={_max:.2f} thr={thr}")
            return res
        except Exception as e:
            self._log(f"检测异常({os.path.basename(template_abs_path)}): {e}")
            return None

    def detect_template_abs_scales(self, template_abs_path, scales=None, threshold=None):
        thr = self.threshold if threshold is None else float(threshold)
        img = self.capturer.capture_background()
        if img is None:
            return None
        if not os.path.isfile(template_abs_path):
            self._log(f"模板不存在: {template_abs_path}")
            return None
        try:
            tpl = cv2.imread(template_abs_path, cv2.IMREAD_COLOR)
            if tpl is None:
                return None
            (ih, iw) = img.shape[:2]
            best = None
            for s in (scales or [1.0, 0.95, 0.9, 1.05, 1.1]):
                th = max(1, int(tpl.shape[0] * s))
                tw = max(1, int(tpl.shape[1] * s))
                if th >= ih or tw >= iw:
                    continue
                rs = cv2.resize(tpl, (tw, th), interpolation=cv2.INTER_LINEAR)
                r = cv2.matchTemplate(img, rs, cv2.TM_CCOEFF_NORMED)
                _min, _max, _minl, _maxl = cv2.minMaxLoc(r)
                if best is None or _max > best[0]:
                    best = (_max, _maxl, rs.shape[1], rs.shape[0])
            if best is None:
                return None
            max_val, max_loc, w, h = best
            if max_val < thr:
                self._log(f"多尺度匹配未达阈值: path={os.path.basename(template_abs_path)} max={max_val:.2f} thr={thr}")
                return None
            x, y = max_loc
            center = (x + w // 2, y + h // 2)
            return {'score': float(max_val), 'rect': (x, y, w, h), 'center': center}
        except Exception as e:
            self._log(f"多尺度检测异常({os.path.basename(template_abs_path)}): {e}")
            return None

    def center_to_client_and_target(self, center_xy):
        try:
            cx, cy = center_xy
            win_left, win_top, _, _ = win32gui.GetWindowRect(self.selected_hwnd)
            sx, sy = win_left + cx, win_top + cy
            tx, ty = win32gui.ScreenToClient(self.selected_hwnd, (sx, sy))
            target = child_from_client_point(self.selected_hwnd, tx, ty)
            return target, tx, ty
        except Exception:
            return self.selected_hwnd, *get_client_center(self.selected_hwnd)

    def click_match_abs(self, template_abs_path, name_alias='', threshold=None, scales=None): 
        if scales:
            m = self.detect_template_abs_scales(template_abs_path, scales=scales, threshold=threshold)
        else:
            m = self.detect_template_abs(template_abs_path, threshold=threshold)
        if not m:
            return False
        cx, cy = m['center']
        target, tx, ty = self.center_to_client_and_target((cx, cy))
        self._log(f"点击 {name_alias or os.path.basename(template_abs_path)} @ ({tx},{ty})")
        lp = _pack_lparam(tx, ty)
        send_mouse_move(target, tx, ty)
        win32gui.SendMessage(target, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lp)
        win32gui.SendMessage(target, win32con.WM_LBUTTONUP, 0, lp)
        waited = 0.0
        step = 0.05
        while waited < self.post_click_wait and self.running and not self.stop_event.is_set():
            time.sleep(step)
            waited += step
        return True

    def send_mouse_wheel(self, delta=120, count=1, client_pos=None):
        try:
            # Per MSDN: lParam holds screen coordinates; message is sent to focus window.
            # For background usage, send to the top-level hwnd with screen coords of client center.
            if client_pos is None:
                cx, cy = get_client_center(self.selected_hwnd)
            else:
                cx, cy = client_pos
            sx, sy = win32gui.ClientToScreen(self.selected_hwnd, (cx, cy))
            lparam = _pack_lparam(sx, sy)
            wparam = (int(delta) & 0xFFFF) << 16
            target = child_from_client_point(self.selected_hwnd, cx, cy)
            for _ in range(max(1, int(count))):
                win32gui.SendMessage(target, win32con.WM_MOUSEWHEEL, wparam, lparam)
                self._log(f"发送滚轮: delta={delta} -> target=0x{target:08X} screen({sx},{sy}) client({cx},{cy})")
                time.sleep(0.06)
            return True
        except Exception:
            return False

    def _append_log_line_ui(self, line):
        try:
            self.log.insert('end', line)
            self.log.see('end')
        except Exception:
            pass

    def _log(self, msg):
        ts = datetime.now().strftime('%H:%M:%S')
        line = f"[{ts}] {msg}\n"
        try:
            self.root.after(0, self._append_log_line_ui, line)
        except Exception:
            pass
        try:
            print(line, end='')
        except Exception:
            pass
        try:
            with open(self.log_file_path, 'a', encoding='utf-8') as _f:
                _f.write(line)
        except Exception:
            pass

    def _compute_base_dir(self):
        try:
            if getattr(sys, 'frozen', False):
                return os.path.dirname(sys.executable)
        except Exception:
            pass
        try:
            return os.path.dirname(os.path.abspath(__file__))
        except Exception:
            return os.getcwd()

    # ------------------ Config persist ------------------
    def _load_config(self):
        try:
            if os.path.isfile(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    cfg = json.load(f)
                self.fail_fallback_random = bool(cfg.get('fail_fallback_random', self.fail_fallback_random))
                self.post_likai_delay = float(cfg.get('post_likai_delay', self.post_likai_delay))
                self.max_loops = int(cfg.get('max_loops', self.max_loops))
                self.auto_stop_seconds = int(cfg.get('auto_stop_seconds', self.auto_stop_seconds))
                # Clamp after load
                self._clamp_settings()
                self._log('已加载本地配置文件。')
        except Exception as e:
            self._log(f"加载配置失败: {e}")

    def _save_config(self):
        try:
            cfg = {
                'fail_fallback_random': self.fail_fallback_random,
                'post_likai_delay': self.post_likai_delay,
                'max_loops': self.max_loops,
                'auto_stop_seconds': self.auto_stop_seconds,
            }
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
            self._log('设置已保存到本地配置文件。')
        except Exception as e:
            self._log(f"保存配置失败: {e}")

    # Validation and clamping
    def _clamp_settings(self):
        try:
            # Bounds: delay [0..10], loops [0..10000], auto_stop [0..86400]
            self.post_likai_delay = max(0.0, min(10.0, float(self.post_likai_delay)))
            self.max_loops = max(0, min(10000, int(self.max_loops)))
            self.auto_stop_seconds = max(0, min(86400, int(self.auto_stop_seconds)))
        except Exception:
            pass

    # Window management
    def refresh_windows(self):
        wins = enumerate_windows()
        keyword = self.auto_keyword.get().strip()
        auto = None
        items = []
        for h, cls, title in wins:
            item = f"0x{h:08X} | {title} | {cls}"
            items.append((item, h, title))
            if keyword and (keyword in title):
                auto = h
        self.combo['values'] = [it[0] for it in items]
        if items:
            self.combo.current(0)
        if auto:
            self._set_target(auto)
            self._log(f"自动锁定窗口: 0x{auto:08X} {win32gui.GetWindowText(auto)}")
        else:
            self._log('未自动锁定窗口，请在下拉列表中选择。')

    def _open_settings(self):
        if hasattr(self, 'settings_window') and self.settings_window and tk.Toplevel.winfo_exists(self.settings_window):
            try:
                self.settings_window.focus_set()
                return
            except Exception:
                pass
        win = tk.Toplevel(self.root)
        self.settings_window = win
        win.title('设置')
        win.geometry('380x260')
        frm = ttk.Frame(win, padding=10)
        frm.pack(fill='both', expand=True)

        # Vars bound to settings
        self.var_fail_random = tk.BooleanVar(value=self.fail_fallback_random)
        self.var_post_delay = tk.DoubleVar(value=self.post_likai_delay)
        self.var_max_loops = tk.IntVar(value=self.max_loops)
        self.var_auto_stop = tk.IntVar(value=self.auto_stop_seconds)

        # Row: checkbox fail random
        chk = ttk.Checkbutton(frm, text='识图失败时随机选择脚本继续', variable=self.var_fail_random)
        chk.pack(anchor='w', pady=(0,6))

        # Row: post likai delay
        row1 = ttk.Frame(frm)
        row1.pack(fill='x', pady=4)
        ttk.Label(row1, text='进入地图后延迟').pack(side='left')
        ent_delay = ttk.Entry(row1, textvariable=self.var_post_delay, width=8)
        ent_delay.pack(side='left', padx=6)
        ttk.Label(row1, text='秒再识别地图').pack(side='left')

        # Row: loop count
        row2 = ttk.Frame(frm)
        row2.pack(fill='x', pady=4)
        ttk.Label(row2, text='循环次数').pack(side='left')
        ent_loops = ttk.Entry(row2, textvariable=self.var_max_loops, width=8)
        ent_loops.pack(side='left', padx=6)
        ttk.Label(row2, text='(0为不限)').pack(side='left')

        # Row: auto stop seconds
        row3 = ttk.Frame(frm)
        row3.pack(fill='x', pady=4)
        ttk.Label(row3, text='定时关闭').pack(side='left')
        ent_auto = ttk.Entry(row3, textvariable=self.var_auto_stop, width=8)
        ent_auto.pack(side='left', padx=6)
        ttk.Label(row3, text='秒 (0为关闭)').pack(side='left')

        # Buttons
        btns = ttk.Frame(frm)
        btns.pack(fill='x', pady=(10,0))
        def _apply_and_close():
            try:
                self.fail_fallback_random = bool(self.var_fail_random.get())
                # Parse and clamp with bounds
                self.post_likai_delay = float(self.var_post_delay.get())
                self.max_loops = int(self.var_max_loops.get())
                self.auto_stop_seconds = int(self.var_auto_stop.get())
                self._clamp_settings()
                self._log(f"已应用设置: 随机脚本={self.fail_fallback_random}, 延迟={self.post_likai_delay}s, 循环次数={self.max_loops}, 定时关闭={self.auto_stop_seconds}s")
                self._save_config()
            except Exception as e:
                self._log(f"应用设置失败: {e}")
            try:
                win.destroy()
            except Exception:
                pass
        ttk.Button(btns, text='保存', command=_apply_and_close).pack(side='right')

    def on_combo_select(self, _evt=None):
        idx = self.combo.current()
        try:
            text = self.combo['values'][idx]
        except Exception:
            return
        # Parse hwnd at start
        try:
            hwnd_str = text.split('|', 1)[0].strip()
            hwnd = int(hwnd_str, 16)
            self._set_target(hwnd)
        except Exception:
            pass

    def _set_target(self, hwnd):
        self.selected_hwnd = hwnd
        self.capturer.set_hwnd(hwnd)

    # Buttons
    def on_start(self):
        if self.running:
            return
        if not self.selected_hwnd or not win32gui.IsWindow(self.selected_hwnd):
            messagebox.showwarning('提示', '请先选择有效窗口')
            return
        # quick env check
        try:
            self._log(f"[自检] base_dir={self.base_dir}")
            self._log(f"[自检] control={self.control_dir} 存在={os.path.isdir(self.control_dir)}")
            for fn in ['xuanzemihan.png','bushiyong.png','querenxuanze.png','likai.png','zaicijinixng.png']:
                p = os.path.join(self.control_dir, fn)
                self._log(f"[自检] 模板 {fn}: {'存在' if os.path.isfile(p) else '缺失'} | {p}")
        except Exception:
            pass
        # resolve selected mode and switch working dirs
        sel_label = (self.mode_var.get() or '').strip()
        mode = self.mode_display_to_key.get(sel_label, sel_label)
        if not mode:
            messagebox.showwarning('提示', '请选择模式')
            return
        self.map_dir = os.path.join(self.map_root, mode)
        self.json_dir = os.path.join(self.json_root, mode)
        # ensure paths
        os.makedirs(self.map_dir, exist_ok=True)
        os.makedirs(self.json_dir, exist_ok=True)
        try:
            map_exists = os.path.isdir(self.map_dir)
            json_exists = os.path.isdir(self.json_dir)
            map_cnt = len([f for f in os.listdir(self.map_dir) if f.lower().endswith('.png')]) if map_exists else 0
            json_cnt = len([f for f in os.listdir(self.json_dir) if f.lower().endswith('.json')]) if json_exists else 0
            self._log(f"[自检] 模式={mode} map_dir={self.map_dir} 存在={map_exists} png数={map_cnt}")
            self._log(f"[自检] 模式={mode} json_dir={self.json_dir} 存在={json_exists} json数={json_cnt}")
            if map_cnt == 0:
                self._log("[警告] 该模式的地图模板目录为空，请将 png 放入 map/" + mode)
            if json_cnt == 0:
                self._log("[警告] 该模式的脚本目录为空，请将 json 放入 json/" + mode)
        except Exception:
            pass
        self.running = True
        self.stop_event.clear()
        # pass selected mode to runner
        self.worker = threading.Thread(target=self._run_mode_loop, args=(mode,), daemon=True)
        self.worker.start()
        self._log('开始运行脚本')
        # Start status updater
        try:
            self._schedule_status_update()
        except Exception:
            pass
        # Start auto-stop timer (F12-style)
        try:
            if self.auto_stop_timer and self.auto_stop_timer.is_alive():
                self.auto_stop_timer.cancel()
        except Exception:
            pass
        try:
            if self.auto_stop_seconds:
                def _timeout_stop():
                    try:
                        self.on_stop()
                    except Exception:
                        pass
                self.auto_stop_timer = threading.Timer(self.auto_stop_seconds, _timeout_stop)
                self.auto_stop_timer.daemon = True
                self.auto_stop_timer.start()
        except Exception:
            pass

    def on_stop(self):
        if not self.running:
            return
        self.running = False
        self.stop_event.set()
        self._log('请求停止，等待当前步骤结束...')
        try:
            if self.auto_stop_timer and self.auto_stop_timer.is_alive():
                self.auto_stop_timer.cancel()
        except Exception:
            pass

    # Core loop
    def _run_mode_loop(self, mode_name: str):
        """Dynamically import logic.<mode_name> and run it with self."""
        try:
            # add base_dir to sys.path to import local packages
            if self.base_dir not in sys.path:
                sys.path.insert(0, self.base_dir)
            mod = importlib.import_module(f'logic.{mode_name}')
        except Exception as e:
            self._log(f'加载模式失败: {mode_name} ({e})')
            self.running = False
            return
        ensure_restored(self.selected_hwnd)
        self._log(f'已启动模式: {mode_name}')
        try:
            if hasattr(mod, 'run'):
                mod.run(self)
            else:
                self._log(f'模式 {mode_name} 不包含 run(app)')
        except Exception as e:
            self._log(f'模式运行异常: {e}')
        finally:
            self._log('脚本已停止。')

    def _wait_and_click(self, template_filename, name_alias, timeout=None):
        if timeout is None:
            timeout = self.timeout_seconds
        deadline = time.time() + timeout
        tpl_path = os.path.join(self.control_dir, template_filename)
        self._log(f"等待 {name_alias}_button，超时{timeout:.0f}s …")
        while self.running and not self.stop_event.is_set() and time.time() < deadline:
            img = self.capturer.capture_background()
            m = match_template(img, tpl_path, self.threshold)
            if m:
                cx, cy = m['center']
                self._log(f"识别到 {name_alias}_button (score={m['score']:.2f})，点击中心: ({cx},{cy})")
                # Convert to client coordinates for SendMessage
                win_left, win_top, _, _ = win32gui.GetWindowRect(self.selected_hwnd)
                client_pt = win32gui.ScreenToClient(self.selected_hwnd, (win_left + cx, win_top + cy))
                tx, ty = client_pt
                target = child_from_client_point(self.selected_hwnd, tx, ty)
                # slightly longer click: down -> short hold -> up
                lp = _pack_lparam(tx, ty)
                send_mouse_move(target, tx, ty)
                win32gui.SendMessage(target, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lp)
                # short, interruptible hold (~80ms)
                hold_ms = 0.08
                waited = 0.0
                step = 0.01
                while waited < hold_ms and self.running and not self.stop_event.is_set():
                    time.sleep(step)
                    waited += step
                win32gui.SendMessage(target, win32con.WM_LBUTTONUP, 0, lp)
                # interruptible wait
                waited = 0.0
                step = 0.05
                while waited < self.post_click_wait and self.running and not self.stop_event.is_set():
                    time.sleep(step)
                    waited += step
                return True
            # interruptible retry sleep
            waited = 0.0
            step = 0.05
            while waited < self.retry_interval and self.running and not self.stop_event.is_set():
                time.sleep(step)
                waited += step
        self._log(f"等待 {name_alias}_button 超时，已停止。")
        self.running = False
        return False

    def _wait_detect(self, template_filename, name_alias, timeout=None):
        if timeout is None:
            timeout = self.timeout_seconds
        deadline = time.time() + timeout
        tpl_path = os.path.join(self.control_dir, template_filename)
        while self.running and not self.stop_event.is_set() and time.time() < deadline:
            img = self.capturer.capture_background()
            m = match_template(img, tpl_path, self.threshold)
            if m:
                self._log(f"检测到 {name_alias} (score={m['score']:.2f})")
                return True
            waited = 0.0
            step = 0.05
            while waited < self.retry_interval and self.running and not self.stop_event.is_set():
                time.sleep(step)
                waited += step
        self._log(f"等待 {name_alias} 超时，已停止。")
        self.running = False
        return False

    def _try_wait_and_click(self, template_filename, name_alias, timeout=3.0):
        """Like _wait_and_click but optional: timeout does NOT stop the loop.
        Returns True if clicked, False if not found in time or interrupted.
        """
        deadline = time.time() + (timeout or 0)
        tpl_path = os.path.join(self.control_dir, template_filename)
        self._log(f"尝试点击 {name_alias}_button（可选），超时{timeout:.1f}s …")
        while self.running and not self.stop_event.is_set() and time.time() < deadline:
            img = self.capturer.capture_background()
            m = match_template(img, tpl_path, self.threshold)
            if m:
                cx, cy = m['center']
                self._log(f"识别到 {name_alias}_button (score={m['score']:.2f})，点击中心: ({cx},{cy})")
                win_left, win_top, _, _ = win32gui.GetWindowRect(self.selected_hwnd)
                tx, ty = win32gui.ScreenToClient(self.selected_hwnd, (win_left + cx, win_top + cy))
                target = child_from_client_point(self.selected_hwnd, tx, ty)
                lp = _pack_lparam(tx, ty)
                send_mouse_move(target, tx, ty)
                win32gui.SendMessage(target, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lp)
                waited = 0.0
                step = 0.01
                while waited < 0.08 and self.running and not self.stop_event.is_set():
                    time.sleep(step)
                    waited += step
                win32gui.SendMessage(target, win32con.WM_LBUTTONUP, 0, lp)
                waited = 0.0
                step = 0.05
                while waited < self.post_click_wait and self.running and not self.stop_event.is_set():
                    time.sleep(step)
                    waited += step
                return True
            # retry interval
            waited = 0.0
            step = 0.05
            while waited < self.retry_interval and self.running and not self.stop_event.is_set():
                time.sleep(step)
                waited += step
        return False

    def _wait_and_click_either(self, choices, timeout=None):
        """choices: list of (filename, alias). Click whichever appears first.
        Returns alias string if clicked, else False when timeout/stop.
        """
        if timeout is None:
            timeout = self.timeout_seconds
        deadline = time.time() + timeout
        # prebuild paths
        files = [(os.path.join(self.control_dir, fn), alias) for fn, alias in choices]
        while self.running and not self.stop_event.is_set() and time.time() < deadline:
            img = self.capturer.capture_background()
            best = None
            for path, alias in files:
                m = match_template(img, path, self.threshold)
                if m:
                    # prefer the highest score among matches in the same frame
                    score = m['score']
                    if (best is None) or (score > best[0]):
                        best = (score, m, alias, path)
            if best is not None:
                score, m, alias, path = best
                cx, cy = m['center']
                self._log(f"识别到 {alias}_button (score={score:.2f})，点击中心: ({cx},{cy})")
                win_left, win_top, _, _ = win32gui.GetWindowRect(self.selected_hwnd)
                tx, ty = win32gui.ScreenToClient(self.selected_hwnd, (win_left + cx, win_top + cy))
                target = child_from_client_point(self.selected_hwnd, tx, ty)
                lp = _pack_lparam(tx, ty)
                send_mouse_move(target, tx, ty)
                win32gui.SendMessage(target, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lp)
                # short, interruptible hold (~80ms)
                hold_ms = 0.08
                waited = 0.0
                step = 0.01
                while waited < hold_ms and self.running and not self.stop_event.is_set():
                    time.sleep(step)
                    waited += step
                win32gui.SendMessage(target, win32con.WM_LBUTTONUP, 0, lp)
                # post click wait
                waited = 0.0
                step = 0.05
                while waited < self.post_click_wait and self.running and not self.stop_event.is_set():
                    time.sleep(step)
                    waited += step
                return alias
            # retry interval
            waited = 0.0
            step = 0.05
            while waited < self.retry_interval and self.running and not self.stop_event.is_set():
                time.sleep(step)
                waited += step
        self._log("等待按钮(确认选择/开始挑战)超时，已停止。")
        self.running = False
        return False

    def _recognize_map_name(self):
        img = self.capturer.capture_background()
        if img is None:
            return None
        edge_img = _edges1ch(img)
        # Three-scale best score for a single template path
        def _score_for(path):
            if path not in self._tpl_edge_cache:
                e, mask = _load_template_edge_and_mask(path)
                self._tpl_edge_cache[path] = (e, mask)
            e, mask = self._tpl_edge_cache[path]
            if e is None or mask is None:
                return None
            try:
                # three-scale matching: pick best among [1.0, 0.95, 0.9]
                if e.size == 0 or mask.size == 0 or np.count_nonzero(mask) == 0:
                    return None
                ih, iw = edge_img.shape[:2]
                base_h, base_w = e.shape[:2]
                best = 0.0
                for scale in (1.0, 0.95, 0.9):
                    if scale != 1.0:
                        th = max(1, int(round(base_h * scale)))
                        tw = max(1, int(round(base_w  * scale)))
                        e_s = cv2.resize(e, (tw, th), interpolation=cv2.INTER_AREA)
                        m_s = cv2.resize(mask, (tw, th), interpolation=cv2.INTER_NEAREST)
                    else:
                        e_s, m_s = e, mask
                    th, tw = e_s.shape[:2]
                    if th > ih or tw > iw:
                        continue
                    if np.count_nonzero(m_s) == 0:
                        continue
                    res = cv2.matchTemplate(edge_img, e_s, cv2.TM_CCORR_NORMED, mask=m_s)
                    _, max_val, _, _ = cv2.minMaxLoc(res)
                    mv = float(max_val)
                    if not np.isfinite(mv):
                        mv = 0.0
                    if mv > best:
                        best = mv
                return best
            except Exception:
                return None

        # Build candidates grouped by canonical base name (strip -2/-3 suffix)
        pngs = [f for f in os.listdir(self.map_dir) if f.lower().endswith('.png')]
        groups = {}
        for f in pngs:
            nm = os.path.splitext(f)[0]
            if nm.endswith('-2'):
                key = nm[:-2]
                groups.setdefault(key, {})['f2'] = f
            elif nm.endswith('-3'):
                key = nm[:-2]
                groups.setdefault(key, {})['f3'] = f
            else:
                key = nm
                groups.setdefault(key, {})['base'] = f
        candidates = []
        for base_name, files in groups.items():
            p1 = os.path.join(self.map_dir, files['base']) if 'base' in files else None
            p2 = os.path.join(self.map_dir, files['f2']) if 'f2' in files else None
            p3 = os.path.join(self.map_dir, files['f3']) if 'f3' in files else None
            s1 = _score_for(p1) if p1 else 0.0
            s2 = _score_for(p2) if p2 else 0.0
            s3 = _score_for(p3) if p3 else 0.0
            hits = int((s1 or 0.0) >= self.feature_presence_thr) + \
                   int((s2 or 0.0) >= self.feature_presence_thr) + \
                   int((s3 or 0.0) >= self.feature_presence_thr)
            total = (s1 or 0.0) + (s2 or 0.0) + (s3 or 0.0)
            mx = max((s1 or 0.0), (s2 or 0.0), (s3 or 0.0))
            candidates.append((base_name, hits, total, mx, (s1 if p1 else -1.0), (s2 if p2 else -1.0), (s3 if p3 else -1.0)))
        if not candidates:
            return None
        candidates.sort(key=lambda x: (x[1], x[2], x[3]), reverse=True)
        def _fmt(n, hits, total, mx, s1, s2, s3):
            parts = [f"hits={hits}", f"sum={total:.2f}", f"max={mx:.2f}"]
            feat_parts = [f"base={s1:.2f}" if s1 >= 0 else "base=-",
                          f"feat2={s2:.2f}" if s2 >= 0 else "feat2=-",
                          f"feat3={s3:.2f}" if s3 >= 0 else "feat3=-"]
            return f"{n}:" + ",".join(parts) + " (" + ",".join(feat_parts) + ")"
        top3 = ', '.join([_fmt(n, h, t, m, s1, s2, s3) for (n, h, t, m, s1, s2, s3) in candidates[:3]])
        self._log(f"地图匹配Top3: {top3}")
        top1_name, top1_hits, top1_total, top1_max, *_ = candidates[0]
        if top1_max >= self.quick_accept_thr and top1_hits >= 1:
            self._log(f"地图识别为 {top1_name} (quick_accept,max={top1_max:.2f},hits={top1_hits})")
            return top1_name
        if top1_hits >= self.min_hits:
            self._log(f"地图识别为 {top1_name} (hits={top1_hits}, sum={top1_total:.2f})")
            return top1_name
        self._log(f"地图匹配不确定 (top1_hits={top1_hits}, top1_max={top1_max:.2f})")
        return None

    def _load_actions(self, map_name):
        def _resolve_json_path(name):
            safe_name = (name or '').strip()
            cand = os.path.join(self.json_dir, f"{safe_name}.json")
            if os.path.isfile(cand):
                return cand
            # case-insensitive fallback
            try:
                target = f"{safe_name}.json".lower()
                for f in os.listdir(self.json_dir):
                    if f.lower() == target:
                        return os.path.join(self.json_dir, f)
            except Exception:
                pass
            return cand

        json_path = _resolve_json_path(map_name)
        if not os.path.isfile(json_path):
            self._log(f"未找到动作文件: {json_path}")
            try:
                files = [f for f in os.listdir(self.json_dir) if f.lower().endswith('.json')]
                self._log(f"可用脚本: {', '.join(files) if files else '无'}")
            except Exception:
                pass
            return None
        try:
            self._log(f"加载脚本: {json_path}")
            with open(json_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            return data.get('steps', [])
        except Exception as e:
            self._log(f"读取 {json_path} 失败: {e}")
            return None


    # Periodic status updater
    def _schedule_status_update(self):
        def _fmt_time(sec):
            try:
                sec = int(max(0, sec))
                h = sec // 3600
                m = (sec % 3600) // 60
                s = sec % 60
                if h:
                    return f"{h}h{m:02d}m{s:02d}s"
                if m:
                    return f"{m}m{s:02d}s"
                return f"{s}s"
            except Exception:
                return "0s"
        try:
            if self.running:
                remain_loops = (self.max_loops - self.loops_done) if self.max_loops else 0
                elapsed = time.time() - (self.started_at or time.time())
                remain_time = (self.auto_stop_seconds - elapsed) if self.auto_stop_seconds else 0
                loops_part = f"剩余循环: {remain_loops}" if self.max_loops else "循环: 不限"
                time_part = f"剩余时间: {_fmt_time(remain_time)}" if self.auto_stop_seconds else "定时关闭: 关闭"
                self.status_var.set(f"状态: 运行中 | {loops_part} | {time_part}")
            else:
                self.status_var.set("状态: 已停止")
        except Exception:
            pass
        # reschedule
        try:
            self.root.after(500, self._schedule_status_update)
        except Exception:
            pass


def main():
    root = tk.Tk()
    app = App(root)
    root.geometry('820x520')
    root.mainloop()


if __name__ == '__main__':
    main()

