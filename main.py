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
from tkinter import messagebox
import importlib
import ttkbootstrap as ttk


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
        self.theme_name = 'cosmo'           # 窗口主题：白天cosmo/黑夜darkly
        self.started_at = None
        self.loops_done = 0
        self.auto_stop_timer = None

        # Try load persisted config
        self._load_config()
        # Apply ttk theme early (no UI rebuild) so initial widgets use correct palette
        try:
            if hasattr(self.root, 'style') and self.root.style:
                self.root.style.theme_use(self.theme_name)
            else:
                from ttkbootstrap import Style  # type: ignore
                st = Style(theme=self.theme_name)
                # attach to root for later use
                try:
                    self.root.style = st
                except Exception:
                    pass
        except Exception:
            pass

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

    # Helpers: theme-based colors
    def _get_page_bg(self):
        try:
            if str(self.theme_name).lower() in ('darkly', 'cyborg', 'superhero', 'solar', 'vapor'):
                return '#1c1f24'
        except Exception:
            pass
        return '#f2f4f7'

    # Helpers: rounded card UI
    def _create_card(self, parent, title: str, padding=10, radius=20):
        dark = str(self.theme_name).lower() in ('darkly', 'cyborg', 'superhero', 'solar', 'vapor')
        bg = self._get_page_bg()
        if dark:
            card_fill = '#2b2f33'
            stroke = '#3a3f45'
            shadow_layers = [('#23282d', 1), ('#20252a', 2), ('#1e2328', 3)]
        else:
            card_fill = '#ffffff'
            stroke = '#e6e9ef'
            shadow_layers = [('#eef2f8', 1), ('#e8edf6', 2), ('#e3e9f2', 3)]
        container = tk.Frame(parent, bg=bg, highlightthickness=0)
        container.grid_propagate(False)
        # canvas for rounded background
        cv = tk.Canvas(container, bg=bg, highlightthickness=0, bd=0)
        cv.pack(fill='both', expand=True)
        inner = ttk.Frame(cv, padding=padding)
        # inner uses grid manager exclusively
        inner.columnconfigure(0, weight=1)
        # header
        if title:
            ttk.Label(inner, text=title).grid(row=0, column=0, sticky='w', pady=(0,6))
        # content frame for caller widgets (return this)
        content = ttk.Frame(inner)
        content.grid(row=1, column=0, sticky='nsew')
        inner.rowconfigure(1, weight=1)

        # place inner as canvas window
        win_id = cv.create_window(0, 0, window=inner, anchor='nw')

        def _redraw(_evt=None):
            try:
                w = max(40, container.winfo_width())
                h = max(40, container.winfo_height())
                cv.config(width=w, height=h)
                cv.delete('card')
                cv.delete('shadow')
                cv.delete('stroke')
                r = radius
                x1, y1, x2, y2 = 2, 2, w-2, h-2
                # layered soft shadow (down-right)
                for col, off in shadow_layers:
                    sx1, sy1, sx2, sy2 = x1+off, y1+off, x2+off, y2+off
                    cv.create_arc(sx1, sy1, sx1+2*r, sy1+2*r, start=90, extent=90, style='pieslice', fill=col, outline='', tags='shadow')
                    cv.create_arc(sx2-2*r, sy1, sx2, sy1+2*r, start=0, extent=90, style='pieslice', fill=col, outline='', tags='shadow')
                    cv.create_arc(sx1, sy2-2*r, sx1+2*r, sy2, start=180, extent=90, style='pieslice', fill=col, outline='', tags='shadow')
                    cv.create_arc(sx2-2*r, sy2-2*r, sx2, sy2, start=270, extent=90, style='pieslice', fill=col, outline='', tags='shadow')
                    cv.create_rectangle(sx1+r, sy1, sx2-r, sy2, fill=col, outline='', tags='shadow')
                    cv.create_rectangle(sx1, sy1+r, sx2, sy2-r, fill=col, outline='', tags='shadow')

                # draw rounded rect via four arcs + rects (card body)
                cv.create_arc(x1, y1, x1+2*r, y1+2*r, start=90, extent=90, style='pieslice', fill=card_fill, outline='', tags='card')
                cv.create_arc(x2-2*r, y1, x2, y1+2*r, start=0, extent=90, style='pieslice', fill=card_fill, outline='', tags='card')
                cv.create_arc(x1, y2-2*r, x1+2*r, y2, start=180, extent=90, style='pieslice', fill=card_fill, outline='', tags='card')
                cv.create_arc(x2-2*r, y2-2*r, x2, y2, start=270, extent=90, style='pieslice', fill=card_fill, outline='', tags='card')
                cv.create_rectangle(x1+r, y1, x2-r, y2, fill=card_fill, outline='', tags='card')
                cv.create_rectangle(x1, y1+r, x2, y2-r, fill=card_fill, outline='', tags='card')
                # stroke: rounded border (arcs + edges)
                cv.create_arc(x1, y1, x1+2*r, y1+2*r, start=90, extent=90, style='arc', outline=stroke, width=1, tags='stroke')
                cv.create_arc(x2-2*r, y1, x2, y1+2*r, start=0, extent=90, style='arc', outline=stroke, width=1, tags='stroke')
                cv.create_arc(x1, y2-2*r, x1+2*r, y2, start=180, extent=90, style='arc', outline=stroke, width=1, tags='stroke')
                cv.create_arc(x2-2*r, y2-2*r, x2, y2, start=270, extent=90, style='arc', outline=stroke, width=1, tags='stroke')
                cv.create_line(x1+r, y1, x2-r, y1, fill=stroke, width=1, tags='stroke')
                cv.create_line(x1+r, y2, x2-r, y2, fill=stroke, width=1, tags='stroke')
                cv.create_line(x1, y1+r, x1, y2-r, fill=stroke, width=1, tags='stroke')
                cv.create_line(x2, y1+r, x2, y2-r, fill=stroke, width=1, tags='stroke')
                # place inner with padding inset
                cv.coords(win_id, x1+8, y1+8)
                cv.itemconfigure(win_id, width=max(20, w-16), height=max(20, h-16))
            except Exception:
                pass

        container.bind('<Configure>', _redraw)
        return container, content

    # UI
    def _build_ui(self):
        # light gray background to enhance card depth
        wrapper = ttk.Frame(self.root, padding=12)
        wrapper.pack(fill='both', expand=True)

        # top bar title
        title = ttk.Label(wrapper, text='皎皎助手后台挂机1.0', font=('Microsoft YaHei', 16, 'bold'))
        title.pack(anchor='w', pady=(0, 8))
        # Settings button at top-right
        ttk.Button(wrapper, text='设置', command=self._open_settings).place(relx=1.0, x=-12, y=10, anchor='ne')

        # three columns container
        # outer page background tweak
        try:
            wrapper.configure(style='')
            wrapper['bootstyle'] = ''
        except Exception:
            pass

        # use tk.Frame to control background color for better contrast
        cols = tk.Frame(wrapper, bg=self._get_page_bg())
        cols.pack(fill='both', expand=True)
        # make middle narrower, right wider; allow row to expand
        cols.columnconfigure(0, weight=1, uniform='col')
        cols.columnconfigure(1, weight=1, uniform='col')
        cols.columnconfigure(2, weight=2, uniform='col')
        cols.rowconfigure(0, weight=1)

        # Left sidebar: rounded card
        left_card, left = self._create_card(cols, '任务列表', padding=10)
        left_card.grid(row=0, column=0, sticky='nsew', padx=(0, 8))
        left.rowconfigure(0, weight=1)

        # task check list
        task_frame = ttk.Frame(left)
        task_frame.grid(row=0, column=0, sticky='nsew')
        self.mode_var = tk.StringVar(value='')
        self.task_vars = {}
        self.tasks_map = {
            '55mod': self.mode_name_map.get('55mod', '55mod'),
            'juesemihan': self.mode_name_map.get('juesemihan', 'juesemihan'),
            'wuqimihan': self.mode_name_map.get('wuqimihan', 'wuqimihan'),
        }
        def _mk_chk(key):
            v = tk.BooleanVar(value=(key == '55mod'))
            self.task_vars[key] = v
            def _cb():
                self._on_task_toggle(key)
            ttk.Checkbutton(task_frame, text=self.tasks_map[key], variable=v, command=_cb).pack(anchor='w', pady=2)
        for k in ['55mod', 'juesemihan', 'wuqimihan']:
            _mk_chk(k)
        # default select 55mod
        self.mode_var.set(self.mode_name_map.get('55mod', '55mod'))

        # bottom controls in left
        left_bottom = ttk.Frame(left)
        left_bottom.grid(row=1, column=0, sticky='ew', pady=(8,0))
        left_bottom.columnconfigure(0, weight=1)
        left_bottom.columnconfigure(1, weight=1)
        self.btn_start = ttk.Button(left_bottom, text='开始任务 (F10)', command=self.on_start, bootstyle='primary')
        self.btn_start.grid(row=0, column=0, sticky='ew', padx=(0,4))
        self.btn_stop = ttk.Button(left_bottom, text='停止 (F12)', command=self.on_stop)
        self.btn_stop.grid(row=0, column=1, sticky='ew', padx=(4,0))

        # Middle: task settings card
        mid_card, mid = self._create_card(cols, '任务设置', padding=10)
        mid_card.grid(row=0, column=1, sticky='nsew', padx=8)

        # 武器密函
        row_m = ttk.Frame(mid)
        row_m.pack(fill='x', pady=4)
        ttk.Label(row_m, text='武器密函选择').pack(side='left')
        self.wuqi_mihan_var = tk.StringVar(value='')
        self.combo_wuqi_mihan = ttk.Combobox(row_m, state='readonly', width=30, textvariable=self.wuqi_mihan_var)
        self.combo_wuqi_mihan.pack(side='left', padx=6)
        ttk.Button(row_m, text='刷新密函', command=self._refresh_wuqi_mihan).pack(side='left')

        row_j = ttk.Frame(mid)
        row_j.pack(fill='x', pady=4)
        ttk.Label(row_j, text='角色密函选择').pack(side='left')
        self.juese_mihan_var = tk.StringVar(value='')
        self.combo_juese_mihan = ttk.Combobox(row_j, state='readonly', width=30, textvariable=self.juese_mihan_var)
        self.combo_juese_mihan.pack(side='left', padx=6)
        ttk.Button(row_j, text='刷新密函', command=self._refresh_juese_mihan).pack(side='left')

        # 温馨提示（小卡片）
        tip_card, tip_box = self._create_card(mid, '温馨提示', padding=8, radius=10)
        tip_card.pack(fill='x', pady=(10,0))
        ttk.Label(tip_box, text='进入对应副本列表，点击前往进入确认选择界面\n再开启脚本\n在运行前您需要确保：\n1.pc设置为100%缩放\n2.游戏窗口设置为16：9，1920x1080\n3.使用管理员权限运行脚本\n5.不要更改文件结构，若要替换模板图片请按原\n名命名！\n6.若无论如何无法运行脚本，请描述设备情况并\n复制日志发给1484413790@qq.com').pack(anchor='w')

        # Right: connection + logs card
        right_card, right = self._create_card(cols, '连接 / 日志', padding=10)
        right_card.grid(row=0, column=2, sticky='nsew', padx=(8,0))
        right.rowconfigure(3, weight=1)

        # refresh and window selector
        row_conn = ttk.Frame(right)
        row_conn.grid(row=0, column=0, sticky='ew')
        self.btn_refresh = ttk.Button(row_conn, text='刷新窗口 (F9)', command=self.refresh_windows)
        self.btn_refresh.pack(side='left')
        ttk.Label(row_conn, text='当前窗口').pack(side='left', padx=(10,4))
        self.combo = ttk.Combobox(row_conn, state='readonly', width=40)
        self.combo.pack(side='left', padx=4, fill='x', expand=True)
        self.combo.bind('<<ComboboxSelected>>', self.on_combo_select)

        # 日志标题 + 清空按钮
        row_log_title = ttk.Frame(right)
        row_log_title.grid(row=2, column=0, sticky='ew', pady=(8,0))
        ttk.Label(row_log_title, text='日志').pack(side='left')
        ttk.Button(row_log_title, text='⟳', width=3, command=self.clear_log).pack(side='right')

        # Log box (read-only)
        self.log = tk.Text(right, height=18, state='disabled')
        self.log.grid(row=3, column=0, sticky='nsew', pady=(4,0))

        # hidden mode combobox kept for compatibility with _refresh_modes
        self.combo_mode = ttk.Combobox(wrapper, state='readonly', textvariable=self.mode_var)
        self.combo_mode.place_forget()

        # Status bar
        status_frame = ttk.Frame(wrapper)
        status_frame.pack(fill='x', pady=(8,0))
        self.status_var = tk.StringVar(value='状态: 就绪')
        ttk.Label(status_frame, textvariable=self.status_var, anchor='w').pack(side='left', fill='x', expand=True)

        # init modes and lists
        self._refresh_modes()
        self._refresh_wuqi_mihan()
        self._refresh_juese_mihan()

    def _bind_hotkeys(self):
        # Local Tk bindings (also support global with keyboard if installed)
        try:
            import keyboard  # type: ignore
            keyboard.add_hotkey('f10', self.on_start)
            keyboard.add_hotkey('f12', self.on_stop)
            keyboard.add_hotkey('f9', self.refresh_windows)
        except Exception:
            pass

    # Left task list single-select behavior
    def _on_task_toggle(self, selected_key: str):
        try:
            # uncheck others
            for k, var in self.task_vars.items():
                if k != selected_key:
                    var.set(False)
            # ensure selected set to True
            self.task_vars[selected_key].set(True)
            # sync mode_var to display label once modes are built
            display = self.mode_key_to_display.get(selected_key, self.tasks_map.get(selected_key, selected_key))
            self.mode_var.set(display)
        except Exception:
            pass

    def clear_log(self):
        try:
            self.log.configure(state='normal')
            self.log.delete('1.0', 'end')
            self.log.configure(state='disabled')
        except Exception:
            pass
        # also truncate log file quietly
        try:
            if os.path.isfile(self.log_file_path):
                with open(self.log_file_path, 'w', encoding='utf-8') as _f:
                    _f.write('')
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

    def _refresh_juese_mihan(self):
        try:
            folder = os.path.join(self.control_dir, '角色密函png')
            names = []
            if os.path.isdir(folder):
                for fn in os.listdir(folder):
                    if fn.lower().endswith('.png'):
                        names.append(os.path.splitext(fn)[0])
            names.sort()
            self.combo_juese_mihan['values'] = names
            if names:
                cur = self.juese_mihan_var.get()
                if cur not in names:
                    self.juese_mihan_var.set(names[0])
            else:
                self.juese_mihan_var.set('')
            self._log(f"已刷新角色密函列表: {', '.join(names) if names else '无'}")
        except Exception as e:
            self._log(f"刷新角色密函失败: {e}")

    def get_selected_juese_mihan_path(self):
        name = (self.juese_mihan_var.get() or '').strip()
        if not name:
            return None
        return os.path.join(self.control_dir, '角色密函png', f"{name}.png")

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
            self.log.configure(state='normal')
            self.log.insert('end', line)
            self.log.see('end')
            self.log.configure(state='disabled')
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
                self.theme_name = str(cfg.get('theme', self.theme_name))
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
                'theme': self.theme_name,
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
        self.var_theme = tk.StringVar(value=self.theme_name)

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

        # Card: 窗口模式（主题）
        lf_theme = ttk.Labelframe(frm, text='窗口模式', padding=8)
        lf_theme.pack(fill='x', pady=(8,4))
        rb_row = ttk.Frame(lf_theme)
        rb_row.pack(anchor='w')
        ttk.Radiobutton(rb_row, text='白天模式', value='cosmo', variable=self.var_theme).pack(side='left', padx=(0,12))
        ttk.Radiobutton(rb_row, text='黑夜模式', value='darkly', variable=self.var_theme).pack(side='left')

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
                # theme
                self.theme_name = str(self.var_theme.get() or self.theme_name)
                self._clamp_settings()
                self._log(f"已应用设置: 随机脚本={self.fail_fallback_random}, 延迟={self.post_likai_delay}s, 循环次数={self.max_loops}, 定时关闭={self.auto_stop_seconds}s")
                # apply theme now
                self._apply_theme(self.theme_name)
                self._save_config()
            except Exception as e:
                self._log(f"应用设置失败: {e}")
            try:
                win.destroy()
            except Exception:
                pass
        ttk.Button(btns, text='保存', command=_apply_and_close).pack(side='right')

    def _apply_theme(self, theme_name: str):
        try:
            # ttkbootstrap Window has style accessible
            if hasattr(self.root, 'style') and self.root.style:
                self.root.style.theme_use(theme_name)
            else:
                # fallback: create a temporary Style
                from ttkbootstrap import Style  # type: ignore
                st = Style()
                st.theme_use(theme_name)
            # Rebuild UI so custom canvas-based cards use the new palette
            try:
                for ch in self.root.winfo_children():
                    ch.destroy()
            except Exception:
                pass
            self._build_ui()
            self._bind_hotkeys()
            self.refresh_windows()
        except Exception as e:
            self._log(f"切换主题失败: {e}")

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
            self._log(f"🔧 [自检] base_dir={self.base_dir}")
            self._log(f"📁 [自检] control={self.control_dir} 存在={os.path.isdir(self.control_dir)}")
            for fn in ['xuanzemihan.png', 'bushiyong.png', 'querenxuanze.png', 'likai.png', 'zaicijinixng.png']:
                p = os.path.join(self.control_dir, fn)
                self._log(f"🖼️ [自检] 模板 {fn}: {'存在' if os.path.isfile(p) else '缺失'} | {p}")
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
            self._log(f"📊 [自检] 模式={mode} map_dir={self.map_dir} 存在={map_exists} png数={map_cnt}")
            self._log(f"📊 [自检] 模式={mode} json_dir={self.json_dir} 存在={json_exists} json数={json_cnt}")
            if map_cnt == 0:
                self._log("⚠️ [警告] 该模式的地图模板目录为空，请将 png 放入 map/" + mode)
            if json_cnt == 0:
                self._log("⚠️ [警告] 该模式的脚本目录为空，请将 json 放入 json/" + mode)
        except Exception:
            pass
        self.running = True
        self.stop_event.clear()
        # pass selected mode to runner
        self.worker = threading.Thread(target=self._run_mode_loop, args=(mode,), daemon=True)
        self.worker.start()
        self._log('🚀 开始运行脚本')
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
        self._log('⏹️ 请求停止，等待当前步骤结束...')
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
            self._log(f'❌ 加载模式失败: {mode_name} ({e})')
            self.running = False
            return
        ensure_restored(self.selected_hwnd)
        self._log(f'▶️ 已启动模式: {mode_name}')
        try:
            if hasattr(mod, 'run'):
                mod.run(self)
            else:
                self._log(f'⚠️ 模式 {mode_name} 不包含 run(app)')
        except Exception as e:
            self._log(f'⚠️ 模式运行异常: {e}')
        finally:
            self._log('🛑 脚本已停止。')

    def _wait_and_click(self, template_filename, name_alias, timeout=None):
        if timeout is None:
            timeout = self.timeout_seconds
        deadline = time.time() + timeout
        tpl_path = os.path.join(self.control_dir, template_filename)
        self._log(f"⏳ 等待 {name_alias}_button，超时{timeout:.0f}s …")
        while self.running and not self.stop_event.is_set() and time.time() < deadline:
            img = self.capturer.capture_background()
            m = match_template(img, tpl_path, self.threshold)
            if m:
                cx, cy = m['center']
                self._log(f"🔍 识别到 {name_alias}_button (score={m['score']:.2f})，点击中心: ({cx},{cy})")
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
        self._log(f"⏰ 等待 {name_alias}_button 超时，已停止。")
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
                self._log(f"🔍 检测到 {name_alias} (score={m['score']:.2f})")
                return True
            waited = 0.0
            step = 0.05
            while waited < self.retry_interval and self.running and not self.stop_event.is_set():
                time.sleep(step)
                waited += step
        self._log(f"⏰ 等待 {name_alias} 超时，已停止。")
        self.running = False
        return False

    def _try_wait_and_click(self, template_filename, name_alias, timeout=3.0):
        """Like _wait_and_click but optional: timeout does NOT stop the loop.
        Returns True if clicked, False if not found in time or interrupted.
        """
        deadline = time.time() + (timeout or 0)
        tpl_path = os.path.join(self.control_dir, template_filename)
        self._log(f"🖱️ 尝试点击 {name_alias}_button（可选），超时{timeout:.1f}s …")
        while self.running and not self.stop_event.is_set() and time.time() < deadline:
            img = self.capturer.capture_background()
            m = match_template(img, tpl_path, self.threshold)
            if m:
                cx, cy = m['center']
                self._log(f"🔍 识别到 {name_alias}_button (score={m['score']:.2f})，点击中心: ({cx},{cy})")
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
                self._log(f"🔍 识别到 {alias}_button (score={score:.2f})，点击中心: ({cx},{cy})")
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
            total = (s1 or 0.0) + (s2 or 0.0) + (s3 or 0.0)
            candidates.append((base_name, total, (s1 if p1 else -1.0), (s2 if p2 else -1.0), (s3 if p3 else -1.0)))
        if not candidates:
            return None
        # sort by total score only
        candidates.sort(key=lambda x: x[1], reverse=True)
        def _fmt(n, total, s1, s2, s3):
            feat_parts = [f"base={s1:.2f}" if s1 >= 0 else "base=-",
                          f"feat2={s2:.2f}" if s2 >= 0 else "feat2=-",
                          f"feat3={s3:.2f}" if s3 >= 0 else "feat3=-"]
            return f"{n}: sum={total:.2f} (" + ",".join(feat_parts) + ")"
        top3 = ', '.join([_fmt(n, t, s1, s2, s3) for (n, t, s1, s2, s3) in candidates[:3]])
        self._log(f"地图匹配Top3: {top3}")
        top1_name, top1_total, *_ = candidates[0]
        self._log(f"地图识别为 {top1_name} (sum={top1_total:.2f})")
        return top1_name

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
    root = ttk.Window(themename='cosmo')
    app = App(root)
    # Larger initial size; lock as minimum so users can only enlarge
    initial_w, initial_h = 1280, 760
    root.geometry(f'{initial_w}x{initial_h}')
    try:
        root.update_idletasks()
        root.minsize(initial_w, initial_h)
    except Exception:
        pass
    root.mainloop()


if __name__ == '__main__':
    main()

