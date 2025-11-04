import os
import json
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading

import win32con
import win32gui
import win32api

# ---- Window helpers ----

def enumerate_windows():
    wins = []
    def _cb(h, _):
        if win32gui.IsWindowVisible(h) and win32gui.GetWindowText(h):
            wins.append((h, win32gui.GetClassName(h), win32gui.GetWindowText(h)))
    win32gui.EnumWindows(_cb, None)
    return wins


def ensure_restored(hwnd):
    try:
        placement = win32gui.GetWindowPlacement(hwnd)
        if placement[1] == win32con.SW_SHOWMINIMIZED:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            time.sleep(0.2)
    except Exception:
        pass


def get_client_center(hwnd):
    l, t, r, b = win32gui.GetClientRect(hwnd)
    w = max(0, r - l)
    h = max(0, b - t)
    return w // 2, h // 2


def child_from_client_point(hwnd_parent, x, y):
    try:
        ch = win32gui.ChildWindowFromPoint(hwnd_parent, (x, y))
        if ch and win32gui.IsWindow(ch):
            return ch
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

# ---- SendMessage input ----

def _pack_lparam(x, y):
    return (y << 16) | (x & 0xFFFF)


def send_mouse_move(hwnd, x, y, wparam=0):
    win32gui.SendMessage(hwnd, win32con.WM_MOUSEMOVE, wparam, _pack_lparam(x, y))


def send_left_click(hwnd, x, y):
    lp = _pack_lparam(x, y)
    send_mouse_move(hwnd, x, y)
    win32gui.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lp)
    win32gui.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lp)


def _vk_from_key_name(name: str) -> int:
    n = name.strip().lower()
    if len(n) == 1:
        return ord(n.upper())
    mapping = {
        'space': win32con.VK_SPACE,
        'shift': win32con.VK_SHIFT,
        'ctrl': win32con.VK_CONTROL,
        'esc': win32con.VK_ESCAPE,
        'escape': win32con.VK_ESCAPE,
    }
    if n in mapping:
        return mapping[n]
    if n.isdigit():
        return ord(n)
    raise ValueError(f'Unsupported key: {name}')


def _make_key_lparam(vk: int, is_keyup: bool) -> int:
    scan = win32api.MapVirtualKey(vk, 0) & 0xFF
    lp = (1) | (scan << 16)
    if is_keyup:
        lp |= (1 << 30) | (1 << 31)
    return lp


def send_key_down(hwnd, key_name: str):
    vk = _vk_from_key_name(key_name)
    win32gui.SendMessage(hwnd, win32con.WM_KEYDOWN, vk, _make_key_lparam(vk, False))


def send_key_up(hwnd, key_name: str):
    vk = _vk_from_key_name(key_name)
    win32gui.SendMessage(hwnd, win32con.WM_KEYUP, vk, _make_key_lparam(vk, True))


def send_key_press(hwnd, key_name: str, hold: float = 0.0):
    send_key_down(hwnd, key_name)
    if hold > 0:
        time.sleep(hold)
    send_key_up(hwnd, key_name)


# ---- Player ----

def play_actions(hwnd, steps, logfn):
    ensure_restored(hwnd)
    cx, cy = get_client_center(hwnd)
    target_child = child_from_client_point(hwnd, cx, cy)
    try:
        logfn(f"目标窗口 0x{hwnd:08X} class={win32gui.GetClassName(hwnd)} 标题={win32gui.GetWindowText(hwnd)}")
        logfn(f"鼠标目标 0x{target_child:08X} class={win32gui.GetClassName(target_child)}")
    except Exception:
        pass
    # If first step is keyboard, do an activation left click first
    try:
        first = next((s for s in steps if s.get('type') in ('key','mouse')), None)
        if first and first.get('type') == 'key':
            tx, ty = map_point_parent_to_child(hwnd, target_child, cx, cy)
            logfn('[激活] 首步为键盘，先发送一次左键点击以激活后台窗口')
            send_left_click(target_child, tx, ty)
            time.sleep(0.05)
    except Exception:
        pass
    for i, st in enumerate(steps):
        kind = st.get('type')
        delay = float(st.get('delay', 0))
        try:
            if kind == 'key':
                key = st['key']
                hold = float(st.get('hold', 0))
                logfn(f"[{i+1}] key {key} hold={hold}s delay={delay}s")
                # deliver to both child and top-level for best compatibility
                send_key_press(target_child, key, hold)
                send_key_press(hwnd, key, 0)
            elif kind == 'mouse':
                btn = st.get('button', 'left').lower()
                hold = float(st.get('hold', 0))
                logfn(f"[{i+1}] mouse {btn} hold={hold}s delay={delay}s")
                if btn == 'left':
                    tx, ty = map_point_parent_to_child(hwnd, target_child, cx, cy)
                    send_left_click(target_child, tx, ty)
                    if hold > 0:
                        time.sleep(hold)
                elif btn == 'right':
                    tx, ty = map_point_parent_to_child(hwnd, target_child, cx, cy)
                    lp = _pack_lparam(tx, ty)
                    send_mouse_move(target_child, tx, ty)
                    win32gui.SendMessage(target_child, win32con.WM_RBUTTONDOWN, win32con.MK_RBUTTON, lp)
                    if hold > 0:
                        time.sleep(hold)
                    win32gui.SendMessage(target_child, win32con.WM_RBUTTONUP, 0, lp)
            else:
                logfn(f"[{i+1}] 未知动作类型: {kind}")
        except Exception as e:
            logfn(f"[{i+1}] 执行错误: {e}")
        if delay > 0:
            time.sleep(delay)


# ---- Simple GUI ----

class JsonTestApp:
    def __init__(self, root):
        self.root = root
        self.root.title('JSON 回放测试 - 50扼守')
        self.hwnd = None
        self.steps = []

        frm = ttk.Frame(root, padding=10)
        frm.pack(fill='both', expand=True)

        row1 = ttk.Frame(frm)
        row1.pack(fill='x', pady=4)
        ttk.Label(row1, text='窗口:').pack(side='left')
        self.combo = ttk.Combobox(row1, state='readonly', width=60)
        self.combo.pack(side='left', padx=6, fill='x', expand=True)
        ttk.Button(row1, text='刷新', command=self.refresh_windows).pack(side='left', padx=4)

        row2 = ttk.Frame(frm)
        row2.pack(fill='x', pady=4)
        ttk.Label(row2, text='JSON 文件:').pack(side='left')
        self.entry = ttk.Entry(row2, width=50)
        self.entry.pack(side='left', padx=6, fill='x', expand=True)
        ttk.Button(row2, text='浏览...', command=self.browse_json).pack(side='left', padx=4)

        ttk.Button(frm, text='开始回放', command=self.start).pack(anchor='w', pady=8)
        ttk.Button(frm, text='后台序列自测(与test2相同)', command=self.run_selfcheck).pack(anchor='w')

        self.log = tk.Text(frm, height=18)
        self.log.pack(fill='both', expand=True)

        self.refresh_windows()

    def _log(self, msg):
        # thread-safe log
        self.root.after(0, lambda: (self.log.insert('end', msg + '\n'), self.log.see('end')))

    def refresh_windows(self):
        wins = enumerate_windows()
        self._wins = [h for h, _, _ in wins]
        self.combo['values'] = [f"0x{h:08X} | {title} | {cls}" for h, cls, title in wins]
        if self._wins:
            self.combo.current(0)

    def browse_json(self):
        p = filedialog.askopenfilename(initialdir=os.path.dirname(__file__),
                                       filetypes=[('JSON Files','*.json')])
        if p:
            self.entry.delete(0, 'end')
            self.entry.insert(0, p)

    def start(self):
        idx = self.combo.current()
        if idx < 0:
            messagebox.showwarning('提示', '请选择窗口')
            return
        self.hwnd = self._wins[idx]
        path = self.entry.get().strip()
        if not path or not os.path.isfile(path):
            messagebox.showwarning('提示', '请选择有效的 JSON 文件')
            return
        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            self.steps = data.get('steps', [])
        except Exception as e:
            messagebox.showerror('错误', f'读取 JSON 失败: {e}')
            return
        self._log(f'加载 {os.path.basename(path)}，步骤数: {len(self.steps)}')
        threading.Thread(target=lambda: (play_actions(self.hwnd, self.steps, self._log), self._log('回放完成')),
                         daemon=True).start()

    def run_selfcheck(self):
        idx = self.combo.current()
        if idx < 0:
            messagebox.showwarning('提示', '请选择窗口')
            return
        hwnd = self._wins[idx]
        def worker():
            ensure_restored(hwnd)
            x, y = get_client_center(hwnd)
            target = child_from_client_point(hwnd, x, y)
            try:
                self._log(f"[自测] hwnd=0x{hwnd:08X}, target=0x{target:08X}")
            except Exception:
                pass
            # Step 1: Left click x3
            self._log('[自测] 左键点击 x3')
            for _ in range(3):
                send_left_click(target, x, y)
                time.sleep(0.15)
            # Step 2: Hold right mouse 2s
            self._log('[自测] 右键按住 2s')
            lp = _pack_lparam(x, y)
            send_mouse_move(target, x, y)
            win32gui.SendMessage(target, win32con.WM_RBUTTONDOWN, win32con.MK_RBUTTON, lp)
            time.sleep(2.0)
            win32gui.SendMessage(target, win32con.WM_RBUTTONUP, 0, lp)
            # Step 3: Hold W 2s
            self._log('[自测] 按住 W 2s')
            send_key_down(target, 'w')
            time.sleep(2.0)
            send_key_up(target, 'w')
            # Step 4: Tap Space x2
            self._log('[自测] 空格 x2')
            send_key_press(target, 'space', 0)
            time.sleep(0.1)
            send_key_press(target, 'space', 0)
            self._log('[自测] 完成')
        threading.Thread(target=worker, daemon=True).start()


if __name__ == '__main__':
    root = tk.Tk()
    app = JsonTestApp(root)
    root.geometry('820x520')
    root.mainloop()
