import os
import sys
import json
import time
import threading
from datetime import datetime

import win32gui
import tkinter as tk
from tkinter import ttk, messagebox


def enumerate_windows():
    wins = []
    def _enum_cb(h, _):
        if win32gui.IsWindowVisible(h) and win32gui.GetWindowText(h):
            wins.append((h, win32gui.GetClassName(h), win32gui.GetWindowText(h)))
    win32gui.EnumWindows(_enum_cb, None)
    return wins


ALLOWED_KEYS = set(list('wasdf') + [str(i) for i in range(10)] + ['space', 'shift', 'ctrl', 'esc'])


def normalize_key(name: str, scan_code: int = None):
    # prefer name, fallback to scan_code
    if name:
        n = name.lower()
    else:
        n = ''
    # unify ctrl
    if n in ('control', 'left ctrl', 'right ctrl', 'ctrl'):
        return 'ctrl'
    # unify shift
    if n in ('shift', 'left shift', 'right shift'):
        return 'shift'
    # unify esc
    if n in ('escape', 'esc'):
        return 'esc'
    # unify numpad like 'num 1'
    if n.startswith('num '):
        d = n[4:].strip()
        if d.isdigit() and len(d) == 1:
            return d
    # accept letters/digits and space
    if n in ('w', 'a', 's', 'd', 'f', 'space'):
        return n
    if n.isdigit() and len(n) == 1:
        return n
    # fallback by scan code mapping (keyboard module on Windows)
    sc_map = {
        30: 'a', 17: 'w', 31: 's', 32: 'd', 33: 'f',
        57: 'space',
        42: 'shift', 54: 'shift',
        29: 'ctrl', 3613: 'ctrl',
        1: 'esc',
        # top row digits 1..0
        2: '1', 3: '2', 4: '3', 5: '4', 6: '5', 7: '6', 8: '7', 9: '8', 10: '9', 11: '0',
        # numpad digits
        82: '0', 79: '1', 80: '2', 81: '3', 75: '4', 76: '5', 77: '6', 71: '7', 72: '8', 73: '9',
    }
    if scan_code in sc_map:
        return sc_map[scan_code]
    return None


class RecorderApp:
    def __init__(self, root):
        self.root = root
        self.root.title('操作录制器 - 50扼守')
        self.base_dir = self._compute_base_dir()
        self.map_dir = os.path.join(self.base_dir, 'map')
        self.json_dir = os.path.join(self.base_dir, 'json')

        self.selected_hwnd = None
        self.selected_map = tk.StringVar(value='mapA')

        self.is_recording = False
        self.is_cancelling = False
        self.records = []
        self.last_event_time = None

        self._build_ui()
        self.refresh_windows()
        self.refresh_maps()
        self._bind_hotkeys()

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

    def _build_ui(self):
        frame = ttk.Frame(self.root, padding=10)
        frame.pack(fill='both', expand=True)

        row1 = ttk.Frame(frame)
        row1.pack(fill='x', pady=4)
        ttk.Label(row1, text='窗口:').pack(side='left')
        self.combo_win = ttk.Combobox(row1, state='readonly', width=60)
        self.combo_win.pack(side='left', padx=6, fill='x', expand=True)
        self.combo_win.bind('<<ComboboxSelected>>', self.on_sel_window)
        ttk.Button(row1, text='刷新窗口', command=self.refresh_windows).pack(side='left', padx=4)

        row2 = ttk.Frame(frame)
        row2.pack(fill='x', pady=4)
        ttk.Label(row2, text='地图模板:').pack(side='left')
        self.combo_map = ttk.Combobox(row2, state='readonly', width=20, textvariable=self.selected_map)
        self.combo_map.pack(side='left', padx=6)
        ttk.Button(row2, text='刷新地图', command=self.refresh_maps).pack(side='left', padx=4)

        row3 = ttk.Frame(frame)
        row3.pack(fill='x', pady=8)
        ttk.Button(row3, text='开始录制 (F9)', command=self.start_record).pack(side='left', padx=6)
        ttk.Button(row3, text='结束并保存 (F10)', command=self.stop_and_save).pack(side='left', padx=6)
        ttk.Button(row3, text='放弃本次 (F12)', command=self.cancel_record).pack(side='left', padx=6)

        # Option: include mouse recording
        self.include_mouse = tk.BooleanVar(value=False)
        ttk.Checkbutton(frame, text='是否包含鼠标操作录制', variable=self.include_mouse).pack(anchor='w', pady=(0,6))

        # Option: debug key capture (log raw names)
        self.debug_keys = tk.BooleanVar(value=False)
        ttk.Checkbutton(frame, text='调试键盘捕获(输出原始键名)', variable=self.debug_keys).pack(anchor='w')

        self.log = tk.Text(frame, height=16)
        self.log.pack(fill='both', expand=True)

    def _log(self, msg):
        ts = datetime.now().strftime('%H:%M:%S')
        self.log.insert('end', f'[{ts}] {msg}\n')
        self.log.see('end')

    def _bind_hotkeys(self):
        try:
            import keyboard  # type: ignore
            keyboard.add_hotkey('f9', self.start_record)
            keyboard.add_hotkey('f10', self.stop_and_save)
            keyboard.add_hotkey('f12', self.cancel_record)
        except Exception:
            pass

    def refresh_windows(self):
        wins = enumerate_windows()
        items = [f"0x{h:08X} | {title} | {cls}" for h, cls, title in wins]
        self._wins_ref = [h for h, _, _ in wins]
        self.combo_win['values'] = items
        if items:
            self.combo_win.current(0)
            self.on_sel_window()
        self._log('已刷新窗口列表')

    def on_sel_window(self, _evt=None):
        idx = self.combo_win.current()
        if idx >= 0:
            self.selected_hwnd = self._wins_ref[idx]

    def refresh_maps(self):
        names = []
        try:
            for fn in os.listdir(self.map_dir):
                if fn.lower().endswith('.png'):
                    names.append(os.path.splitext(fn)[0])
            names.sort()
            self.combo_map['values'] = names
            if names and (self.selected_map.get() not in names):
                self.combo_map.current(0)
            self._log(f'已刷新地图模板 (路径: {self.map_dir}, 数量: {len(names)})')
        except FileNotFoundError:
            self.combo_map['values'] = []
            self._log(f'地图目录不存在: {self.map_dir}，请将 map 文件夹放在与 exe 同目录')
        except Exception as e:
            self._log(f'读取地图目录失败: {e}')

    # Recording
    def start_record(self):
        if self.is_recording:
            return
        if not self.selected_hwnd:
            messagebox.showwarning('提示', '请选择目标窗口')
            return
        self.records = []
        self.last_event_time = None
        self.is_cancelling = False
        self.is_recording = True
        self._log('开始录制 (F10 结束保存，F12 放弃)')
        self._start_hooks()

    def stop_and_save(self):
        if not self.is_recording:
            return
        self.is_recording = False
        self._stop_hooks()
        if self.is_cancelling:
            self._log('已放弃本次录制')
            return
        map_name = self.selected_map.get().strip()
        if not map_name:
            messagebox.showwarning('提示', '请选择地图模板')
            return
        os.makedirs(self.json_dir, exist_ok=True)
        out_path = os.path.join(self.json_dir, f'{map_name}.json')
        data = {
            'name': map_name,
            'steps': self.records,
        }
        try:
            with open(out_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            self._log(f'已保存到 {out_path} (共 {len(self.records)} 步)')
        except Exception as e:
            self._log(f'保存失败: {e}')

    def cancel_record(self):
        if not self.is_recording:
            return
        self.is_cancelling = True
        self.is_recording = False
        self._stop_hooks()
        self.records = []
        self._log('已取消本次录制')

    def _start_hooks(self):
        # Use keyboard & mouse global hooks
        try:
            import keyboard  # type: ignore
            import mouse  # type: ignore

            # Keyboard: use unified hook for down/up
            self._kb_handle = keyboard.hook(self._on_kb_event, suppress=False)

            # Mouse global hook (optional)
            self._mouse_hooked = False
            if self.include_mouse.get():
                mouse.hook(self._on_mouse_event)
                self._mouse_hooked = True
            self._log('键鼠全局钩子已启动')
        except Exception as e:
            self._log(f'启动钩子失败: {e}')
            import keyboard  # type: ignore
            import mouse  # type: ignore
            if getattr(self, '_kb_down_handle', None):
                keyboard.unhook(self._kb_down_handle)
                self._kb_down_handle = None

    def _stop_hooks(self):
        try:
            import keyboard  # type: ignore
            import mouse  # type: ignore
            if getattr(self, '_kb_handle', None):
                keyboard.unhook(self._kb_handle)
                self._kb_handle = None
            if self._mouse_hooked:
                mouse.unhook(self._on_mouse_event)
            self._log('已解除键鼠钩子')
        except Exception as e:
            self._log(f'解除钩子时出错: {e}')

    def _append_step(self, step):
        now = time.time()
        if self.last_event_time is None:
            delay = 0.0
        else:
            delay = max(0.0, now - self.last_event_time)
        self.last_event_time = now
        step['delay'] = round(delay, 3)
        self.records.append(step)
        self._log(f"记录: {step}")

    def _on_kb_event(self, e):
        if not self.is_recording:
            return
        try:
            raw = (e.name or '')
            etype = getattr(e, 'event_type', '')
            rawl = raw.lower()
            if self.debug_keys.get():
                self._log(f"[DBG] key {etype}: '{raw}' sc={getattr(e, 'scan_code', None)}")
            # ignore our hotkeys & win keys
            if rawl in ('left windows', 'right windows', 'f9', 'f10', 'f12'):
                return
            name = normalize_key(raw, getattr(e, 'scan_code', None))
            if name is None or name not in ALLOWED_KEYS:
                return
            if etype == 'down':
                if not hasattr(self, '_key_down_at'):
                    self._key_down_at = {}
                if name not in getattr(self, '_key_down_at', {}):
                    self._key_down_at[name] = time.time()
            elif etype == 'up':
                start = None
                if hasattr(self, '_key_down_at'):
                    start = self._key_down_at.pop(name, None)
                hold = 0.0 if start is None else max(0.0, time.time() - start)
                self._append_step({'type': 'key', 'key': name, 'hold': round(hold, 3)})
        except Exception:
            pass

    def _on_mouse_event(self, e):
        if not self.is_recording:
            return
        try:
            # mouse library provides attributes: event_type in ('down','up') and button in ('left','right','middle')
            if hasattr(e, 'button') and hasattr(e, 'event_type'):
                btn = 'left' if str(getattr(e, 'button')).lower() == 'left' else 'right'
                etype = str(getattr(e, 'event_type')).lower()
                if etype == 'down':
                    if not hasattr(self, '_mouse_down_at'):
                        self._mouse_down_at = {}
                    self._mouse_down_at[btn] = time.time()
                elif etype == 'up':
                    start = None
                    if hasattr(self, '_mouse_down_at'):
                        start = self._mouse_down_at.pop(btn, None)
                    hold = 0.0 if start is None else max(0.0, time.time() - start)
                    self._append_step({'type': 'mouse', 'button': btn, 'hold': round(hold, 3)})
        except Exception:
            pass


def main():
    root = tk.Tk()
    app = RecorderApp(root)
    root.geometry('780x460')
    root.mainloop()


if __name__ == '__main__':
    main()

