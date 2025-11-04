import time
import sys
import ctypes
import win32con
import win32gui
import win32api
import win32process

# Utilities
LOWORD = lambda dword: dword & 0xFFFF
HIWORD = lambda dword: (dword >> 16) & 0xFFFF

user32 = ctypes.WinDLL('user32', use_last_error=True)


def pack_lparam(x, y):
    return (y << 16) | (x & 0xFFFF)


def get_client_center(hwnd):
    # Get client rect size and center point in client coordinates
    left, top, right, bottom = win32gui.GetClientRect(hwnd)
    width = max(0, right - left)
    height = max(0, bottom - top)
    cx = width // 2
    cy = height // 2
    return cx, cy


def child_from_client_point(hwnd_parent, x, y):
    try:
        child = win32gui.ChildWindowFromPoint(hwnd_parent, (x, y))
        if child and win32gui.IsWindow(child):
            return child
    except Exception:
        pass
    return hwnd_parent


def ensure_restored(hwnd):
    # If minimized, try to restore (some apps ignore background messages while minimized)
    try:
        placement = win32gui.GetWindowPlacement(hwnd)
        show_cmd = placement[1]
        if show_cmd == win32con.SW_SHOWMINIMIZED:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            time.sleep(0.2)
    except Exception:
        pass


# Keyboard helpers
VK_SPACE = win32con.VK_SPACE


def vk_from_key_name(name: str) -> int:
    name = name.strip().lower()
    if len(name) == 1:
        ch = name.upper()
        return ord(ch)
    if name == 'space':
        return VK_SPACE
    # Extend here for other names if needed
    raise ValueError(f"Unsupported key name: {name}")


def make_key_lparam(vk: int, is_keyup: bool) -> int:
    scan = win32api.MapVirtualKey(vk, 0) & 0xFF
    lparam = (1) | (scan << 16)
    if is_keyup:
        lparam |= (1 << 30) | (1 << 31)
    return lparam


# Mouse actions
def send_mouse_move(hwnd, x, y, wparam=0):
    lparam = pack_lparam(x, y)
    win32gui.SendMessage(hwnd, win32con.WM_MOUSEMOVE, wparam, lparam)


def send_left_click(hwnd, x, y):
    lparam = pack_lparam(x, y)
    send_mouse_move(hwnd, x, y)
    win32gui.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
    win32gui.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)


def send_right_down(hwnd, x, y):
    lparam = pack_lparam(x, y)
    send_mouse_move(hwnd, x, y)
    win32gui.SendMessage(hwnd, win32con.WM_RBUTTONDOWN, win32con.MK_RBUTTON, lparam)


def send_right_up(hwnd, x, y):
    lparam = pack_lparam(x, y)
    win32gui.SendMessage(hwnd, win32con.WM_RBUTTONUP, 0, lparam)


# Keyboard actions
def send_key_down(hwnd, key_name: str):
    vk = vk_from_key_name(key_name)
    lparam = make_key_lparam(vk, is_keyup=False)
    win32gui.SendMessage(hwnd, win32con.WM_KEYDOWN, vk, lparam)


def send_key_up(hwnd, key_name: str):
    vk = vk_from_key_name(key_name)
    lparam = make_key_lparam(vk, is_keyup=True)
    win32gui.SendMessage(hwnd, win32con.WM_KEYUP, vk, lparam)


def send_key_press(hwnd, key_name: str, hold_seconds: float = 0.0):
    send_key_down(hwnd, key_name)
    if hold_seconds > 0:
        time.sleep(hold_seconds)
    send_key_up(hwnd, key_name)


# Window enumeration and selection

def list_top_windows():
    wins = []
    def _enum_cb(h, _):
        if win32gui.IsWindowVisible(h) and win32gui.GetWindowText(h):
            cls = win32gui.GetClassName(h)
            title = win32gui.GetWindowText(h)
            wins.append((h, cls, title))
    win32gui.EnumWindows(_enum_cb, None)
    return wins


def choose_window():
    wins = list_top_windows()
    print('== Select Target Window ==')
    for i, (_, cls, title) in enumerate(wins):
        print(f"[{i}] {title}  (class={cls})")
    while True:
        try:
            idx = int(input('Enter index: ').strip())
            if 0 <= idx < len(wins):
                return wins[idx][0]
        except Exception:
            pass
        print('Invalid index, try again.')


# Test sequence

def run_test_sequence(hwnd):
    ensure_restored(hwnd)

    # Use center of client area; optionally hit-test a child window
    x, y = get_client_center(hwnd)
    target = child_from_client_point(hwnd, x, y)

    # Log handles for troubleshooting
    try:
        print(f"Parent hwnd=0x{hwnd:08X}, target hwnd=0x{target:08X}")
    except Exception:
        print(f"Parent hwnd={hwnd}, target hwnd={target}")

    # Left click 3 times
    print('Step 1: Left click x3')
    for _ in range(3):
        send_left_click(target, x, y)
        time.sleep(0.15)

    # Hold right mouse 2s
    print('Step 2: Hold right mouse 2s')
    send_right_down(target, x, y)
    time.sleep(2.0)
    send_right_up(target, x, y)

    # Hold W for 2s
    print('Step 3: Hold W 2s')
    send_key_press(target, 'w', hold_seconds=2.0)

    # Tap Space twice
    print('Step 4: Tap Space x2')
    send_key_press(target, 'space', hold_seconds=0)
    time.sleep(0.1)
    send_key_press(target, 'space', hold_seconds=0)

    print('Done.')


if __name__ == '__main__':
    print('Background Input Test (SendMessage)')
    hwnd = choose_window()
    run_test_sequence(hwnd)
