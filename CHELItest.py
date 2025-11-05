import os
import sys
import time
import argparse
import cv2
import numpy as np
import win32gui
import win32con
import ctypes

# Reuse background capture and background input utilities
from test import BackgroundScreenshot
from test2 import (
    choose_window,
    ensure_restored,
)

CONTROL_DIR = os.path.join(os.path.dirname(__file__), 'control')


def _pack_lparam(x, y):
    return (int(y) << 16) | (int(x) & 0xFFFF)


def _makelong(low, high):
    return ((int(high) & 0xFFFF) << 16) | (int(low) & 0xFFFF)


def find_deepest_child_at_screen_point(root_hwnd, sx, sy):
    """Return (deepest_hwnd, chain, child_client_xy)
    chain: list of (hwnd, class, client_x, client_y) from parent to deepest.
    """
    chain = []
    current = root_hwnd
    flags = 0x0001 | 0x0002 | 0x0004  # CWP_SKIPINVISIBLE | CWP_SKIPDISABLED | CWP_SKIPTRANSPARENT
    while True:
        try:
            cx, cy = win32gui.ScreenToClient(current, (sx, sy))
        except Exception:
            break
        try:
            cls = win32gui.GetClassName(current)
        except Exception:
            cls = ''
        chain.append((current, cls, cx, cy))
        try:
            child = win32gui.ChildWindowFromPointEx(current, (cx, cy), flags)
        except Exception:
            child = None
        if not child or child == current or not win32gui.IsWindow(child):
            break
        current = child
    # deepest is last in chain
    if not chain:
        return root_hwnd, [], (0, 0)
    d_hwnd, d_cls, d_cx, d_cy = chain[-1]
    return d_hwnd, chain, (d_cx, d_cy)


def match_template_multiscale(bgr_img, template_path, threshold=0.80, scales=(1.1, 1.05, 1.0, 0.95, 0.9)):
    if bgr_img is None:
        return None
    if not os.path.isfile(template_path):
        print(f"[Match] 模板不存在: {template_path}")
        return None
    tpl = cv2.imread(template_path, cv2.IMREAD_COLOR)
    if tpl is None:
        print(f"[Match] 模板读取失败: {template_path}")
        return None
    best = None
    for s in scales:
        try:
            if abs(s - 1.0) > 1e-6:
                t = cv2.resize(tpl, None, fx=s, fy=s, interpolation=cv2.INTER_AREA if s < 1.0 else cv2.INTER_CUBIC)
            else:
                t = tpl
            res = cv2.matchTemplate(bgr_img, t, cv2.TM_CCOEFF_NORMED)
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
            if best is None or max_val > best[0]:
                h, w = t.shape[:2]
                x, y = max_loc
                center = (x + w // 2, y + h // 2)
                best = (max_val, (x, y, w, h), center, s)
        except Exception as e:
            print(f"[Match] scale={s} 异常: {e}")
    if best and best[0] >= threshold:
        score, rect, center, scale = best
        print(f"[Match] 命中: score={score:.3f} scale={scale} rect={rect} center={center}")
        return { 'score': float(score), 'rect': rect, 'center': center, 'scale': float(scale) }
    if best:
        print(f"[Match] 未达阈值: max={best[0]:.3f} thr={threshold}")
    else:
        print("[Match] 未找到匹配")
    return None


def run_once(hwnd, target_key, threshold=0.80, scales=(1.1, 1.05, 1.0, 0.95, 0.9), seq_delay_ms=60):
    ensure_restored(hwnd)
    # Map target key to control template path
    name_map = {
        'cheli': os.path.join(CONTROL_DIR, 'cheli.png'),
        'jixutiaozhan': os.path.join(CONTROL_DIR, 'jixutiaozhan.png'),
    }
    tpl_path = name_map.get(target_key)
    if not tpl_path:
        print(f"未知目标: {target_key}")
        return False
    print(f"[Info] 目标模板: {tpl_path} 存在={os.path.isfile(tpl_path)}")

    capturer = BackgroundScreenshot(hwnd)
    img = capturer.capture_background()
    if img is None:
        print("[Error] 截图失败(None)")
        return False
    h, w = img.shape[:2]
    print(f"[Info] 截图尺寸: {w}x{h}")

    print(f"[Match] 参数: thr={threshold:.3f} scales={list(scales)}")
    m = match_template_multiscale(img, tpl_path, threshold=threshold, scales=scales)
    if not m:
        return False

    # Convert template center (image coords relative to window) to client coords for SendMessage
    cx_img, cy_img = m['center']
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    sx, sy = left + cx_img, top + cy_img  # screen coords
    client_pt = win32gui.ScreenToClient(hwnd, (sx, sy))
    tx, ty = client_pt

    # Resolve DEEPEST child window at the point and convert to its client coords
    child, chain, (csx, csy) = find_deepest_child_at_screen_point(hwnd, sx, sy)
    chain_str = ' -> '.join([f"0x{h:08X}:{c or ''}@({x},{y})" for (h,c,x,y) in chain])
    print(f"[Click] chain: {chain_str}")
    print(f"[Click] parent=0x{hwnd:08X} deepest_child=0x{child:08X} center_img=({cx_img},{cy_img}) parent_client=({tx},{ty}) child_client=({csx},{csy}) screen=({sx},{sy})")
    # Strict sequence only: Activate -> MapVirtualKey -> MOVE -> DOWN -> delay -> UP
    try:
        # Activate parent window first
        try:
            win32gui.SendMessage(hwnd, win32con.WM_ACTIVATE, win32con.WA_ACTIVE, 0)
            time.sleep(0.03)
        except Exception:
            pass
        # Parent expects WM_MOUSEACTIVATE prior to click; simulate HTCLIENT/LBUTTONDOWN
        try:
            lparam_ma = _makelong(win32con.WM_LBUTTONDOWN, win32con.HTCLIENT)
            win32gui.SendMessage(hwnd, win32con.WM_MOUSEACTIVATE, hwnd, lparam_ma)
            time.sleep(0.01)
        except Exception:
            pass
        # Child may expect cursor/hover setup and focus
        try:
            lparam_cursor = _makelong(win32con.HTCLIENT, win32con.WM_MOUSEMOVE)
            win32gui.SendMessage(child, win32con.WM_SETCURSOR, child, lparam_cursor)
            win32gui.SendMessage(child, win32con.WM_SETFOCUS, 0, 0)
            time.sleep(0.01)
        except Exception:
            pass
        try:
            scan = ctypes.windll.user32.MapVirtualKeyW(win32con.VK_LBUTTON, 0)
        except Exception:
            scan = 0
        print(f"[Seq] VK_LBUTTON scan={scan}")
        # Per MSDN, WM_* mouse messages (except wheel) expect client coords in lParam (use deepest child)
        lparam_child  = _pack_lparam(csx, csy)
        # MOVE
        win32gui.SendMessage(child,  win32con.WM_MOUSEMOVE, 0, lparam_child)
        # DOWN
        win32gui.SendMessage(child,  win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam_child)
        time.sleep(max(0.0, float(seq_delay_ms) / 1000.0))
        # UP
        win32gui.SendMessage(child,  win32con.WM_LBUTTONUP, 0, lparam_child)
        print("[Seq] 完成严格顺序点击")
        return True
    except Exception as e:
        print(f"[Seq] 异常: {e}")
        return False


def main():
    parser = argparse.ArgumentParser(description='CHELI/JIXUTIAOZHAN 严格顺序点击测试')
    parser.add_argument('--target', choices=['cheli', 'jixutiaozhan'], default=None, help='测试目标按钮')
    parser.add_argument('--thr', type=float, default=0.80, help='匹配阈值（默认0.80）')
    parser.add_argument('--scales', type=str, default='1.1,1.05,1.0,0.95,0.9', help='匹配尺度，逗号分隔，例如 1.2,1.1,1.0,0.9')
    parser.add_argument('--seq-delay-ms', type=int, default=60, help='按下后的延迟毫秒（再抬起）')
    args = parser.parse_args()

    print('== 选择目标窗口 ==')
    hwnd = choose_window()
    if not hwnd or not win32gui.IsWindow(hwnd):
        print('[Error] 无效窗口')
        sys.exit(1)

    target = args.target
    if not target:
        print('选择测试目标:')
        print('  [1] cheli (撤离)')
        print('  [2] jixutiaozhan (继续挑战)')
        sel = input('请输入序号: ').strip()
        target = 'cheli' if sel == '1' else 'jixutiaozhan'

    # parse scales
    try:
        scales = tuple(float(s.strip()) for s in args.scales.split(',') if s.strip())
    except Exception:
        scales = (1.1, 1.05, 1.0, 0.95, 0.9)
    ok = run_once(hwnd, target, threshold=float(args.thr), scales=scales, seq_delay_ms=args.seq_delay_ms)
    sys.exit(0 if ok else 2)


if __name__ == '__main__':
    main()
