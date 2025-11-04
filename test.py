import ctypes
import sys
import time
import argparse
import win32gui
import win32ui
import win32con
import win32api
import numpy as np
import cv2


class BackgroundScreenshot:
    def __init__(self, hwnd=None):
        self.hwnd = hwnd

    def set_hwnd(self, hwnd):
        self.hwnd = hwnd

    def capture_background(self):
        """返回窗口图像（BGR np.ndarray）或 None"""
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

            # 2 = PW_RENDERFULLCONTENT, 能在部分窗口（UWP/无效）回退
            res = ctypes.windll.user32.PrintWindow(self.hwnd, saveDC.GetSafeHdc(), 2)
            if res != 1:
                # 回退 BitBlt（有时只能拿到被遮挡后的内容）
                saveDC.BitBlt((0, 0), (width, height), mfcDC, (0, 0), win32con.SRCCOPY)

            bmpinfo = saveBitMap.GetInfo()
            bmpstr = saveBitMap.GetBitmapBits(True)
            img = np.frombuffer(bmpstr, dtype=np.uint8)
            img.shape = (bmpinfo['bmHeight'], bmpinfo['bmWidth'], 4)
            img = img[:, :, :3]  # BGRA -> BGR

            win32gui.DeleteObject(saveBitMap.GetHandle())
            saveDC.DeleteDC()
            mfcDC.DeleteDC()
            win32gui.ReleaseDC(self.hwnd, hwndDC)
            return img
        except Exception:
            try:
                # 清理（防止句柄泄漏）
                win32gui.DeleteObject(saveBitMap.GetHandle())
                saveDC.DeleteDC()
                mfcDC.DeleteDC()
                win32gui.ReleaseDC(self.hwnd, hwndDC)
            except Exception:
                pass
            return None


def enum_visible_windows():
    windows = []
    def _enum(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if title and title.strip():
                windows.append((title, hwnd))
    win32gui.EnumWindows(_enum, None)
    return windows


def pick_window_interactive():
    windows = enum_visible_windows()
    if not windows:
        print("未找到可见窗口")
        return None
    print("可见窗口列表：")
    for i, (title, hwnd) in enumerate(windows):
        print(f"[{i}] {title} (HWND=0x{hwnd:08X})")
    while True:
        try:
            idx = input("请输入要截图的窗口序号并回车：")
            if not idx:
                return None
            idx = int(idx)
            if 0 <= idx < len(windows):
                return windows[idx][1]
        except Exception:
            pass
        print("输入无效，请重试。")


def main():
    parser = argparse.ArgumentParser(description="Test BackgroundScreenshot via PrintWindow")
    parser.add_argument("--hwnd", type=lambda x: int(x, 0), default=None, help="窗口句柄（十进制或0x十六进制）")
    parser.add_argument("--loop", action="store_true", help="循环捕获并显示")
    parser.add_argument("--interval", type=float, default=1.0, help="循环间隔（秒）")
    args = parser.parse_args()

    if args.hwnd is None:
        hwnd = pick_window_interactive()
    else:
        hwnd = args.hwnd

    if not hwnd:
        print("未选择窗口，退出。")
        sys.exit(1)

    if not win32gui.IsWindow(hwnd):
        print(f"无效窗口句柄: {hwnd}")
        sys.exit(1)

    title = win32gui.GetWindowText(hwnd)
    print(f"目标窗口: {title} (HWND=0x{hwnd:08X})")

    capturer = BackgroundScreenshot(hwnd)

    def do_capture_once(seq=None):
        img = capturer.capture_background()
        if img is None:
            print("截图失败（返回None）")
            return False
        h, w = img.shape[:2]
        print(f"截图成功: {w}x{h}")
        out = f"test_capture{'_{:04d}'.format(seq) if seq is not None else ''}.png"
        cv2.imwrite(out, img)
        cv2.imshow("Background Capture (BGR)", img)
        cv2.waitKey(1)
        print(f"已保存: {out}")
        return True

    if args.loop:
        i = 0
        try:
            while True:
                ok = do_capture_once(seq=i)
                i += 1
                time.sleep(max(0.01, args.interval))
        except KeyboardInterrupt:
            print("已停止循环。")
        finally:
            cv2.destroyAllWindows()
    else:
        ok = do_capture_once()
        if not ok:
            sys.exit(2)
        print("按任意键关闭窗口...")
        cv2.waitKey(0)
        cv2.destroyAllWindows()


if __name__ == "__main__":
    main()
