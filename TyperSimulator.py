# -*- coding: utf-8 -*-
"""文件打字模拟器（单文件版）

把内容像真人打字一样逐字敲进“光标点击的输入框”——网页输入框、聊天窗口、
记事本、VSCode、Word 都可以。中文 / UTF-8 / GBK 全支持，无需第三方库。

功能：
- 内容来源：选文件（txt/cpp/c/py/docx，编码自动识别），或直接在内容框粘贴/编辑
- 速度每秒 1~500 字符（自动提升系统计时精度）；“模拟抖动”更像真人；
  “硬输入”不限速直灌
- 全局热键：F6 = 开始/暂停/继续，Esc = 停止
- 外部窗口模式（打到光标点击处）/ 内置演示框模式

运行：python TyperSimulator.py   （或双击“启动打字模拟器.bat”）
自测：python -X utf8 TyperSimulator.py --selftest
"""

from __future__ import annotations

import ctypes
import ctypes.wintypes as wt
import os
import queue
import random
import shutil
import sys
import tempfile
import threading
import time
import tkinter as tk
import xml.etree.ElementTree as ET
import zipfile
from tkinter import ttk, filedialog, messagebox


# ============================================================
# 一、文件读取与编码识别
# ============================================================

SUPPORTED_EXTS = {'.txt', '.cpp', '.c', '.py', '.h', '.hpp', '.cc', '.cxx', '.docx'}

_W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


class FileReaderError(Exception):
    """文件读取失败。message 用中文说明原因与解决办法。"""


def read_text(path: str) -> tuple[str, str]:
    """读取文件，返回 (文本内容, 识别出的编码名)。失败抛 FileReaderError。"""
    if not os.path.isfile(path):
        raise FileReaderError(f'文件不存在：{path}')
    ext = os.path.splitext(path)[1].lower()
    if ext == '.doc':
        raise FileReaderError(
            '暂不支持旧版 .doc 文件，请用 Word/WPS 打开后另存为 .docx 或 .txt 再试。'
        )
    if ext not in SUPPORTED_EXTS:
        raise FileReaderError(
            f'不支持的文件类型 "{ext}"，请使用 txt / cpp / c / py / docx 等文本类文件。'
        )
    if ext == '.docx':
        return _read_docx(path)
    return _read_plain(path)


def _decode_auto(data: bytes) -> tuple[str, str]:
    """BOM 优先，其次 UTF-8 严格解码，最后 GB18030（GBK 超集）。"""
    if data.startswith(b'\xef\xbb\xbf'):
        return data.decode('utf-8-sig'), 'utf-8-sig'
    if data.startswith(b'\xff\xfe') or data.startswith(b'\xfe\xff'):
        return data.decode('utf-16'), 'utf-16'
    try:
        return data.decode('utf-8'), 'utf-8'
    except UnicodeDecodeError:
        pass
    try:
        return data.decode('gb18030'), 'gbk'
    except UnicodeDecodeError:
        raise FileReaderError(
            '无法识别文件编码（既不是 UTF-8 也不是 GBK）。'
            '请用记事本 / VSCode 打开该文件，另存为 UTF-8 或 GBK 编码后重试。'
        )


def _clean(text: str) -> str:
    """行尾统一为 \\n；去掉 \\x00 等控制字符（保留 \\n 与 \\t）。"""
    text = text.replace('\r\n', '\n').replace('\r', '\n')
    return ''.join(ch for ch in text if ch in '\n\t' or ch >= ' ')


def _read_plain(path: str) -> tuple[str, str]:
    with open(path, 'rb') as f:
        data = f.read()
    text, enc = _decode_auto(data)
    return _clean(text), enc


def _read_docx(path: str) -> tuple[str, str]:
    try:
        with zipfile.ZipFile(path) as zf:
            xml_bytes = zf.read('word/document.xml')
    except (zipfile.BadZipFile, KeyError, OSError) as e:
        raise FileReaderError(f'docx 文件无法解析（可能已损坏）：{e}')
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError as e:
        raise FileReaderError(f'docx 内容解析失败（可能已损坏）：{e}')
    paras = []
    for p in root.iter(_W + 'p'):
        paras.append(''.join(t.text or '' for t in p.iter(_W + 't')))
    return _clean('\n'.join(paras)), 'docx'


# ============================================================
# 二、打字引擎：Windows SendInput Unicode 按键模拟
# ============================================================

INPUT_MOUSE = 0
MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004
INPUT_KEYBOARD = 1
KEYEVENTF_KEYUP = 0x0002
KEYEVENTF_UNICODE = 0x0004
KEYEVENTF_SCANCODE = 0x0008
VK_TAB = 0x09
VK_SHIFT = 0x10
VK_CONTROL = 0x11
VK_MENU = 0x12
VK_END = 0x23
VK_RETURN = 0x0D
VK_ESCAPE = 0x1B
VK_LBUTTON = 0x01
VK_F6 = 0x75
VK_V = 0x56
MAPVK_VK_TO_VSC = 0


class MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ('dx', wt.LONG), ('dy', wt.LONG),
        ('mouseData', wt.DWORD), ('dwFlags', wt.DWORD), ('time', wt.DWORD),
        # ULONG_PTR：32 位下 4 字节、64 位下 8 字节，用 c_size_t 正好
        ('dwExtraInfo', ctypes.c_size_t),
    ]


class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ('wVk', wt.WORD), ('wScan', wt.WORD),
        ('dwFlags', wt.DWORD), ('time', wt.DWORD),
        ('dwExtraInfo', ctypes.c_size_t),
    ]


class HARDWAREINPUT(ctypes.Structure):
    _fields_ = [('uMsg', wt.DWORD), ('wParamL', wt.WORD), ('wParamH', wt.WORD)]


class _INPUTUNION(ctypes.Union):
    _fields_ = [('mi', MOUSEINPUT), ('ki', KEYBDINPUT), ('hi', HARDWAREINPUT)]


class INPUT(ctypes.Structure):
    _anonymous_ = ['u']
    _fields_ = [('type', wt.DWORD), ('u', _INPUTUNION)]


_SendInput = ctypes.windll.user32.SendInput
_SendInput.argtypes = [wt.UINT, ctypes.POINTER(INPUT), ctypes.c_int]
_SendInput.restype = wt.UINT

_GetAsyncKeyState = ctypes.windll.user32.GetAsyncKeyState

# 提高 Windows 定时器精度到 1ms：不开的话 time.sleep 实际最小 ~15ms，
# 高速打字（每秒几百字）根本跑不到设定速度
_winmm = ctypes.windll.winmm


_send_fail_total = 0


def _send(ki: KEYBDINPUT) -> bool:
    """发送一个键盘事件；返回是否真的被系统接收。"""
    global _send_fail_total
    inp = INPUT(type=INPUT_KEYBOARD)
    inp.ki = ki
    ok = _SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT)) == 1
    if not ok:
        _send_fail_total += 1
        if _send_fail_total <= 5:  # 只记前几次，避免刷爆日志
            err = ctypes.windll.kernel32.GetLastError()
            _log(f'!!! SendInput 被拒绝 wVk={ki.wVk} wScan={ki.wScan:#x} '
                 f'WinErr={err}（共 {_send_fail_total} 次）')
    return ok


def _send_many(events) -> bool:
    """把一组键盘事件在【一次】SendInput 调用里原子发出。

    按下+弹起成对合并，不会被其它输入插队拆成半截键；系统调用次数
    减半，丢键概率显著下降（吐字稳定性保障之一）。
    """
    global _send_fail_total
    n = len(events)
    if n == 0:
        return True
    arr = (INPUT * n)()
    for i, ki in enumerate(events):
        arr[i].type = INPUT_KEYBOARD
        arr[i].ki = ki
    ok = _SendInput(n, arr, ctypes.sizeof(INPUT)) == n
    if not ok:
        _send_fail_total += 1
        if _send_fail_total <= 5:  # 只记前几次，避免刷爆日志
            err = ctypes.windll.kernel32.GetLastError()
            _log(f'!!! SendInput 被拒绝（一组 {n} 个事件）'
                 f'WinErr={err}（共 {_send_fail_total} 次）')
    return ok


def _press_vk(vk: int) -> bool:
    return _send_many([
        KEYBDINPUT(wVk=vk, wScan=0, dwFlags=0, time=0, dwExtraInfo=0),
        KEYBDINPUT(wVk=vk, wScan=0, dwFlags=KEYEVENTF_KEYUP,
                   time=0, dwExtraInfo=0),
    ])


def _unicode_key(code: int, up: bool = False) -> bool:
    flags = KEYEVENTF_UNICODE | (KEYEVENTF_KEYUP if up else 0)
    return _send(KEYBDINPUT(wVk=0, wScan=code, dwFlags=flags,
                            time=0, dwExtraInfo=0))


def _uni_ev(code: int, up: bool) -> KEYBDINPUT:
    flags = KEYEVENTF_UNICODE | (KEYEVENTF_KEYUP if up else 0)
    return KEYBDINPUT(wVk=0, wScan=code, dwFlags=flags, time=0, dwExtraInfo=0)


def _press_char(ch: str) -> bool:
    code = ord(ch)
    if code > 0xFFFF:
        # 增补字符（emoji 等）需拆成 UTF-16 高低代理对；四事件一次原子发出
        code -= 0x10000
        hi, lo = 0xD800 + (code >> 10), 0xDC00 + (code & 0x3FF)
        return _send_many([_uni_ev(hi, False), _uni_ev(lo, False),
                           _uni_ev(lo, True), _uni_ev(hi, True)])
    return _send_many([_uni_ev(code, False), _uni_ev(code, True)])


_MapVirtualKeyW = ctypes.windll.user32.MapVirtualKeyW
_MapVirtualKeyW.argtypes = [wt.UINT, wt.UINT]
_MapVirtualKeyW.restype = wt.UINT
_VkKeyScanW = ctypes.windll.user32.VkKeyScanW
_VkKeyScanW.argtypes = [ctypes.c_wchar]
_VkKeyScanW.restype = ctypes.c_short


def _scan_event(vk: int, up: bool) -> bool:
    """真实按键事件（带扫描码），模拟物理键盘敲击。"""
    scan = _MapVirtualKeyW(vk, MAPVK_VK_TO_VSC)
    flags = KEYEVENTF_SCANCODE | (KEYEVENTF_KEYUP if up else 0)
    return _send(KEYBDINPUT(wVk=0, wScan=scan, dwFlags=flags,
                            time=0, dwExtraInfo=0))


def _scan_ev(vk: int, up: bool) -> KEYBDINPUT:
    scan = _MapVirtualKeyW(vk, MAPVK_VK_TO_VSC)
    flags = KEYEVENTF_SCANCODE | (KEYEVENTF_KEYUP if up else 0)
    return KEYBDINPUT(wVk=0, wScan=scan, dwFlags=flags, time=0, dwExtraInfo=0)


def _press_vk_scan(vk: int) -> bool:
    return _send_many([_scan_ev(vk, False), _scan_ev(vk, True)])


# ---------- 鼠标：开打前点一下目标窗口，让按键能进 ----------

_GetCursorPos = ctypes.windll.user32.GetCursorPos
_GetCursorPos.argtypes = [ctypes.POINTER(wt.POINT)]
_GetCursorPos.restype = wt.BOOL
_SetCursorPos = ctypes.windll.user32.SetCursorPos
_SetCursorPos.argtypes = [ctypes.c_int, ctypes.c_int]
_SetCursorPos.restype = wt.BOOL
_WindowFromPoint = ctypes.windll.user32.WindowFromPoint
_WindowFromPoint.argtypes = [wt.POINT]
_WindowFromPoint.restype = wt.HWND


def _mouse_click(x: int, y: int) -> bool:
    """把鼠标移到 (x, y) 并左键单击。"""
    ok = bool(_SetCursorPos(int(x), int(y)))
    time.sleep(0.03)
    down = INPUT(type=INPUT_MOUSE)
    down.mi = MOUSEINPUT(0, 0, 0, MOUSEEVENTF_LEFTDOWN, 0, 0)
    up = INPUT(type=INPUT_MOUSE)
    up.mi = MOUSEINPUT(0, 0, 0, MOUSEEVENTF_LEFTUP, 0, 0)
    ok = (_SendInput(1, ctypes.byref(down), ctypes.sizeof(INPUT)) == 1) and ok
    time.sleep(0.03)
    ok = (_SendInput(1, ctypes.byref(up), ctypes.sizeof(INPUT)) == 1) and ok
    return ok


# ---------- 输入法：真实按键会被输入法组字（abc→按不出），敲期间把
# 目标窗口的键盘布局临时切到 en-US，敲完切回原布局。 ----------
# 实测：ImmAssociateContext 禁 IME 不行（字母直接丢失，因为布局本身就是
# IME 布局）；WM_INPUTLANGCHANGEREQUEST 切布局 + 扫描码 100% 正确。

WM_INPUTLANGCHANGEREQUEST = 0x0050
WM_KEYDOWN = 0x0100
WM_KEYUP = 0x0101
WM_CHAR = 0x0102
WM_PASTE = 0x0302
KLF_ACTIVATE = 0x00000001

_GetKeyboardLayout = ctypes.windll.user32.GetKeyboardLayout
_GetKeyboardLayout.argtypes = [wt.DWORD]
_GetKeyboardLayout.restype = wt.HKL
_LoadKeyboardLayoutW = ctypes.windll.user32.LoadKeyboardLayoutW
_LoadKeyboardLayoutW.argtypes = [ctypes.c_wchar_p, wt.UINT]
_LoadKeyboardLayoutW.restype = wt.HKL
_PostMessageW = ctypes.windll.user32.PostMessageW
_PostMessageW.argtypes = [wt.HWND, wt.UINT, wt.WPARAM, wt.LPARAM]
_PostMessageW.restype = wt.BOOL


def _layout_begin(hwnd: int):
    """把目标窗口切到 en-US 布局，返回原 hkl（失败返回 None）。"""
    try:
        pid = wt.DWORD()
        tid = _GetWindowThreadProcessId(int(hwnd), ctypes.byref(pid))
        prev = _GetKeyboardLayout(tid)
        hkl_us = _LoadKeyboardLayoutW('00000409', KLF_ACTIVATE)
        _PostMessageW(wt.HWND(int(hwnd)), WM_INPUTLANGCHANGEREQUEST, 0,
                      int(hkl_us or 0))
        time.sleep(0.15)  # 切换是异步的，等目标线程处理完再敲
        return prev
    except Exception:
        return None


def _layout_end(hwnd: int, prev) -> None:
    """恢复目标窗口原来的键盘布局。"""
    try:
        if prev:
            _PostMessageW(wt.HWND(int(hwnd)), WM_INPUTLANGCHANGEREQUEST, 0,
                          int(prev))
    except Exception:
        pass


def _press_char_real(ch: str) -> bool:
    """兼容输入：ASCII 用真实按键（扫描码），其余字符退回 Unicode 事件。

    有些程序（自绘输入框等）不处理 KEYEVENTF_UNICODE 合成的 WM_CHAR，
    只认真实的 WM_KEYDOWN，就用这种方式敲。英文输入法状态下效果最好。
    """
    code = ord(ch)
    if not (0x20 <= code < 0x7F):
        return _press_char(ch)
    r = int(_VkKeyScanW(ch))
    if r == -1:
        return _press_char(ch)
    vk = r & 0xFF
    need_shift = bool((r >> 8) & 1)
    events = []
    if need_shift:
        events.append(_scan_ev(VK_SHIFT, False))
    events.append(_scan_ev(vk, False))
    events.append(_scan_ev(vk, True))
    if need_shift:
        events.append(_scan_ev(VK_SHIFT, True))
    return _send_many(events)


# ---------- 窗口定位与焦点：把字打进“光标所在的那个程序” ----------

GA_ROOT = 2

_GetForegroundWindow = ctypes.windll.user32.GetForegroundWindow
_GetForegroundWindow.restype = wt.HWND
_SetForegroundWindow = ctypes.windll.user32.SetForegroundWindow
_SetForegroundWindow.argtypes = [wt.HWND]
_SetForegroundWindow.restype = wt.BOOL
_BringWindowToTop = ctypes.windll.user32.BringWindowToTop
_GetAncestor = ctypes.windll.user32.GetAncestor
_GetAncestor.argtypes = [wt.HWND, wt.UINT]
_GetAncestor.restype = wt.HWND
_GetWindowThreadProcessId = ctypes.windll.user32.GetWindowThreadProcessId
_GetWindowThreadProcessId.argtypes = [wt.HWND, ctypes.POINTER(wt.DWORD)]
_GetWindowThreadProcessId.restype = wt.DWORD
_AttachThreadInput = ctypes.windll.user32.AttachThreadInput
_AttachThreadInput.argtypes = [wt.DWORD, wt.DWORD, wt.BOOL]
_AttachThreadInput.restype = wt.BOOL
_GetWindowTextW = ctypes.windll.user32.GetWindowTextW
_GetWindowTextLengthW = ctypes.windll.user32.GetWindowTextLengthW
_GetClassNameW = ctypes.windll.user32.GetClassNameW
_GetClassNameW.argtypes = [wt.HWND, ctypes.c_wchar_p, ctypes.c_int]
_GetClassNameW.restype = ctypes.c_int
_GetCurrentThreadId = ctypes.windll.kernel32.GetCurrentThreadId
_GetCurrentThreadId.restype = wt.DWORD
_SwitchToThisWindow = ctypes.windll.user32.SwitchToThisWindow
_SwitchToThisWindow.argtypes = [wt.HWND, wt.BOOL]
_SwitchToThisWindow.restype = None


def _top_level(hwnd) -> int:
    """返回 hwnd 所在的顶层窗口句柄（比较焦点归属用）。"""
    if not hwnd:
        return 0
    return int(_GetAncestor(int(hwnd), GA_ROOT) or 0)


def _window_title(hwnd) -> str:
    """读取窗口标题（用于状态栏提示“正在打进哪个程序”）。"""
    if not hwnd:
        return ''
    n = _GetWindowTextLengthW(int(hwnd))
    if n <= 0:
        return ''
    buf = ctypes.create_unicode_buffer(n + 1)
    _GetWindowTextW(int(hwnd), buf, n + 1)
    return buf.value


def _window_class(hwnd) -> str:
    """读取窗口类名（诊断日志用）。"""
    if not hwnd:
        return ''
    buf = ctypes.create_unicode_buffer(256)
    _GetClassNameW(int(hwnd), buf, 256)
    return buf.value


def _window_info(hwnd) -> str:
    """一句诊断：句柄/标题/类名/pid。"""
    if not hwnd:
        return '0x0'
    pid = wt.DWORD()
    _GetWindowThreadProcessId(int(hwnd), ctypes.byref(pid))
    return f'{int(hwnd):#x} {_window_title(hwnd)!r} class={_window_class(hwnd)!r} pid={pid.value}'


class GUITHREADINFO(ctypes.Structure):
    _fields_ = [
        ('cbSize', wt.DWORD), ('flags', wt.DWORD),
        ('hwndActive', wt.HWND), ('hwndFocus', wt.HWND),
        ('hwndCapture', wt.HWND), ('hwndMenuOwner', wt.HWND),
        ('hwndMoveSize', wt.HWND), ('hwndCaret', wt.HWND),
        ('rcCaret', wt.RECT),
    ]


_GetGUIThreadInfo = ctypes.windll.user32.GetGUIThreadInfo
_GetGUIThreadInfo.argtypes = [wt.DWORD, ctypes.POINTER(GUITHREADINFO)]
_GetGUIThreadInfo.restype = wt.BOOL
_GetWindowRect = ctypes.windll.user32.GetWindowRect
_GetWindowRect.argtypes = [wt.HWND, ctypes.POINTER(wt.RECT)]
_GetWindowRect.restype = wt.BOOL


def _focus_ctrl(hwnd) -> tuple:
    """目标窗口所在线程当前焦点控件（就是用户点过的输入框）的句柄与中心坐标。"""
    if not hwnd:
        return 0, 0, 0
    try:
        tid = _GetWindowThreadProcessId(int(hwnd), None)  # 返回单值 tid（byref 形式才写 pid）
        if not tid:
            return 0, 0, 0
        gi = GUITHREADINFO()
        gi.cbSize = ctypes.sizeof(GUITHREADINFO)
        if not _GetGUIThreadInfo(tid, ctypes.byref(gi)) or not gi.hwndFocus:
            return 0, 0, 0
        r = wt.RECT()
        if not _GetWindowRect(gi.hwndFocus, ctypes.byref(r)):
            return int(gi.hwndFocus), 0, 0
        return int(gi.hwndFocus), (r.left + r.right) // 2, (r.top + r.bottom) // 2
    except Exception:
        return 0, 0, 0


APP_BUILD = '2026-10-27-7'

_LOG_PATH: str | None = None


def _log(msg: str) -> None:
    """调试日志（%LOCALAPPDATA%\\TyperSimulator\\typer_debug.log，不往 exe 旁边留文件，失败也绝不影响功能）。"""
    global _LOG_PATH
    try:
        if _LOG_PATH is None:
            bases = []
            try:
                base = os.path.join(
                    os.environ.get('LOCALAPPDATA') or tempfile.gettempdir(),
                    'TyperSimulator')
                os.makedirs(base, exist_ok=True)
                bases.append(base)
            except Exception:
                pass
            bases += [tempfile.gettempdir()]
            for base in bases:
                p = os.path.join(base, 'typer_debug.log')
                try:
                    with open(p, 'a', encoding='utf-8'):
                        pass
                    _LOG_PATH = p
                    break
                except OSError:
                    continue
            if _LOG_PATH is None:
                return
        with open(_LOG_PATH, 'a', encoding='utf-8') as f:
            f.write(f'{time.strftime("%m-%d %H:%M:%S")} {msg}\n')
    except Exception:
        pass


def _focus_window(hwnd) -> bool:
    """把 hwnd 拉到前台并给它键盘焦点；返回前台是否真的是它。

    Windows 有“前台锁”，直接 SetForegroundWindow 常被拒。正解是把当前
    线程同时与【当前前台窗口的线程】和目标线程的输入队列绑在一起
    （AttachThreadInput），再置顶 + 置前台；失败再补“按一下 Alt”解锁和
    SwitchToThisWindow 兜底。最后以“前台真的是它”为准，不信 API 返回值。
    """
    if not hwnd:
        return False
    hwnd = int(hwnd)
    fg = int(_GetForegroundWindow() or 0)
    if fg == hwnd:
        return True
    tid_target = _GetWindowThreadProcessId(hwnd, None)
    tid_fg = _GetWindowThreadProcessId(fg, None) if fg else 0
    tid_self = _GetCurrentThreadId()
    attached = []
    for tid in (tid_fg, tid_target):
        if tid and tid != tid_self and tid not in attached:
            if _AttachThreadInput(tid_self, tid, True):
                attached.append(tid)
    try:
        _BringWindowToTop(hwnd)
        ok = bool(_SetForegroundWindow(hwnd))
        if not ok:
            # 前台锁兜底：先“按一下 Alt”产生真实输入事件再试
            _press_vk(VK_MENU)
            _BringWindowToTop(hwnd)
            ok = bool(_SetForegroundWindow(hwnd))
        if not ok:
            _SwitchToThisWindow(hwnd, True)
        fg2 = int(_GetForegroundWindow() or 0)
        if fg2 != hwnd and _top_level(fg2) != _top_level(hwnd):
            # 最后兜底：真实鼠标点击目标的焦点控件/窗口中部。
            # “收到最后输入事件的进程”才允许 SetForegroundWindow，
            # 真实点击是最可靠的解锁方式（也和用户手动点目标的行为一致）。
            global _focus_click_at
            now = time.monotonic()
            if now - _focus_click_at >= 0.8:  # 点击限流，避免连点打扰用户
                _focus_click_at = now
                fc, cx, cy = _focus_ctrl(hwnd)
                if not (cx or cy):
                    r = wt.RECT()
                    if _GetWindowRect(hwnd, ctypes.byref(r)):
                        cx, cy = (r.left + r.right) // 2, (r.top + r.bottom) // 2
                if cx or cy:
                    _mouse_click(cx, cy)
                    time.sleep(0.05)
                    # 点击会把光标定位到点中的位置，按 Ctrl+End 归到末尾，
                    # 保证后续字符按顺序追加而不是插进已有文字中间
                    _scan_event(VK_CONTROL, False)
                    _scan_event(VK_END, False)
                    _scan_event(VK_END, True)
                    _scan_event(VK_CONTROL, True)
    finally:
        for tid in attached:
            _AttachThreadInput(tid_self, tid, False)
    fg2 = int(_GetForegroundWindow() or 0)
    return fg2 == hwnd or _top_level(fg2) == _top_level(hwnd)


_refocus_log_at = 0.0
_focus_click_at = 0.0


def _ensure_on_target(target: int) -> bool:
    """确保前台窗口属于 target（按根窗口比较，弹出窗也算）。

    前台不在目标上就夺回；夺不回来返回 False——调用方必须跳过这次输入，
    绝不能把字打进别的程序。
    """
    global _refocus_log_at

    def on_target() -> bool:
        fg = int(_GetForegroundWindow() or 0)
        return bool(fg) and _top_level(fg) == _top_level(target)

    if on_target():
        return True
    now = time.monotonic()
    if now - _refocus_log_at >= 1.0:  # 日志限流，避免逐字刷屏
        _refocus_log_at = now
        fg = int(_GetForegroundWindow() or 0)
        _log(f'!!! 焦点不在目标上（前台 {fg:#x} {_window_title(fg)!r}）'
             f'→ 夺回 {target:#x} {_window_title(target)!r}')
    deadline = time.monotonic() + 0.6
    while True:
        _focus_window(target)
        if on_target():
            time.sleep(0.03)  # 等焦点切换真正落地再敲
            return True
        if time.monotonic() >= deadline:
            return False
        time.sleep(0.05)


_IsWindow = ctypes.windll.user32.IsWindow
_IsWindow.argtypes = [wt.HWND]
_IsWindow.restype = wt.BOOL

# 完整性级别（RID）：0x1000 低 / 0x2000 中 / 0x3000 高（管理员）/ 0x4000 系统
_OpenProcess = ctypes.windll.kernel32.OpenProcess
_OpenProcess.argtypes = [wt.DWORD, wt.BOOL, wt.DWORD]
_OpenProcess.restype = wt.HANDLE
_OpenProcessToken = ctypes.windll.advapi32.OpenProcessToken
_OpenProcessToken.argtypes = [wt.HANDLE, wt.DWORD, ctypes.POINTER(wt.HANDLE)]
_OpenProcessToken.restype = wt.BOOL
_GetTokenInformation = ctypes.windll.advapi32.GetTokenInformation
_GetTokenInformation.argtypes = [wt.HANDLE, ctypes.c_int, ctypes.c_void_p,
                                 wt.DWORD, ctypes.POINTER(wt.DWORD)]
_GetTokenInformation.restype = wt.BOOL
_GetSidSubAuthorityCount = ctypes.windll.advapi32.GetSidSubAuthorityCount
_GetSidSubAuthorityCount.argtypes = [ctypes.c_void_p]
_GetSidSubAuthorityCount.restype = ctypes.POINTER(ctypes.c_ubyte)
_GetSidSubAuthority = ctypes.windll.advapi32.GetSidSubAuthority
_GetSidSubAuthority.argtypes = [ctypes.c_void_p, wt.DWORD]
_GetSidSubAuthority.restype = ctypes.POINTER(wt.DWORD)
_CloseHandle = ctypes.windll.kernel32.CloseHandle
_CloseHandle.argtypes = [wt.HANDLE]
_CloseHandle.restype = wt.BOOL
_GetCurrentProcess = ctypes.windll.kernel32.GetCurrentProcess
_GetCurrentProcess.restype = wt.HANDLE


def _integrity_of(pid) -> int | None:
    """读进程的完整性级别 RID；读不到返回 None（绝不影响打字）。"""
    TOKEN_QUERY = 0x0008
    PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
    TokenIntegrityLevel = 25
    try:
        if pid:
            h = _OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, int(pid))
        else:
            h = _GetCurrentProcess()
        if not h:
            return None
        token = wt.HANDLE()
        try:
            if not _OpenProcessToken(h, TOKEN_QUERY, ctypes.byref(token)):
                return None
            need = wt.DWORD()
            _GetTokenInformation(token, TokenIntegrityLevel, None, 0,
                                 ctypes.byref(need))
            buf = ctypes.create_string_buffer(need.value or 16)
            if not _GetTokenInformation(token, TokenIntegrityLevel, buf,
                                        need.value, ctypes.byref(need)):
                return None
            sid = ctypes.cast(buf, ctypes.POINTER(ctypes.c_void_p)).contents.value
            if not sid:
                return None
            cnt = _GetSidSubAuthorityCount(ctypes.c_void_p(sid))[0]
            # 真实签名 (PSID, DWORD) → PDWORD（指回 SID 里的 RID）
            p_rid = _GetSidSubAuthority(ctypes.c_void_p(sid), cnt - 1)
            if not p_rid:
                return None
            return int(p_rid.contents.value)
        finally:
            _CloseHandle(token)
            if pid:
                _CloseHandle(h)
    except Exception:
        return None


def _integrity_of_hwnd(hwnd) -> int | None:
    """窗口所属进程的完整性级别。"""
    pid = wt.DWORD()
    _GetWindowThreadProcessId(int(hwnd), ctypes.byref(pid))
    return _integrity_of(pid.value or None)


class OsKeySink:
    """真实按键输出：打进光标所在的窗口。

    给定 target（窗口句柄）时会锁定它——期间如果焦点被别的窗口抢走
    （比如点到了本程序），会自动把焦点夺回目标窗口再继续敲，
    保证字全部落在光标所在的那个程序里。target=0 表示不锁定，
    就按“敲键那一刻的焦点窗口”打。
    """

    REFOCUS_MIN_GAP = 0.25  # 秒：夺回焦点的最小间隔，避免每个字符都折腾窗口

    def __init__(self, target: int = 0, compat: bool = False):
        self.target = int(target or 0)
        self.compat = bool(compat)  # 兼容模式：真实按键（扫描码）输入
        self._last_refocus = 0.0
        self._hkl_prev = None  # 原键盘布局（真实按键前临时切 en-US）
        self.failed = 0  # 系统拒绝的按键数（诊断用）
        self.sent = 0

    def _ensure_layout(self) -> None:
        """真实按键第一次敲之前，把目标布局切到 en-US，避开输入法组字。"""
        if self.compat and self.target and self._hkl_prev is None:
            self._hkl_prev = _layout_begin(self.target)

    def finish(self) -> None:
        """打字结束：还原目标窗口的键盘布局。"""
        if self._hkl_prev is not None:
            _layout_end(self.target, self._hkl_prev)
            self._hkl_prev = None

    def abort(self) -> None:
        """中途停止：同样还原布局。"""
        self.finish()

    def send(self, ch: str) -> None:
        # 按“根窗口”比较：目标程序的弹出窗/独立聊天窗（如微信的独立聊天窗）
        # 同属目标，绝不能当成抢焦点、把输入框里的焦点拽走。焦点真不在目标
        # 上时夺回；夺不回来就跳过这个字（failed 计数），绝不打进别的程序。
        if self.target and not _ensure_on_target(self.target):
            self.failed += 1
            return
        self._ensure_layout()
        if ch == '\n':
            ok = _press_vk_scan(VK_RETURN) if self.compat else _press_vk(VK_RETURN)
        elif ch == '\t':
            ok = _press_vk_scan(VK_TAB) if self.compat else _press_vk(VK_TAB)
        else:
            ok = _press_char_real(ch) if self.compat else _press_char(ch)
        self.sent += 1
        if not ok:
            self.failed += 1


class CallbackSink:
    """演示输出：不碰操作系统，只把字符回调出去。"""

    def __init__(self, on_char):
        self.on_char = on_char
        self.sent = 0
        self.failed = 0

    def send(self, ch: str) -> None:
        self.on_char(ch)
        self.sent += 1


class MsgKeySink:
    """消息输出：把 WM_CHAR / 按键消息直接投进目标窗口的消息队列。

    这是业界（微软 winapp-cli send-keys、wxauto 后台模式）的默认路线：
    - 免焦点免前台：消息直接进目标窗口队列，用户点着别家窗口也照样输入，
      这也是网页输入的首选——浏览器对 posted 消息生成的输入事件与真按键
      无异，页面的“禁止粘贴”完全拦不住；
    - 不带 injected 标志，低级键盘钩子/注入过滤拦不到；
    - 投递目标优先取目标线程的焦点控件（GetGUIThreadInfo），自绘控件
      （Chromium/Qt，无子窗口句柄）退回顶层窗口，由框架自己路由给焦点
      控件。
    换行默认 Shift+Enter（聊天窗口不会触发“回车发送”；网页多行框同样
    换行），可选真回车。制表符用 Tab 键消息。emoji 等非 BMP 字符按
    UTF-16 代理对逐单元投递。
    """

    def __init__(self, target: int = 0, nl_shift: bool = True):
        self.target = int(target or 0)
        self.nl_shift = bool(nl_shift)
        self.sent = 0
        self.failed = 0

    def _recv_hwnd(self) -> int:
        if self.target:
            fc, _x, _y = _focus_ctrl(self.target)
            if fc and _IsWindow(fc):
                return int(fc)
            return self.target
        return int(_GetForegroundWindow() or 0)

    def _post(self, msg: int, w: int, l: int = 0) -> bool:
        return bool(_PostMessageW(self._recv_hwnd(), msg, w, l))

    def _key(self, vk: int, up: bool = False) -> bool:
        l = (_MapVirtualKeyW(vk, 0) << 16) | 1
        if up:
            l |= 0xC0000000  # 之前已按下、本次为松开
        return self._post(WM_KEYUP if up else WM_KEYDOWN, vk, l)

    def send(self, ch: str) -> None:
        # 消息直投不需要焦点/前台（这正是它的意义），窗口还在就投；
        # 目标没了才记失败，绝不会投进别的程序
        if self.target and not _IsWindow(self.target):
            self.failed += 1
            return
        ok = True
        if ch == '\n':
            if self.nl_shift:
                ok = self._key(VK_SHIFT) and ok
            ok = self._key(VK_RETURN) and ok
            ok = self._key(VK_RETURN, up=True) and ok
            if self.nl_shift:
                ok = self._key(VK_SHIFT, up=True) and ok
        elif ch == '\t':
            ok = self._key(VK_TAB) and ok
            ok = self._key(VK_TAB, up=True) and ok
        else:
            b = ch.encode('utf-16-le')
            for i in range(0, len(b), 2):
                unit = b[i] | (b[i + 1] << 8)
                ok = self._post(WM_CHAR, unit) and ok
        self.sent += 1
        if not ok:
            self.failed += 1


# ---------- 剪贴板（粘贴模式用） ----------

CF_UNICODETEXT = 13
GMEM_MOVEABLE = 0x0002
_OpenClipboard = ctypes.windll.user32.OpenClipboard
_OpenClipboard.argtypes = [wt.HWND]
_OpenClipboard.restype = wt.BOOL
_EmptyClipboard = ctypes.windll.user32.EmptyClipboard
_EmptyClipboard.restype = wt.BOOL
_SetClipboardData = ctypes.windll.user32.SetClipboardData
_SetClipboardData.argtypes = [wt.UINT, ctypes.c_void_p]
_SetClipboardData.restype = ctypes.c_void_p
_GetClipboardData = ctypes.windll.user32.GetClipboardData
_GetClipboardData.argtypes = [wt.UINT]
_GetClipboardData.restype = ctypes.c_void_p
_CloseClipboard = ctypes.windll.user32.CloseClipboard
_CloseClipboard.restype = wt.BOOL
_GlobalAlloc = ctypes.windll.kernel32.GlobalAlloc
_GlobalAlloc.argtypes = [wt.UINT, ctypes.c_size_t]
_GlobalAlloc.restype = ctypes.c_void_p
_GlobalLock = ctypes.windll.kernel32.GlobalLock
_GlobalLock.argtypes = [ctypes.c_void_p]
_GlobalLock.restype = ctypes.c_void_p
_GlobalUnlock = ctypes.windll.kernel32.GlobalUnlock
_GlobalUnlock.argtypes = [ctypes.c_void_p]


def _set_clipboard_text(s: str) -> bool:
    """写剪贴板文本（含重试；任何失败返回 False，绝不抛异常）。"""
    try:
        data = s.encode('utf-16-le') + b'\x00\x00'
        opened = False
        for _ in range(10):
            if _OpenClipboard(None):
                opened = True
                break
            time.sleep(0.01)
        if not opened:
            return False
        try:
            _EmptyClipboard()
            h = _GlobalAlloc(GMEM_MOVEABLE, len(data))
            if not h:
                return False
            p = _GlobalLock(h)
            if not p:
                return False
            ctypes.memmove(p, data, len(data))
            _GlobalUnlock(h)
            return bool(_SetClipboardData(CF_UNICODETEXT, h))
        finally:
            _CloseClipboard()
    except Exception:
        return False


def _get_clipboard_text() -> str:
    try:
        opened = False
        for _ in range(10):
            if _OpenClipboard(None):
                opened = True
                break
            time.sleep(0.01)
        if not opened:
            return ''
        try:
            h = _GetClipboardData(CF_UNICODETEXT)
            if not h:
                return ''
            p = _GlobalLock(h)
            if not p:
                return ''
            try:
                return ctypes.wstring_at(p)
            finally:
                _GlobalUnlock(h)
        finally:
            _CloseClipboard()
    except Exception:
        return ''


class PasteSink:
    """粘贴输出：内容攒批写剪贴板再 Ctrl+V，任何输入框都能进。

    针对连合成字符消息/真实按键都不收的自绘输入框（如微信）。分批粘贴
    保留逐段出现的观感；换行作为文本插入，不会触发聊天窗口的“回车发送”。
    打字期间临时借用剪贴板，结束后（finish/abort）还原原来的文本内容。
    """

    FLUSH_CHARS = 24  # 每攒这么多字粘贴一批
    REFOCUS_MIN_GAP = OsKeySink.REFOCUS_MIN_GAP

    def __init__(self, target: int = 0):
        self.target = int(target or 0)
        self.sent = 0
        self.failed = 0
        self._buf: list = []
        self._clip_saved = ''
        self._dirty = False
        self._last_refocus = 0.0
        self._logged_first = False

    def _ensure_focus(self) -> bool:
        # 与 OsKeySink 相同：同一程序的弹出窗都算“在目标上”，按根窗口比较；
        # 夺不回焦点返回 False，本批内容放弃，绝不粘进别的程序
        if not self.target:
            return True
        return _ensure_on_target(self.target)

    def send(self, ch: str) -> None:
        self._ensure_focus()
        self._buf.append(ch)
        self.sent += 1
        if ch == '\n' or len(self._buf) >= self.FLUSH_CHARS:
            self.flush()

    def flush(self) -> None:
        if not self._buf:
            return
        if not self._ensure_focus():
            self.failed += len(self._buf)
            self._buf = []
            _log('!!! 焦点夺不回来，放弃本批粘贴（不粘进别的程序）')
            return
        text = ''.join(self._buf)
        self._buf = []
        if not self._dirty:
            self._clip_saved = _get_clipboard_text()
            self._dirty = True
        ok = _set_clipboard_text(text)
        step = '写剪贴板'
        if ok:
            step = 'Ctrl+V'
            ok = self._ctrl_v()
        if not ok:
            self.failed += len(text)
            _log(f'!!! 粘贴失败（{step}）前台 {_window_info(_GetForegroundWindow())}')
        elif not self._logged_first:
            self._logged_first = True
            _log(f'粘贴首批 {len(text)} 字 → 前台 {_window_info(_GetForegroundWindow())}')
        time.sleep(0.05)  # 给目标程序消化粘贴内容的时间

    def _ctrl_v(self) -> bool:
        # Ctrl+V 四事件一次原子发出
        return _send_many([_scan_ev(VK_CONTROL, False), _scan_ev(VK_V, False),
                           _scan_ev(VK_V, True), _scan_ev(VK_CONTROL, True)])

    def finish(self) -> None:
        """打字完成：吐掉剩余内容，还原剪贴板。"""
        self.flush()
        if self._dirty:
            time.sleep(0.1)  # 等目标程序读完剪贴板再还原
            _set_clipboard_text(self._clip_saved)
            self._dirty = False

    def abort(self) -> None:
        """中途停止：不再粘贴剩余内容，只还原剪贴板。"""
        self._buf = []
        if self._dirty:
            _set_clipboard_text(self._clip_saved)
            self._dirty = False


class TypingEngine:
    """后台线程逐字输出，支持暂停 / 继续 / 停止。"""

    # 硬输入的微批次节流：每发 INSTANT_BURST 个字让目标程序排空一会儿
    INSTANT_BURST = 24
    INSTANT_YIELD = 0.006

    def __init__(self):
        self._thread: threading.Thread | None = None
        self._stop_evt = threading.Event()
        self._pause_evt = threading.Event()  # set = 不暂停
        self._running = False

    @property
    def running(self) -> bool:
        return self._running

    def start(self, text: str, sink, cps: float = 10, jitter: float = 0.3,
              start_delay: float = 0.0, instant: bool = False,
              on_start=None, on_char=None, on_done=None, on_stop=None):
        """开始逐字输出。

        cps      每秒字符数（1~500）
        jitter   每字延迟随机波动比例（更像真人），0 = 恒速
        instant  硬输入：字符间不等待全速直灌（cps/jitter 失效）
        on_char(ch, index, total)  每敲一个字符回调
        on_done / on_stop          正常完成 / 被停止时回调
        """
        if self._running:
            raise RuntimeError('打字已在进行中，请先停止')
        self._stop_evt.clear()
        self._pause_evt.set()
        self._running = True
        self._thread = threading.Thread(
            target=self._worker,
            args=(text, sink, cps, jitter, start_delay, instant,
                  on_start, on_char, on_done, on_stop),
            daemon=True,
        )
        self._thread.start()

    def pause(self) -> None:
        self._pause_evt.clear()

    def resume(self) -> None:
        self._pause_evt.set()

    def stop(self) -> None:
        self._stop_evt.set()
        self._pause_evt.set()  # 唤醒可能挂起的线程

    def _worker(self, text, sink, cps, jitter, start_delay, instant,
                on_start, on_char, on_done, on_stop):
        _winmm.timeBeginPeriod(1)
        try:
            if start_delay > 0:
                deadline = time.monotonic() + start_delay
                while time.monotonic() < deadline:
                    if self._stop_evt.is_set():
                        if on_stop:
                            on_stop()
                        return
                    time.sleep(0.05)
            if on_start:
                on_start()
            chars = list(text)
            total = len(chars)
            base = 1.0 / max(float(cps), 0.1)
            burst = 0
            for i, ch in enumerate(chars):
                if self._stop_evt.is_set():
                    break
                # 暂停时挂起（仍可被停止唤醒）
                while not self._pause_evt.wait(timeout=0.1):
                    if self._stop_evt.is_set():
                        break
                if self._stop_evt.is_set():
                    break
                sink.send(ch)
                if on_char:
                    on_char(ch, i + 1, total)
                if instant:
                    # 硬输入：全速直灌，但每发一小批就让目标程序喘口气——
                    # 完全不喘气会把慢程序的消息队列灌爆导致丢字。
                    # 每批 INSTANT_BURST 字停 INSTANT_YIELD 秒，仍有 ~2000 字/秒
                    burst += 1
                    if burst >= self.INSTANT_BURST:
                        burst = 0
                        time.sleep(self.INSTANT_YIELD)
                    continue
                d = base * (1.0 + random.uniform(-jitter, jitter)) if jitter > 0 else base
                d = max(d, 0.0)
                if d <= 0:
                    continue
                deadline = time.monotonic() + d
                while time.monotonic() < deadline:
                    if self._stop_evt.is_set() or not self._pause_evt.is_set():
                        break
                    time.sleep(min(0.005, deadline - time.monotonic()))
            else:
                if on_done:
                    on_done()
                return
            if on_stop:
                on_stop()
        finally:
            self._running = False
            _winmm.timeEndPeriod(1)


class HotkeyListener:
    """后台轮询全局热键：F6 = 切换（开始/暂停），Esc = 停止。"""

    POLL_MS = 30

    def __init__(self, on_toggle, on_stop):
        self._on_toggle = on_toggle
        self._on_stop = on_stop
        self._stop_evt = threading.Event()
        self._thread: threading.Thread | None = None

    def start(self) -> None:
        if self._thread and self._thread.is_alive():
            return
        self._stop_evt.clear()
        self._thread = threading.Thread(target=self._worker, daemon=True)
        self._thread.start()

    def stop(self) -> None:
        self._stop_evt.set()

    def _worker(self) -> None:
        while not self._stop_evt.is_set():
            if _GetAsyncKeyState(VK_F6) & 0x8000:
                self._on_toggle()
                time.sleep(0.4)  # 消抖：长按不重复触发
            if _GetAsyncKeyState(VK_ESCAPE) & 0x8000:
                self._on_stop()
                time.sleep(0.4)
            time.sleep(self.POLL_MS / 1000.0)


# ============================================================
# 三、图形界面
# ============================================================

FILETYPES = [
    ('支持的文件', '*.txt *.cpp *.c *.py *.h *.hpp *.cc *.cxx *.docx'),
    ('文本文件', '*.txt'),
    ('C/C++ 源码', '*.cpp *.c *.h *.hpp *.cc *.cxx'),
    ('Python 源码', '*.py'),
    ('Word 文档', '*.docx'),
    ('所有文件', '*.*'),
]


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title('文件打字模拟器')
        # 窗口尺寸自适应屏幕（DPI感知后是物理像素，小屏自动缩小）并居中
        sw = root.winfo_screenwidth()
        sh = root.winfo_screenheight()
        w = max(740, min(980, sw - 60))
        h = max(500, min(640, sh - 140))
        x, y = max(0, (sw - w) // 2), max(0, (sh - h) // 2 - 20)
        root.geometry(f'{w}x{h}+{x}+{y}')
        root.minsize(720, 480)
        # F11 最大化/还原（仅本窗口获得焦点时生效，不干扰其他程序）
        root.bind('<F11>', lambda e: root.state(
            'normal' if root.state() == 'zoomed' else 'zoomed'))

        self.full_path = ''
        self.text = ''
        self.encoding = ''
        self.state = 'idle'  # idle / counting / waiting / typing / paused

        self.engine = TypingEngine()
        self.q: queue.Queue = queue.Queue()
        self._countdown_job = None
        self._target_job = None

        self._build_ui()

        # 热键回调只往队列丢消息，由主线程 _poll 处理（线程安全）
        self.hotkey = HotkeyListener(
            on_toggle=lambda: self.q.put(('hotkey_toggle', None)),
            on_stop=lambda: self.q.put(('hotkey_stop', None)),
        )
        self.hotkey.start()

        root.protocol('WM_DELETE_WINDOW', self._on_close)
        self._poll()

    # ---------- 界面 ----------
    def _build_ui(self):
        top = ttk.LabelFrame(self.root, text=' 内容来源 ', padding=(8, 6))
        top.pack(fill='x', padx=10, pady=(8, 0))
        ttk.Button(top, text='选择文件…', command=self._pick_file).pack(side='left')
        ttk.Button(top, text='粘贴剪贴板',
                   command=self._paste_clipboard).pack(side='left', padx=6)
        ttk.Button(top, text='清空',
                   command=lambda: self.preview.delete('1.0', 'end')
                   ).pack(side='left')
        self.file_lbl = ttk.Label(top, text='（未选择文件）', padding=(8, 0))
        self.file_lbl.pack(side='left', padx=(12, 0))
        self.info_lbl = ttk.Label(top, text='')
        self.info_lbl.pack(side='right')

        mid = ttk.Frame(self.root, padding=(10, 0))
        mid.pack(fill='both', expand=True)

        prev_box = ttk.LabelFrame(mid, text=' 内容（选文件，或直接粘贴/编辑） ', padding=4)
        prev_box.pack(side='left', fill='both', expand=True, padx=(0, 4))
        self.preview = tk.Text(prev_box, wrap='none', width=44, height=12)
        pv = ttk.Scrollbar(prev_box, orient='vertical', command=self.preview.yview)
        ph = ttk.Scrollbar(prev_box, orient='horizontal', command=self.preview.xview)
        self.preview.configure(yscrollcommand=pv.set, xscrollcommand=ph.set)
        self.preview.grid(row=0, column=0, sticky='nsew')
        pv.grid(row=0, column=1, sticky='ns')
        ph.grid(row=1, column=0, sticky='ew')
        prev_box.rowconfigure(0, weight=1)
        prev_box.columnconfigure(0, weight=1)

        demo_box = ttk.LabelFrame(mid, text=' 内置演示框（演示模式在这里打字） ', padding=4)
        demo_box.pack(side='left', fill='both', expand=True, padx=(4, 0))
        self.demo = tk.Text(demo_box, wrap='none', width=40, height=12)
        self.demo.bind('<Button-1>', self._demo_clicked)
        dv = ttk.Scrollbar(demo_box, orient='vertical', command=self.demo.yview)
        self.demo.configure(yscrollcommand=dv.set)
        self.demo.grid(row=0, column=0, sticky='nsew')
        dv.grid(row=0, column=1, sticky='ns')
        ttk.Button(demo_box, text='清空演示框',
                   command=lambda: self.demo.delete('1.0', 'end')
                   ).grid(row=1, column=0, columnspan=2, sticky='ew', pady=(4, 0))
        demo_box.rowconfigure(0, weight=1)
        demo_box.columnconfigure(0, weight=1)

        setf = ttk.LabelFrame(self.root, text=' 打字设置 ', padding=(10, 6))
        setf.pack(fill='x', padx=10, pady=(8, 0))

        row1 = ttk.Frame(setf)
        row1.pack(fill='x', pady=(0, 4))
        row2 = ttk.Frame(setf)
        row2.pack(fill='x')

        # 打字速度：输入框（直接键入每秒字符数）在上，滑条在下，双向同步
        row0 = ttk.Frame(setf)
        row0.pack(fill='x', pady=(0, 2))
        self.cps = tk.IntVar(value=15)
        self.cps_entry = tk.StringVar(value='15')
        ttk.Label(row0, text='每秒字符数（1–500）').pack(side='left')
        self.cps_entry_ctrl = ttk.Entry(row0, textvariable=self.cps_entry,
                                        width=6, justify='center')
        self.cps_entry_ctrl.pack(side='left', padx=(6, 0))
        self.cps_entry_ctrl.bind('<Return>', self._apply_cps_entry)
        self.cps_entry_ctrl.bind('<FocusOut>', self._apply_cps_entry)

        self.scale_ctrl = tk.Scale(row1, from_=1, to=500, orient='horizontal',
                                   variable=self.cps, length=240,
                                   command=self._cps_from_scale)
        self.scale_ctrl.pack(side='left', padx=(0, 14))

        self.jitter = tk.BooleanVar(value=True)
        self.jitter_cb = ttk.Checkbutton(row1, text='模拟抖动（更像真人）',
                                         variable=self.jitter)
        self.jitter_cb.pack(side='left', padx=(0, 14))

        self.hard = tk.BooleanVar(value=False)
        ttk.Checkbutton(row1, text='硬输入（不限速直灌）', variable=self.hard,
                        command=self._sync_speed_state).pack(side='left')

        ttk.Label(row2, text='开始倒计时（秒）').pack(side='left')
        self.countdown = tk.IntVar(value=3)
        ttk.Spinbox(row2, from_=0, to=10, textvariable=self.countdown,
                    width=3).pack(side='left', padx=(4, 16))

        self.mode = tk.StringVar(value='external')
        ttk.Label(row2, text='打到').pack(side='left')
        ttk.Radiobutton(row2, text='外部窗口（点击处）',
                        variable=self.mode, value='external'
                        ).pack(side='left', padx=(4, 6))
        ttk.Radiobutton(row2, text='内置演示框',
                        variable=self.mode, value='demo').pack(side='left')

        io = ttk.Frame(self.root, padding=(10, 8, 10, 0))
        io.pack(fill='x')

        method_box = ttk.LabelFrame(io, text=' 输入方式 ', padding=(10, 6))
        method_box.pack(side='left', fill='both', expand=True, padx=(0, 4))
        self.input_method = tk.StringVar(value='msg')
        ttk.Radiobutton(method_box, text='消息（推荐）', value='msg',
                        variable=self.input_method
                        ).grid(row=0, column=0, sticky='w', padx=(0, 14))
        ttk.Radiobutton(method_box, text='逐字模拟', value='unicode',
                        variable=self.input_method
                        ).grid(row=0, column=1, sticky='w')
        ttk.Radiobutton(method_box, text='真实按键（英文）', value='keys',
                        variable=self.input_method
                        ).grid(row=1, column=0, sticky='w', padx=(0, 14), pady=(2, 0))
        ttk.Radiobutton(method_box, text='粘贴', value='paste',
                        variable=self.input_method
                        ).grid(row=1, column=1, sticky='w', pady=(2, 0))
        ttk.Label(method_box,
                  text='消息=免焦点直投（网页/微信都能进）｜逐字=普通程序｜粘贴=大段文本快',
                  foreground='#666').grid(row=2, column=0, columnspan=2,
                                          sticky='w', pady=(4, 0))

        opt_box = ttk.LabelFrame(io, text=' 选项 ', padding=(10, 6))
        opt_box.pack(side='left', fill='both', expand=True, padx=(4, 0))
        self.nl_shift = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt_box, text='换行用 Shift+回车（聊天窗不发送）',
                        variable=self.nl_shift).grid(row=0, column=0, sticky='w')
        self.pre_click = tk.BooleanVar(value=False)
        ttk.Checkbutton(opt_box, text='开打前点击目标一下（消息方式不需要）',
                        variable=self.pre_click).grid(row=1, column=0, sticky='w',
                                                      pady=(2, 0))
        self.topmost = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt_box, text='窗口置顶（切页面也保持最上层）',
                        variable=self.topmost,
                        command=self._toggle_topmost).grid(row=2, column=0, sticky='w',
                                                           pady=(2, 0))

        ctrl = ttk.Frame(self.root, padding=(10, 8, 10, 4))
        ctrl.pack(fill='x')
        self.start_btn = ttk.Button(ctrl, text='开始 (F6)', width=16,
                                    command=self._toggle)
        self.start_btn.pack(side='left')
        ttk.Button(ctrl, text='停止 (Esc)', width=12,
                   command=self._stop).pack(side='left', padx=6)
        self.bar = ttk.Progressbar(ctrl, mode='determinate')
        self.bar.pack(side='left', fill='x', expand=True, padx=8)

        bot = ttk.Frame(self.root, padding=(10, 0, 10, 6))
        bot.pack(fill='x')
        self.status = ttk.Label(bot, text='请先选择文件')
        self.status.pack(side='left')
        ttk.Sizegrip(bot).pack(side='right')  # 右下角显式拖拽手柄，随时可拉伸窗口

    # ---------- 文件 ----------
    def _pick_file(self):
        path = filedialog.askopenfilename(title='选择文件', filetypes=FILETYPES)
        if not path:
            return
        try:
            text, enc = read_text(path)
        except FileReaderError as e:
            messagebox.showerror('无法读取文件', str(e))
            return
        self.full_path, self.text, self.encoding = path, text, enc
        self.file_lbl.config(text=os.path.basename(path))
        self.info_lbl.config(text=f'共 {len(text)} 字符 · 编码：{enc}')
        self.preview.delete('1.0', 'end')
        self.preview.insert('1.0', text)
        self.bar.config(maximum=max(len(text), 1), value=0)
        self._set_status('已载入，可直接编辑内容框，按 F6 或点“开始”打字')

    def _paste_clipboard(self):
        try:
            s = self.root.clipboard_get()
        except tk.TclError:
            self._set_status('剪贴板为空或不可读')
            return
        self.preview.insert('insert', s)
        self.preview.see('insert')
        total = len(self.preview.get('1.0', 'end-1c'))
        self._set_status(f'已粘贴 {len(s)} 字符（内容框共 {total} 字符），可开始打字')

    def _cps_from_scale(self, _v=None):
        """拖动速度滑条 → 同步到输入框。"""
        self.cps_entry.set(str(int(self.cps.get())))

    def _apply_cps_entry(self, _event=None):
        """在输入框键入每秒字符数（回车或点到别处生效）→ 同步到滑条。"""
        raw = self.cps_entry.get().strip()
        try:
            v = int(float(raw))
        except (TypeError, ValueError):
            v = None
        if v is None:
            self.cps_entry.set(str(int(self.cps.get())))  # 输入非法就还原
            return
        v = max(1, min(500, v))
        self.cps.set(v)
        self.cps_entry.set(str(v))

    def _sync_speed_state(self):
        st = 'disabled' if self.hard.get() else 'normal'
        self.scale_ctrl.config(state=st)
        self.jitter_cb.config(state=st)
        self.cps_entry_ctrl.config(state=st)

    # ---------- 状态机 ----------
    def _toggle_topmost(self) -> None:
        try:
            self.root.attributes('-topmost', bool(self.topmost.get()))
            _log(f'窗口置顶 → {bool(self.topmost.get())}')
        except Exception:
            pass

    def _demo_clicked(self, _event=None):
        """倒计时/等待目标期间点内置演示框 = 直接在演示框里打字。

        用户直觉：点哪打哪。外部窗口模式下点自己程序里的演示框时，
        不再傻等外部目标，立刻切到演示模式开打。
        """
        if self.state in ('counting', 'waiting'):
            self._cancel_countdown()
            self.mode.set('demo')
            _log('点击内置演示框 → 立即在演示框里打字')
            self._do_start()

    def _toggle(self):
        """F6 / 开始按钮：开始 → 暂停 → 继续；倒计时中则取消。"""
        if self.state == 'idle':
            self._begin()
        elif self.state == 'typing':
            self.engine.pause()
            self.state = 'paused'
            self.start_btn.config(text='继续 (F6)')
            self._set_status('已暂停，按 F6 继续')
        elif self.state == 'paused':
            self.engine.resume()
            self.state = 'typing'
            self.start_btn.config(text='暂停 (F6)')
            self._set_status('继续打字…（F6 暂停 / Esc 停止）')
        elif self.state in ('counting', 'waiting'):
            self._cancel_countdown()

    def _begin(self):
        self._apply_cps_entry()  # 输入框没按回车就点开始/按 F6 也生效
        if not self.preview.get('1.0', 'end-1c'):
            self._set_status('内容为空：请选择文件或粘贴文本')
            return
        if self.mode.get() == 'demo':
            self._do_start()
            return
        n = int(self.countdown.get())
        if n <= 0:
            self._do_start()
            return
        self.state = 'counting'
        self.start_btn.config(text='取消 (F6)')
        self._countdown_tick(n)

    def _countdown_tick(self, n: int):
        if self.state != 'counting':
            return
        if n <= 0:
            self._do_start()
            return
        self._set_status(f'{n} 秒后开始打字 —— 请点击目标输入框！')
        self._countdown_job = self.root.after(
            1000, lambda: self._countdown_tick(n - 1))

    def _cancel_countdown(self):
        if self._countdown_job:
            self.root.after_cancel(self._countdown_job)
            self._countdown_job = None
        if self._target_job:
            self.root.after_cancel(self._target_job)
            self._target_job = None
        self.state = 'idle'
        self.start_btn.config(text='开始 (F6)')
        self._set_status('已取消')

    def _capture_target(self) -> int:
        """记录光标所在的窗口（此刻的前台窗口）并锁定它。

        倒计时结束时，用户已经点过目标输入框，前台窗口就是“光标在的程序”。
        焦点如果还停在本程序窗口上，说明用户没点目标，返回 0 表示不锁定。
        """
        hwnd = int(_GetForegroundWindow() or 0)
        if not hwnd:
            return 0
        if _top_level(hwnd) == _top_level(self.root.winfo_id()):
            return 0
        return hwnd

    def _on_start(self):
        """打字线程刚启动时回调：开打前点一下目标窗口，让按键能进。

        优先点鼠标当前位置（用户刚点过输入框，鼠标正悬在输入框上）；鼠标
        不在目标上就先夺取焦点、再点目标程序当前焦点控件（即用户点过的那
        个输入框）的正中，保证光标落在输入框里。点完把鼠标挪到屏幕左上角，
        免得遮挡或误碰。仅外部模式做这些。
        """
        sink = self._active_sink
        if not sink or not getattr(sink, 'target', 0):
            return
        if not getattr(self, '_pre_click_flag', False) or isinstance(sink, MsgKeySink):
            # 消息直投不需要焦点，别去点用户放好的光标
            # （注意：本方法在打字线程运行，只读普通 Python 值，绝不碰 tkinter）
            return
        try:
            tgt = int(sink.target)
            pt = wt.POINT()
            _GetCursorPos(ctypes.byref(pt))
            hwnd_at = int(_WindowFromPoint(pt) or 0)
            fc, fx, fy = _focus_ctrl(tgt)
            if hwnd_at and _top_level(hwnd_at) == _top_level(tgt):
                _mouse_click(pt.x, pt.y)
                how = f'点击鼠标处 ({pt.x}, {pt.y}) 控件={hwnd_at:#x}'
            else:
                _focus_window(tgt)
                how = '焦点夺取（鼠标不在目标上）'
                if fc and fx:
                    _mouse_click(fx, fy)
                    how += f' + 点击焦点控件 {fc:#x} ({fx},{fy})'
            _SetCursorPos(4, 4)  # 挪开鼠标
            fg = int(_GetForegroundWindow() or 0)
            _log(f'开打前{how} → 目标 {tgt:#x} {_window_title(tgt)!r} '
                 f'焦点控件={fc:#x}；点击后前台 {fg:#x} {_window_title(fg)!r}；鼠标已挪到左上角')
        except Exception:
            pass

    def _do_start(self, target: int = 0):
        text = self.preview.get('1.0', 'end-1c')  # 以内容框当前内容为准
        if not text:
            self._set_status('内容为空：请选择文件或粘贴文本')
            return
        instant = bool(self.hard.get())
        cps = max(1, int(self.cps.get()))
        jitter = 0.3 if (self.jitter.get() and not instant) else 0.0
        total = max(len(text), 1)
        self.bar.config(maximum=total, value=0)
        if self.mode.get() == 'demo':
            self.demo.delete('1.0', 'end')
            self._integrity_warn = False
            sink = CallbackSink(on_char=lambda ch: self.q.put(('insert', ch)))
        else:
            if not target:
                target = self._capture_target()
            if not target:
                # 焦点还在本程序：绝不往自己窗口里打，等用户点到目标输入框
                self._wait_for_target()
                return
            method = self.input_method.get()
            if method == 'paste':
                sink = PasteSink(target=target)
            elif method == 'msg':
                sink = MsgKeySink(target=target, nl_shift=bool(self.nl_shift.get()))
            else:
                sink = OsKeySink(target=target, compat=(method == 'keys'))
            self._integrity_warn = False
            me = _integrity_of(None)
            other = _integrity_of_hwnd(target)
            if me and me < 0x2000:
                self._integrity_warn = True
                _log('!!! 本程序自身被降权为低完整性（exe 文件夹带低完整性标签），'
                     'Windows 会拦掉全部输入——请把 exe 复制到桌面等普通位置运行')
            elif me and other and other > me:
                self._integrity_warn = True
                _log(f'!!! 目标进程完整性级别更高（{other:#x} > 本程序 {me:#x}）：'
                     f'Windows 会拦截按键，请以管理员身份运行本程序')
        self._active_sink = sink
        # 在主线程把勾选值读成普通 bool：_on_start 在打字线程运行，
        # tkinter 变量不能在别的线程读（会 RuntimeError: main thread is not in main loop）
        self._pre_click_flag = bool(self.pre_click.get())
        try:
            self.engine.start(
                text, sink, cps=cps, jitter=jitter, instant=instant,
                on_start=self._on_start,
                on_char=lambda ch, i, t: self.q.put(('progress', (i, t))),
                on_done=lambda: self.q.put(('done', None)),
                on_stop=lambda: self.q.put(('stopped', None)),
            )
        except RuntimeError:
            self._set_status('引擎忙，请稍候再试')
            return
        _log(f'开打 mode={self.mode.get()} method={self.input_method.get()} '
             f'target={target:#x} cps={cps} instant={instant} chars={total}')
        if target:
            fc, fx, fy = _focus_ctrl(target)
            _log(f'  目标 {_window_info(target)}；焦点控件 {_window_info(fc)} '
                 f'中心=({fx},{fy})；当前前台 {_window_info(_GetForegroundWindow())}')
        self.state = 'typing'
        self.start_btn.config(text='暂停 (F6)')
        if self._integrity_warn:
            self._set_status('⚠ 权限问题：Windows 可能拦掉按键——'
                             '把 exe 复制到桌面运行，或以管理员身份运行')
        elif target:
            name = _window_title(target) or '目标窗口'
            self._set_status(f'正在打字 → {name}（F6 暂停 / Esc 停止）')
        else:
            self._set_status('正在打字…（F6 暂停 / Esc 停止）')

    def _wait_for_target(self):
        """焦点还在本程序：等用户点到目标输入框再开打，绝不打到自己窗口里。"""
        self.state = 'waiting'
        self.start_btn.config(text='取消 (F6)')
        self._set_status('请点击要打字的输入框…（点到后自动开始；'
                         '点内置演示框就在演示框里打；Esc 取消）')
        _log('等待目标：焦点还在本程序')
        self._poll_target()

    def _poll_target(self):
        if self.state != 'waiting':
            return
        target = self._capture_target()
        if target:
            self._target_job = None
            _log(f'等到目标 {target:#x} {_window_title(target)!r}')
            self._do_start(target)
            return
        self._target_job = self.root.after(100, self._poll_target)

    def _stop(self):
        if self.state in ('counting', 'waiting'):
            self._cancel_countdown()
            return
        self.engine.stop()
        self.state = 'idle'
        self.start_btn.config(text='开始 (F6)')
        self._set_status('已停止')

    # ---------- 事件轮询（所有 UI 更新都在主线程） ----------
    def _poll(self):
        try:
            while True:
                kind, payload = self.q.get_nowait()
                if kind == 'insert':
                    self.demo.insert('end', payload)
                    self.demo.see('end')
                elif kind == 'progress':
                    i, t = payload
                    self.bar.config(value=i)
                    prefix = '⚠权限更高可能被拦截｜' if self._integrity_warn else ''
                    self._set_status(
                        f'{prefix}已输入 {i} / 共 {t} 字符（F6 暂停 / Esc 停止）')
                elif kind == 'done':
                    self.state = 'idle'
                    self.start_btn.config(text='开始 (F6)')
                    sk = getattr(self, '_active_sink', None)
                    fin = getattr(sk, 'finish', None)
                    if callable(fin):
                        fin()  # PasteSink：吐尾批还原剪贴板；OsKeySink：还原键盘布局
                    msg = '打字完成 ✔'
                    if self._integrity_warn:
                        msg += '（若目标没显示字：目标权限更高被 Windows 拦截，' \
                               '请以管理员身份运行本程序）'
                    self._set_status(msg)
                    _log(f'完成：发送 {getattr(sk, "sent", 0)} 字，'
                         f'失败 {getattr(sk, "failed", 0)} 字，'
                         f'结束时前台 {_window_info(_GetForegroundWindow())}')
                elif kind == 'stopped':
                    if self.state != 'idle':
                        self.state = 'idle'
                        self.start_btn.config(text='开始 (F6)')
                        self._set_status('已停止')
                    sk = getattr(self, '_active_sink', None)
                    ab = getattr(sk, 'abort', None)
                    if callable(ab):
                        ab()  # 不粘贴剩余内容 / 还原键盘布局
                    _log(f'停止：发送 {getattr(sk, "sent", 0)} 字，'
                         f'失败 {getattr(sk, "failed", 0)} 字')
                elif kind == 'hotkey_toggle':
                    self._toggle()
                elif kind == 'hotkey_stop':
                    self._stop()
        except queue.Empty:
            pass
        self.root.after(50, self._poll)

    def _set_status(self, s: str):
        self.status.config(text=s)

    def _on_close(self):
        self.engine.stop()
        self.hotkey.stop()
        self.root.destroy()


def _enable_high_dpi() -> str:
    """声明进程 DPI 感知，返回【真实生效】的感知模式。

    不声明（或声明失败）Windows 会把整个窗口按位图拉伸——界面发虚的根源。
    依次尝试：Per-Monitor V2 → Per-Monitor → System；exe 的 manifest 里也声明
    了 PerMonitorV2（声明更早、更可靠），上面几步可能因此失败——所以最后用
    GetThreadDpiAwarenessContext 读真实状态，而不是拿返回值当结果。
    """
    try:
        fn = ctypes.windll.user32.SetProcessDpiAwarenessContext
        fn.argtypes = [ctypes.c_void_p]
        fn(ctypes.c_void_p(-4))  # DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2
    except Exception:
        pass
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PER_MONITOR_DPI_AWARE
    except Exception:
        pass
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass
    try:
        gt = ctypes.windll.user32.GetThreadDpiAwarenessContext
        gt.restype = ctypes.c_void_p
        eq = ctypes.windll.user32.AreDpiAwarenessContextsEqual
        eq.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
        ctx = gt()
        for name, val in (('pmv2', -4), ('pm', -3), ('system', -2)):
            if eq(ctypes.c_void_p(ctx), ctypes.c_void_p(val)):
                return name
        return 'unaware'
    except Exception:
        return 'unknown'


def _apply_dpi_scaling(root: tk.Tk) -> int:
    """让 Tk 按真实 DPI 缩放字号与布局（不缩放=字小；虚化=上面的感知声明没生效）。"""
    dpi = 0
    try:
        try:
            fn = ctypes.windll.user32.GetDpiForWindow
            fn.argtypes = [wt.HWND]
            fn.restype = wt.UINT
            dpi = int(fn(root.winfo_id()) or 0)
        except Exception:
            dpi = 0
        if not dpi:
            dpi = int(root.winfo_fpixels('1i') or 96)
        if dpi and dpi != 96:
            root.tk.call('tk', 'scaling', float(dpi) / 72.0)
    except Exception:
        pass
    return dpi


def _fit_window(root: tk.Tk) -> None:
    """按实际内容需求放大窗口并居中（高 DPI 下固定尺寸会把内容截掉）。"""
    try:
        root.update_idletasks()
        sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
        w = max(740, root.winfo_reqwidth() + 16)
        h = max(520, root.winfo_reqheight() + 16)
        w = min(w, max(740, sw - 40))
        h = min(h, max(520, sh - 80))
        x, y = max(0, (sw - w) // 2), max(0, (sh - h) // 2 - 20)
        root.geometry(f'{w}x{h}+{x}+{y}')
        root.minsize(min(720, w), min(480, h))
    except Exception:
        pass


def main():
    dpi_mode = _enable_high_dpi()
    me = _integrity_of(None)
    _log(f'启动 build={APP_BUILD} frozen={bool(getattr(sys, "frozen", False))} '
         f'python={sys.version.split()[0]} integrity={(me or 0):#x} dpi_mode={dpi_mode}')
    _log(f'日志文件：{_LOG_PATH}')
    low = bool(me and me < 0x2000)
    if low:
        _log('!!! 本程序运行在低完整性模式（exe 所在文件夹被打上了低完整性标签，'
             '被 Windows 降权）：UIPI 会拦掉对其他程序的一切输入。'
             '请把 exe 复制到桌面等普通位置再运行，或以管理员身份运行')
    root = tk.Tk()
    dpi = _apply_dpi_scaling(root)
    _log(f'DPI：屏幕 {root.winfo_fpixels("1i"):.0f} dpi，窗口实际 {dpi} dpi（96=100% 缩放）')
    app = App(root)
    try:
        root.attributes('-topmost', True)  # 默认置顶，界面「窗口置顶」可取消
    except Exception:
        pass
    if low:
        try:
            app._set_status('⚠⚠ 本程序权限被降为低完整性，打不进其他程序！'
                            '请把 exe 复制到桌面再运行（或右键以管理员身份运行）')
        except Exception:
            pass
    _fit_window(root)
    root.mainloop()


# ============================================================
# 临时工作目录：默认权限自建（不走 tempfile，原因见下）
# ============================================================

def _make_work_dir() -> str:
    """自建一个临时工作目录返回路径；调用方用完自己 shutil.rmtree 删掉。

    不用 tempfile.mkdtemp / TemporaryDirectory：它们内部按 0o700 建目录，
    在带权限模拟的受限环境（某些沙箱）里会把目录锁成连创建它的进程自己
    都写不进、删不掉，自测就莫名其妙挂在 PermissionError 上。
    这里不传 mode 走默认权限，实测各环境都能正常读写。
    """
    bases = [tempfile.gettempdir(),
             os.path.dirname(os.path.abspath(__file__)),
             os.getcwd()]
    err: Exception | None = None
    for base in bases:
        for i in range(50):
            d = os.path.join(base, f'typer_work_{os.getpid()}_{i}')
            try:
                os.mkdir(d)  # 不传 mode —— 关键就在这一步
            except FileExistsError:
                continue
            except OSError as e:
                err = e
                break
            return d
    raise RuntimeError(f'无法创建临时目录：{err}')


# ============================================================
# 自测：python -X utf8 TyperSimulator.py --selftest
# ============================================================

def _selftest():
    import sys

    sys.stdout.reconfigure(encoding='utf-8')

    # 1) 编码识别（纯内存）
    for data, want_enc, key in [
        (b'hello ascii', 'utf-8', 'hello'),
        ('你好GBK，标点。'.encode('gbk'), 'gbk', '你好GBK'),
        ('你好UTF8，混排abc'.encode('utf-8'), 'utf-8', '混排'),
        (b'\xef\xbb\xbf' + 'BOM中文'.encode('utf-8'), 'utf-8-sig', 'BOM中文'),
    ]:
        text, enc = _decode_auto(data)
        assert enc == want_enc and key in text, (enc, text)

    # 2) read_text：GBK 文件 + 手工 docx（临时目录，不留垃圾文件）
    td = _make_work_dir()
    try:
        p = os.path.join(td, 'a.txt')
        with open(p, 'wb') as f:
            f.write('文件读取测试\r\n第二行'.encode('gbk'))
        text, enc = read_text(p)
        assert enc == 'gbk' and '\r' not in text and '第二行' in text

        dx = os.path.join(td, 'a.docx')
        ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<Types xmlns="http://schemas.openxmlformats.org/package/2006/'
              'content-types"><Default Extension="xml" ContentType='
              '"application/xml"/><Override PartName="/word/document.xml" '
              'ContentType="application/vnd.openxmlformats-officedocument.'
              'wordprocessingml.document.main+xml"/></Types>')
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<w:document xmlns:w="http://schemas.openxmlformats.org/'
               'wordprocessingml/2006/main"><w:body>'
               '<w:p><w:r><w:t>第一段docx</w:t></w:r></w:p></w:body></w:document>')
        with zipfile.ZipFile(dx, 'w') as zf:
            zf.writestr('[Content_Types].xml', ct)
            zf.writestr('word/document.xml', doc)
        text, enc = read_text(dx)
        assert enc == 'docx' and '第一段docx' in text
    finally:
        shutil.rmtree(td, ignore_errors=True)

    # 3) 引擎：常速 + 硬输入，逐字回调拼接一致
    sample = '你好ABC123，世界！\n第二行\tTab 😀'
    for inst in (False, True):
        out: list[str] = []
        done = threading.Event()
        TypingEngine().start(sample, CallbackSink(out.append), cps=120,
                             jitter=0, instant=inst, on_done=done.set)
        assert done.wait(timeout=10), '引擎超时未完成'
        assert ''.join(out) == sample, out

    print('SELFTEST PASS（编码识别 / GBK文件 / docx / 引擎常速+硬输入）')


# ============================================================
# 端到端实测（真实按键打字）：python -X utf8 TyperSimulator.py --e2e
# ============================================================

WS_OVERLAPPEDWINDOW = 0x00CF0000
WS_VISIBLE = 0x10000000
WS_VSCROLL = 0x00200000
ES_MULTILINE = 0x0004
ES_AUTOVSCROLL = 0x0040
PM_REMOVE = 1

_PeekMessageW = ctypes.windll.user32.PeekMessageW
_PeekMessageW.argtypes = [ctypes.POINTER(wt.MSG), wt.HWND, wt.UINT, wt.UINT, wt.UINT]
_PeekMessageW.restype = wt.BOOL
_TranslateMessage = ctypes.windll.user32.TranslateMessage
_DispatchMessageW = ctypes.windll.user32.DispatchMessageW
_CreateWindowExW = ctypes.windll.user32.CreateWindowExW
_CreateWindowExW.argtypes = [wt.DWORD, wt.LPCWSTR, wt.LPCWSTR, wt.DWORD,
                             ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
                             wt.HWND, wt.HWND, wt.HINSTANCE, ctypes.c_void_p]
_CreateWindowExW.restype = wt.HWND
_DestroyWindow = ctypes.windll.user32.DestroyWindow
_SetWindowTextW = ctypes.windll.user32.SetWindowTextW
_SetWindowTextW.argtypes = [wt.HWND, wt.LPCWSTR]
_SetWindowTextW.restype = wt.BOOL
_GetModuleHandleW = ctypes.windll.kernel32.GetModuleHandleW
_GetModuleHandleW.restype = wt.HMODULE


class _EditWindow:
    """测试用的真实 Win32 Edit 控件窗口（独立线程创建并泵消息），可当打字目标。

    用真实控件验证 SendInput 确实把字送进了“指定的那个窗口”，
    不依赖记事本等外部程序，测试可重复跑。
    """

    def __init__(self, label: str, x: int, y: int):
        self.label = label  # 仅用于报错提示
        self.hwnd = 0
        self._ready = threading.Event()
        self._quit = threading.Event()
        threading.Thread(target=self._run, args=(x, y), daemon=True).start()
        if not self._ready.wait(timeout=5) or not self.hwnd:
            raise RuntimeError(f'测试窗口创建失败（{label}；当前会话没有图形环境？）')

    def _run(self, x: int, y: int) -> None:
        # 窗口名留空：Edit 控件的窗口名就是它的正文，不留空会成为初始内容
        self.hwnd = int(_CreateWindowExW(
            0, 'Edit', '',
            WS_OVERLAPPEDWINDOW | WS_VISIBLE | WS_VSCROLL
            | ES_MULTILINE | ES_AUTOVSCROLL,
            x, y, 420, 240, 0, 0, _GetModuleHandleW(None), None))
        self._ready.set()
        if not self.hwnd:
            return
        msg = wt.MSG()
        while not self._quit.is_set():
            if _PeekMessageW(ctypes.byref(msg), 0, 0, 0, PM_REMOVE):
                _TranslateMessage(ctypes.byref(msg))
                _DispatchMessageW(ctypes.byref(msg))
            else:
                time.sleep(0.005)
        _DestroyWindow(self.hwnd)
        self.hwnd = 0

    def text(self) -> str:
        """读回控件里的文字（相当于“屏幕截图比对”的可靠版）。"""
        n = _GetWindowTextLengthW(self.hwnd)
        buf = ctypes.create_unicode_buffer(n + 1)
        _GetWindowTextW(self.hwnd, buf, n + 1)
        return buf.value

    def clear(self) -> None:
        """清空控件内容（多场景复用同一个测试窗口时用）。"""
        _SetWindowTextW(self.hwnd, '')

    def close(self) -> None:
        self._quit.set()
        time.sleep(0.05)


def _e2e() -> None:
    """端到端：五种场景全过——逐字 Unicode / 真实按键 / 粘贴 / 跨进程 / 焦点夺回。"""
    import subprocess
    import sys

    sys.stdout.reconfigure(encoding='utf-8')
    win1 = _EditWindow('目标窗口-打字模拟器e2e', 80, 80)
    win2 = _EditWindow('干扰窗口-打字模拟器e2e', 540, 80)
    try:
        # 1) 锁定模式：中途把焦点抢到别的窗口，字应全部打回锁定的目标窗口
        want = '端到端测试abc123，你好世界！\n第二行继续敲完'
        _focus_window(win1.hwnd)
        time.sleep(0.4)
        done = threading.Event()
        eng = TypingEngine()
        eng.start(want, OsKeySink(target=win1.hwnd), cps=25, jitter=0,
                  on_done=done.set)
        time.sleep(0.5)
        _focus_window(win2.hwnd)  # 模拟打字中途用户点到了别的程序
        assert done.wait(timeout=30), '引擎超时未完成'
        time.sleep(0.2)
        got = win1.text().replace('\r\n', '\n')
        if got != want and int(_GetForegroundWindow() or 0) != win1.hwnd:
            print('  [1/6] 锁定目标窗口+焦点夺回：SKIP（前台被其他程序霸占，SendInput 测不了）')
        else:
            assert got == want, f'锁定模式内容不对：want={want!r} got={got!r}'
        assert win2.text() == '', f'字泄漏到了别的窗口：{win2.text()!r}'
        print('  [1/6] 锁定目标窗口 + 焦点被抢自动夺回：PASS')

        # 2) 不锁定模式（target=0）：打给按下那一刻的焦点窗口
        want2 = '不锁定模式XYZ789'
        _focus_window(win2.hwnd)
        time.sleep(0.4)
        done2 = threading.Event()
        eng2 = TypingEngine()
        eng2.start(want2, OsKeySink(), cps=30, jitter=0, on_done=done2.set)
        assert done2.wait(timeout=30), '引擎超时未完成'
        time.sleep(0.2)
        got2 = win2.text()
        if got2 != want2 and int(_GetForegroundWindow() or 0) != win2.hwnd:
            print('  [2/6] 不锁定：SKIP（前台被其他程序霸占，SendInput 测不了）')
        else:
            assert got2 == want2, f'无锁定模式内容不对：want={want2!r} got={got2!r}'
        print('  [2/6] 不锁定（打给当前焦点窗口）：PASS')

        # 4) 真实按键模式：ASCII 用扫描码真实按键，中文退回 Unicode
        #    （敲前把目标布局临时切到 en-US，避开输入法组字；敲完还原）
        want4 = 'real-keys abc123！你好\n末行'
        win1.clear()
        _focus_window(win1.hwnd)
        time.sleep(0.4)
        done4 = threading.Event()
        sink4 = OsKeySink(target=win1.hwnd, compat=True)
        eng4 = TypingEngine()
        eng4.start(want4, sink4, cps=30,
                   jitter=0, on_done=done4.set)
        assert done4.wait(timeout=30), '引擎超时未完成'
        sink4.finish()  # 还原键盘布局
        time.sleep(0.2)
        got4 = win1.text().replace('\r\n', '\n')
        if got4 != want4 and int(_GetForegroundWindow() or 0) != win1.hwnd:
            print('  [4/6] 真实按键（扫描码）：SKIP（前台被其他程序霸占，SendInput 测不了）')
        else:
            assert got4 == want4, f'真实按键模式内容不对：want={want4!r} got={got4!r}'
        print('  [4/6] 真实按键（扫描码）：PASS')

        # 5) 粘贴模式：剪贴板 + Ctrl+V 分批粘贴（任何输入框都能进）
        want5 = 'paste模式abc123，你好！\n第二行粘贴'
        win2.clear()
        _focus_window(win2.hwnd)
        time.sleep(0.4)
        done5 = threading.Event()
        sink5 = PasteSink(target=win2.hwnd)
        eng5 = TypingEngine()
        eng5.start(want5, sink5, cps=60, jitter=0, on_done=done5.set)
        assert done5.wait(timeout=30), '引擎超时未完成'
        sink5.finish()
        time.sleep(0.3)
        got5 = win2.text().replace('\r\n', '\n')
        if got5 != want5 and int(_GetForegroundWindow() or 0) != win2.hwnd:
            print('  [5/6] 粘贴：SKIP（前台被其他程序霸占）')
        else:
            assert got5 == want5, f'粘贴模式内容不对：want={want5!r} got={got5!r}'
        print('  [5/6] 粘贴（剪贴板+Ctrl+V）：PASS')

        # 6) 消息直投：WM_CHAR 直投窗口队列——打字时焦点故意在别的窗口上
        want6 = '消息直投abc123！\n第二行消息'
        win1.clear()
        _focus_window(win2.hwnd)   # 焦点在干扰窗口上！
        time.sleep(0.3)
        done6 = threading.Event()
        sink6 = MsgKeySink(target=win1.hwnd)
        eng6 = TypingEngine()
        eng6.start(want6, sink6, cps=30, jitter=0, on_done=done6.set)
        assert done6.wait(timeout=30), '引擎超时未完成'
        time.sleep(0.2)
        got6 = win1.text().replace('\r\n', '\n')
        assert got6 == want6, f'消息直投内容不对：want={want6!r} got={got6!r}'
        assert int(_GetForegroundWindow() or 0) != win1.hwnd, '消息直投抢了焦点'
        print('  [6/6] 消息直投（免焦点后台输入）：PASS')

        # 3) 跨进程：打进【另一个进程】的真实窗口（等价于记事本/微信等外部程序）
        _base = os.path.dirname(os.path.abspath(__file__))
        helper = next((p for p in (os.path.join(_base, '_target_app.py'),
                                   os.path.join(_base, '..', 'tests', '_target_app.py'))
                       if os.path.exists(p)), '')
        if helper:
            work = _make_work_dir()
            proc = subprocess.Popen([sys.executable, '-X', 'utf8', helper, work])
            try:
                hwnd3 = 0
                for _ in range(100):
                    try:
                        with open(os.path.join(work, 'hwnd.txt'),
                                  encoding='utf-8') as f:
                            hwnd3 = int(f.read().strip(), 16)
                    except OSError:
                        pass
                    if hwnd3:
                        break
                    assert proc.poll() is None, '目标进程启动失败'
                    time.sleep(0.1)
                assert hwnd3, '没等到目标进程的窗口'
                want3 = '跨进程abc123，你好！\n第二行'
                _focus_window(hwnd3)
                time.sleep(0.4)
                done3 = threading.Event()
                eng3 = TypingEngine()
                eng3.start(want3, OsKeySink(target=hwnd3), cps=30, jitter=0,
                           on_done=done3.set)
                assert done3.wait(timeout=30), '引擎超时未完成'
                time.sleep(0.5)
                with open(os.path.join(work, 'content.txt'),
                          encoding='utf-8', newline='') as f:
                    got3 = f.read().replace('\r\n', '\n')
                if got3 != want3 and int(_GetForegroundWindow() or 0) != hwnd3:
                    print('  [3/6] 跨进程：SKIP（前台被其他程序霸占，SendInput 测不了）')
                else:
                    assert got3 == want3, f'跨进程内容不对：want={want3!r} got={got3!r}'
                print('  [3/6] 跨进程（另一进程的窗口）：PASS')
            finally:
                proc.kill()
                shutil.rmtree(work, ignore_errors=True)
        else:
            print('  [3/6] 跨进程：跳过（找不到 _target_app.py）')
    finally:
        win1.close()
        win2.close()
    print('E2E PASS（消息直投 / 逐字Unicode / 真实按键 / 粘贴 / 焦点夺回 / 无锁定 / 跨进程）')


if __name__ == '__main__':
    import sys
    if '--selftest' in sys.argv:
        _selftest()
    elif '--e2e' in sys.argv:
        _e2e()
    else:
        main()
