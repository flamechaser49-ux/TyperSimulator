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

INPUT_KEYBOARD = 1
KEYEVENTF_KEYUP = 0x0002
KEYEVENTF_UNICODE = 0x0004
VK_TAB = 0x09
VK_RETURN = 0x0D
VK_ESCAPE = 0x1B
VK_F6 = 0x75


class MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ('dx', wt.LONG), ('dy', wt.LONG),
        ('mouseData', wt.DWORD), ('dwFlags', wt.DWORD), ('time', wt.DWORD),
        ('dwExtraInfo', ctypes.POINTER(ctypes.c_ulong)),
    ]


class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ('wVk', wt.WORD), ('wScan', wt.WORD),
        ('dwFlags', wt.DWORD), ('time', wt.DWORD),
        ('dwExtraInfo', ctypes.POINTER(ctypes.c_ulong)),
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


def _send(ki: KEYBDINPUT) -> None:
    inp = INPUT(type=INPUT_KEYBOARD)
    inp.ki = ki
    _SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT))


def _press_vk(vk: int) -> None:
    _send(KEYBDINPUT(wVk=vk, wScan=0, dwFlags=0, time=0, dwExtraInfo=None))
    _send(KEYBDINPUT(wVk=vk, wScan=0, dwFlags=KEYEVENTF_KEYUP, time=0, dwExtraInfo=None))


def _unicode_key(code: int, up: bool = False) -> None:
    flags = KEYEVENTF_UNICODE | (KEYEVENTF_KEYUP if up else 0)
    _send(KEYBDINPUT(wVk=0, wScan=code, dwFlags=flags, time=0, dwExtraInfo=None))


def _press_char(ch: str) -> None:
    code = ord(ch)
    if code > 0xFFFF:
        # 增补字符（emoji 等）需拆成 UTF-16 高低代理对分别发送
        code -= 0x10000
        _unicode_key(0xD800 + (code >> 10))
        _unicode_key(0xDC00 + (code & 0x3FF))
        _unicode_key(0xDC00 + (code & 0x3FF), up=True)
        _unicode_key(0xD800 + (code >> 10), up=True)
    else:
        _unicode_key(code)
        _unicode_key(code, up=True)


class OsKeySink:
    """真实按键输出：打进当前焦点窗口。"""

    def send(self, ch: str) -> None:
        if ch == '\n':
            _press_vk(VK_RETURN)
        elif ch == '\t':
            _press_vk(VK_TAB)
        else:
            _press_char(ch)


class CallbackSink:
    """演示输出：不碰操作系统，只把字符回调出去。"""

    def __init__(self, on_char):
        self.on_char = on_char

    def send(self, ch: str) -> None:
        self.on_char(ch)


class TypingEngine:
    """后台线程逐字输出，支持暂停 / 继续 / 停止。"""

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
        self.state = 'idle'  # idle / counting / typing / paused

        self.engine = TypingEngine()
        self.q: queue.Queue = queue.Queue()
        self._countdown_job = None

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
        top = ttk.Frame(self.root, padding=(10, 8))
        top.pack(fill='x')
        ttk.Button(top, text='选择文件…', command=self._pick_file).pack(side='left')
        self.file_lbl = ttk.Label(top, text='（未选择文件）', padding=(8, 0))
        self.file_lbl.pack(side='left')
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
        pbtns = ttk.Frame(prev_box)
        pbtns.grid(row=2, column=0, columnspan=2, sticky='ew', pady=(4, 0))
        ttk.Button(pbtns, text='粘贴剪贴板',
                   command=self._paste_clipboard).pack(side='left')
        ttk.Button(pbtns, text='清空',
                   command=lambda: self.preview.delete('1.0', 'end')
                   ).pack(side='left', padx=6)
        prev_box.rowconfigure(0, weight=1)
        prev_box.columnconfigure(0, weight=1)

        demo_box = ttk.LabelFrame(mid, text=' 内置演示框（演示模式在这里打字） ', padding=4)
        demo_box.pack(side='left', fill='both', expand=True, padx=(4, 0))
        self.demo = tk.Text(demo_box, wrap='none', width=40, height=12)
        dv = ttk.Scrollbar(demo_box, orient='vertical', command=self.demo.yview)
        self.demo.configure(yscrollcommand=dv.set)
        self.demo.grid(row=0, column=0, sticky='nsew')
        dv.grid(row=0, column=1, sticky='ns')
        ttk.Button(demo_box, text='清空演示框',
                   command=lambda: self.demo.delete('1.0', 'end')
                   ).grid(row=1, column=0, columnspan=2, sticky='ew', pady=(4, 0))
        demo_box.rowconfigure(0, weight=1)
        demo_box.columnconfigure(0, weight=1)

        setf = ttk.LabelFrame(self.root, text=' 参数 ', padding=(10, 6))
        setf.pack(fill='x', padx=10, pady=(8, 0))

        row1 = ttk.Frame(setf)
        row1.pack(fill='x', pady=(0, 4))
        row2 = ttk.Frame(setf)
        row2.pack(fill='x')

        self.cps = tk.IntVar(value=15)
        self.scale_ctrl = tk.Scale(row1, from_=1, to=500, orient='horizontal',
                                   label='每秒字符数（上限 500）',
                                   variable=self.cps, length=230)
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
        ttk.Radiobutton(row2, text='外部窗口（打到光标点击处）',
                        variable=self.mode, value='external'
                        ).pack(side='left', padx=(0, 6))
        ttk.Radiobutton(row2, text='内置演示框',
                        variable=self.mode, value='demo').pack(side='left')

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

    def _sync_speed_state(self):
        st = 'disabled' if self.hard.get() else 'normal'
        self.scale_ctrl.config(state=st)
        self.jitter_cb.config(state=st)

    # ---------- 状态机 ----------
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
        elif self.state == 'counting':
            self._cancel_countdown()

    def _begin(self):
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
        self.state = 'idle'
        self.start_btn.config(text='开始 (F6)')
        self._set_status('已取消')

    def _do_start(self):
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
            sink = CallbackSink(on_char=lambda ch: self.q.put(('insert', ch)))
        else:
            sink = OsKeySink()
        try:
            self.engine.start(
                text, sink, cps=cps, jitter=jitter, instant=instant,
                on_char=lambda ch, i, t: self.q.put(('progress', (i, t))),
                on_done=lambda: self.q.put(('done', None)),
                on_stop=lambda: self.q.put(('stopped', None)),
            )
        except RuntimeError:
            self._set_status('引擎忙，请稍候再试')
            return
        self.state = 'typing'
        self.start_btn.config(text='暂停 (F6)')
        self._set_status('正在打字…（F6 暂停 / Esc 停止）')

    def _stop(self):
        if self.state == 'counting':
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
                    self._set_status(f'已输入 {i} / 共 {t} 字符（F6 暂停 / Esc 停止）')
                elif kind == 'done':
                    self.state = 'idle'
                    self.start_btn.config(text='开始 (F6)')
                    self._set_status('打字完成 ✔')
                elif kind == 'stopped':
                    if self.state != 'idle':
                        self.state = 'idle'
                        self.start_btn.config(text='开始 (F6)')
                        self._set_status('已停止')
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


def _enable_high_dpi():
    """声明进程 DPI 感知：不声明的话 Windows 会位图拉伸整个窗口，界面发虚。"""
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PER_MONITOR_DPI_AWARE
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def main():
    _enable_high_dpi()
    root = tk.Tk()
    App(root)
    root.mainloop()


# ============================================================
# 自测：python -X utf8 TyperSimulator.py --selftest
# ============================================================

def _selftest():
    import sys
    import tempfile

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
    with tempfile.TemporaryDirectory() as td:
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


if __name__ == '__main__':
    import sys
    if '--selftest' in sys.argv:
        _selftest()
    else:
        main()
