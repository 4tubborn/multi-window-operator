import pyperclip
import pygetwindow as gw
import pyautogui
import time
import os
import re
from pynput import mouse
import ctypes
import win32gui

def is_cursor_ibeam():
    """判断当前鼠标光标是否为输入状态 (I-Beam)"""
    # 结构体定义：获取光标信息
    class CURSORINFO(ctypes.Structure):
        _fields_ = [
            ("cbSize", ctypes.c_uint),
            ("flags", ctypes.c_uint),
            ("hCursor", ctypes.c_void_p),
            ("ptScreenPos", ctypes.c_long * 2)
        ]

    # 加载系统预定义的 I-Beam 光标句柄
    # 32513 是 Windows 系统中 IDC_IBEAM 的常量 ID
    IDC_IBEAM = win32gui.LoadCursor(0, 32513)

    info = CURSORINFO()
    info.cbSize = ctypes.sizeof(CURSORINFO)
    
    # 获取当前光标信息
    if ctypes.windll.user32.GetCursorInfo(ctypes.byref(info)):
        if info.flags == 1:  # CURSOR_SHOWING
            return info.hCursor == IDC_IBEAM
    return False

class MouseBlocker:
    def __init__(self):
        self.listener = None
    
    def _on_click(self, x, y, button, pressed):
        return False  # 拦截点击

    def _on_scroll(self, x, y, dx, dy):
        return False  # 拦截滚动
    
    def _on_move(self, x, y):
        return False

    def start(self):
        if self.listener is None:
            self.listener = mouse.Listener(
                on_click=self._on_click,
                on_scroll=self._on_scroll,
                on_move=self._on_move
            )
            self.listener.start()

    def stop(self):
        if self.listener:
            self.listener.stop()
            self.listener.join()  # 确保完全停止
            self.listener = None

# === 配置 ===

CMD_TITLE = "AI Sync Commander"
os.system(f"title {CMD_TITLE}")

TARGET_TITLES = [
    "文心一言",
    "Qwen Chat",
    "Kimi",
    "DeepSeek",
    "Claude",
    "ChatGPT",
    "Gemini",
    "Grok",
    "Copilot",
    "Doubao",
    "豆包",
    "智谱",
    "天工"
]

def try_focus_input_box(win):
    #可能有bug, 如果未选择到任何文本的话会错判为成功
    """尝试点击输入框位置，成功后停止"""
    left, top, width, height = win.left, win.top, win.width, win.height
    
    candidates = [
        (left + width // 2, top + height - 150, "底部-150"),
        (left + width // 2, top + height - 240, "底部-240"),
        (left + width // 2, top + height - 250, "底部-250"),
        (left + width // 2, top + height - 260, "底部-260"),
        (left + width // 2, top + height // 2, "中心"),
        (left + width // 2, top + height // 5, "顶部")        
    ]
    
    # 保存原始剪贴板内容
    test_char = "@(*&+ad"  # 测试
    
    for x, y, desc in candidates:
        print(f"  → 尝试点击 {desc} 位置 ({x}, {y})")
        pyautogui.moveTo(x, y)
        time.sleep(0.05)
        if is_cursor_ibeam():
            #pyautogui.hotkey('ctrl','a')
            #pyautogui.press('backspace')
            pyautogui.click()
            return True
        #time.sleep(0.1)
        

        # 设置测试文本到剪贴板
        """
        pyperclip.copy(test_char)
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.hotkey('ctrl','a')
        pyautogui.hotkey('ctrl','c')
        result=pyperclip.paste()
        if result in test_char:
            pyautogui.press('backspace')
            return True
        else:
            continue
        """
    
    # 全部失败
    print("无法定位至输入框")
    return False

def find_windows():
    """查找所有匹配的窗口"""
    found = []
    for win in gw.getAllWindows():
        if not win.isMinimized and win.title.strip():
            for kw in TARGET_TITLES:
                if kw in win.title:
                    found.append(win)
                    break          
    return found

print(f"\n将在 1 秒后开始操作... 请确保以下窗口已打开：")
for t in TARGET_TITLES:
    print(f" - 标题包含 '{t}'")

time.sleep(1)

# 初始查找
found_windows = find_windows()
if not found_windows:
    print("未找到任何匹配的窗口！")
    #exit()
else:
    print(f"初始找到 {len(found_windows)} 个窗口")
    for w in found_windows:
        print(f"  - {w.title}")


# === 主循环 ===
while True:
    text = input("\n请输入要同步粘贴的文本（支持命令: /r, /q）: ").strip()
    
    if text.lower() == "/q":
        break
        
    elif text.lower() == "/r":
        print("正在重新查找窗口...")
        found_windows = find_windows()
        if found_windows:
            print(f"找到 {len(found_windows)} 个窗口:")
            for w in found_windows:
                print(f"  - {w.title}")
        else:
            print("未找到任何匹配窗口")
        continue

    # 正常粘贴流程
    original_x, original_y = pyautogui.position()
    

    blocker = MouseBlocker()
    blocker.start()
    try:
        for win in found_windows:       
            try:            
                print(f"正在操作: {win.title}")
                win.activate()
                try_focus_input_box(win)
                pyperclip.copy(text)
                time.sleep(0.03)
                #try_focus_input_box(win)
                pyautogui.hotkey('ctrl', 'v')
                pyautogui.press('enter')
            except Exception as e:
                print(f"操作失败: {e}")
    finally:
        blocker.stop()
        pyautogui.moveTo(original_x, original_y)

    # 粘贴完切回 CMD
    try:
        cmd_win = gw.getWindowsWithTitle(CMD_TITLE)
        if cmd_win:
            cmd_win[0].activate()
            time.sleep(0.01)
    except:
        pass

    print("已完成同步粘贴！")
