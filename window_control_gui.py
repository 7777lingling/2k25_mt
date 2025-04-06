import tkinter as tk
from tkinter import ttk, messagebox
import win32gui
import win32con
import time
import keyboard
import threading
from game_loop import GameLoop

class WindowControlGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("NBA 2K25 視窗控制")
        self.root.geometry("400x400")
        
        # 創建主框架
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 視窗名稱輸入
        ttk.Label(main_frame, text="視窗名稱:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.window_name = tk.StringVar(value="NBA 2K25")
        self.window_name_entry = ttk.Entry(main_frame, textvariable=self.window_name, width=30)
        self.window_name_entry.grid(row=0, column=1, columnspan=2, sticky=tk.W, pady=5)
        
        # 查找按鈕
        ttk.Button(main_frame, text="查找視窗", command=self.find_window).grid(row=1, column=0, columnspan=3, pady=10)
        
        # 置頂控制按鈕
        self.topmost_var = tk.BooleanVar(value=False)
        self.topmost_button = ttk.Checkbutton(main_frame, text="視窗置頂", variable=self.topmost_var, 
                                            command=self.toggle_topmost)
        self.topmost_button.grid(row=2, column=0, columnspan=3, pady=5)
        
        # 稱霸賽功能按鈕
        self.domination_var = tk.BooleanVar(value=False)
        self.domination_button = ttk.Checkbutton(main_frame, text="稱霸賽功能", variable=self.domination_var,
                                               command=self.toggle_domination)
        self.domination_button.grid(row=3, column=0, columnspan=3, pady=5)
        
        # 狀態顯示
        self.status_var = tk.StringVar(value="就緒")
        ttk.Label(main_frame, textvariable=self.status_var).grid(row=4, column=0, columnspan=3, pady=5)
        
        # 視窗句柄顯示
        self.hwnd_var = tk.StringVar(value="視窗句柄: 未找到")
        ttk.Label(main_frame, textvariable=self.hwnd_var).grid(row=5, column=0, columnspan=3, pady=5)
        
        # 自動刷新選項
        self.auto_refresh = tk.BooleanVar(value=False)
        ttk.Checkbutton(main_frame, text="自動刷新置頂", variable=self.auto_refresh, 
                       command=self.toggle_auto_refresh).grid(row=6, column=0, columnspan=3, pady=5)
        
        self.refresh_job = None
        self.current_hwnd = 0
        self.game_loop = None
        self.is_running = False

    def find_window(self):
        window_name = self.window_name.get()
        hwnd = win32gui.FindWindow(None, window_name)
        
        if hwnd == 0:
            self.status_var.set(f"未找到名為 '{window_name}' 的視窗")
            self.hwnd_var.set("視窗句柄: 未找到")
            self.current_hwnd = 0
            messagebox.showerror("錯誤", f"未找到名為 '{window_name}' 的視窗")
            return False
        
        self.status_var.set(f"已找到視窗 '{window_name}'")
        self.hwnd_var.set(f"視窗句柄: {hwnd}")
        self.current_hwnd = hwnd
        return True

    def toggle_topmost(self):
        if not self.current_hwnd:
            if not self.find_window():
                self.topmost_var.set(False)
                return
        
        if self.topmost_var.get():
            # 設置置頂
            win32gui.SetWindowPos(self.current_hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                                win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            self.status_var.set("視窗已置頂")
        else:
            # 取消置頂
            win32gui.SetWindowPos(self.current_hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                                win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            self.status_var.set("視窗已取消置頂")

    def toggle_domination(self):
        if self.domination_var.get():
            if not self.current_hwnd:
                if not self.find_window():
                    self.domination_var.set(False)
                    return
            self.start_domination()
        else:
            self.stop_domination()

    def start_domination(self):
        if not self.current_hwnd:
            if not self.find_window():
                self.domination_var.set(False)
                return
        
        # 設置視窗置頂
        win32gui.SetWindowPos(self.current_hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                            win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
        self.status_var.set("視窗已置頂，稱霸賽功能已啟動")
        self.topmost_var.set(True)
        
        # 創建並啟動GameLoop
        self.game_loop = GameLoop()
        self.is_running = True
        self.domination_thread = threading.Thread(target=self.game_loop.start)
        self.domination_thread.daemon = True
        self.domination_thread.start()

    def stop_domination(self):
        self.is_running = False
        if self.game_loop:
            self.game_loop.stop()
        if self.domination_thread:
            self.domination_thread.join(timeout=1)
        
        # 取消視窗置頂
        if self.current_hwnd:
            win32gui.SetWindowPos(self.current_hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                                win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            self.status_var.set("視窗已取消置頂，稱霸賽功能已停止")
            self.topmost_var.set(False)
        else:
            self.status_var.set("稱霸賽功能已停止")

    def toggle_auto_refresh(self):
        if self.auto_refresh.get():
            self.start_auto_refresh()
        else:
            self.stop_auto_refresh()

    def start_auto_refresh(self):
        self.refresh_job = self.root.after(1000, self.auto_refresh_window)

    def stop_auto_refresh(self):
        if self.refresh_job:
            self.root.after_cancel(self.refresh_job)
            self.refresh_job = None

    def auto_refresh_window(self):
        if self.auto_refresh.get():
            if self.topmost_var.get():
                self.toggle_topmost()
            self.refresh_job = self.root.after(1000, self.auto_refresh_window)

def main():
    root = tk.Tk()
    app = WindowControlGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main() 