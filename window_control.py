import win32gui
import win32con
import time

def find_and_focus_window(window_name, set_topmost=True):
    """
    查找指定名稱的視窗並控制其置頂狀態
    :param window_name: 視窗名稱
    :param set_topmost: True為置頂，False為取消置頂
    """
    # 查找視窗
    hwnd = win32gui.FindWindow(None, window_name)
    
    if hwnd == 0:
        print(f"未找到名為 '{window_name}' 的視窗")
        return False
    
    print(f"找到視窗 '{window_name}', 句柄: {hwnd}")
    
    # 控制視窗置頂狀態
    if set_topmost:
        win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                            win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
        print("視窗已置頂")
    else:
        win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                            win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
        print("視窗已取消置頂")
    
    return True

if __name__ == "__main__":
    window_name = "NBA 2K25"
    # 設置置頂
    find_and_focus_window(window_name, True)
    # 等待5秒後取消置頂
    time.sleep(5)
    find_and_focus_window(window_name, False) 