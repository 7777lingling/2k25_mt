import cv2
import numpy as np
import win32gui
import win32ui
import win32con
import win32api
import time
from pathlib import Path
import win32process
from PIL import ImageGrab
from ctypes import windll
import pyautogui  # 添加 pyautogui 導入

def get_window_screenshot(hwnd):
    try:
        # 獲取窗口位置和大小
        rect = win32gui.GetWindowRect(hwnd)
        left, top, right, bottom = rect
        width = right - left
        height = bottom - top
        
        print(f"視窗大小: 寬度={width}, 高度={height}")
        
        # 使用 pyautogui 截圖
        screenshot = pyautogui.screenshot(region=(left, top, width, height))
        # 轉換為 OpenCV 格式 (BGR)
        screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
        
        # 保存截圖用於調試
        cv2.imwrite("debug_screenshot.png", screenshot)
        print(f"已保存截圖，尺寸: {screenshot.shape}")
        
        return screenshot
        
    except Exception as e:
        print(f"截圖時發生錯誤: {e}")
        import traceback
        traceback.print_exc()
        return None

# 另一種方法：使用 OpenCV 的方式（備用）
def get_window_screenshot_cv2(hwnd):
    try:
        # 獲取窗口位置和大小
        rect = win32gui.GetWindowRect(hwnd)
        left, top, right, bottom = rect
        width = right - left
        height = bottom - top
        
        print(f"視窗大小: 寬度={width}, 高度={height}")
        
        # 使用 OpenCV 截圖
        import mss
        with mss.mss() as sct:
            monitor = {"top": top, "left": left, "width": width, "height": height}
            screenshot = np.array(sct.grab(monitor))
            # 轉換顏色空間從 BGRA 到 BGR
            screenshot = cv2.cvtColor(screenshot, cv2.COLOR_BGRA2BGR)
        
        # 保存截圖用於調試
        cv2.imwrite("debug_screenshot.png", screenshot)
        print(f"已保存截圖，尺寸: {screenshot.shape}")
        
        return screenshot
        
    except Exception as e:
        print(f"截圖時發生錯誤: {e}")
        import traceback
        traceback.print_exc()
        return None

def set_foreground_window(hwnd):
    # 獲取當前前台窗口
    current_hwnd = win32gui.GetForegroundWindow()
    
    # 獲取當前窗口的線程ID
    current_thread = win32process.GetWindowThreadProcessId(current_hwnd)[0]
    # 獲取目標窗口的線程ID
    target_thread = win32process.GetWindowThreadProcessId(hwnd)[0]
    # 獲取當前線程ID
    current_process = win32api.GetCurrentThreadId()
    
    # 連接線程輸入
    win32process.AttachThreadInput(target_thread, current_process, True)
    win32process.AttachThreadInput(current_thread, current_process, True)
    
    try:
        # 顯示窗口
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        else:
            win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
        
        # 設置窗口位置
        win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, 
                            win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
        win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0, 
                            win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
        
        # 設置前台窗口
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.1)
        return True
        
    except Exception as e:
        print(f"設置前台窗口時出錯：{e}")
        return False
        
    finally:
        # 斷開線程輸入連接
        win32process.AttachThreadInput(target_thread, current_process, False)
        win32process.AttachThreadInput(current_thread, current_process, False)

def detect_image(hwnd, template_path, threshold=0.5):
    if not Path(template_path).exists():
        print(f"找不到圖片：{template_path}")
        return False
        
    # 讀取模板圖片
    template = cv2.imread(str(template_path))
    if template is None:
        print(f"無法讀取圖片：{template_path}")
        return False
        
    # 嘗試將窗口置於前台
    if not set_foreground_window(hwnd):
        print("無法將窗口置於前台")
        
    # 獲取截圖
    screenshot_cv = get_window_screenshot(hwnd)
    if screenshot_cv is None:
        print("截圖失敗")
        return False
        
    print(f"截圖尺寸: {screenshot_cv.shape}, 模板尺寸: {template.shape}")
    
    # 計算解析度比例
    screen_height = screenshot_cv.shape[0]
    template_height = template.shape[0]
    scale_ratio = screen_height / 1080  # 假設模板是在1080p下截取的
    
    # 根據解析度調整模板大小
    if abs(scale_ratio - 1.0) > 0.1:  # 如果解析度差異超過10%
        new_width = int(template.shape[1] * scale_ratio)
        new_height = int(template.shape[0] * scale_ratio)
        template = cv2.resize(template, (new_width, new_height))
        print(f"調整模板尺寸至: {template.shape} (縮放比例: {scale_ratio:.2f})")
    
    # 圖像預處理
    screenshot_gray = cv2.cvtColor(screenshot_cv, cv2.COLOR_BGR2GRAY)
    template_gray = cv2.cvtColor(template, cv2.COLOR_BGR2GRAY)
    
    # 使用不同的匹配方法
    methods = [cv2.TM_CCOEFF_NORMED, cv2.TM_CCORR_NORMED]
    max_val = 0
    max_loc = None
    best_method = None
    
    for method in methods:
        result = cv2.matchTemplate(screenshot_gray, template_gray, method)
        _, val, _, loc = cv2.minMaxLoc(result)
        if val > max_val:
            max_val = val
            max_loc = loc
            best_method = method
    
    # 顯示結果
    match_img = screenshot_cv.copy()
    cv2.rectangle(match_img, (0, 0), 
                 (screenshot_cv.shape[1]-1, screenshot_cv.shape[0]-1),
                 (0, 255, 0), 2)
    
    if max_val >= threshold:
        width = template.shape[1]
        height = template.shape[0]
        cv2.rectangle(match_img, max_loc, 
                     (max_loc[0] + width, max_loc[1] + height),
                     (0, 0, 255), 2)
        
        # 顯示匹配區域的放大圖
        x, y = max_loc
        roi = screenshot_cv[y:y+height, x:x+width]
        if roi.size > 0:
            roi_resized = cv2.resize(roi, (roi.shape[1]*2, roi.shape[0]*2))
            cv2.imshow("匹配區域", roi_resized)
            
            # 同時顯示調整後的模板以供比較
            template_resized = cv2.resize(template, (template.shape[1]*2, template.shape[0]*2))
            cv2.imshow("模板圖片", template_resized)
    
    cv2.imshow("匹配結果", match_img)
    cv2.waitKey(1)
    
    method_names = {
        cv2.TM_CCOEFF_NORMED: "TM_CCOEFF_NORMED",
        cv2.TM_CCORR_NORMED: "TM_CCORR_NORMED"
    }
    
    print(f"最佳匹配度: {max_val:.3f}")
    print(f"使用匹配方法: {method_names[best_method]}")
    print(f"匹配位置: {max_loc}")
    
    return max_val >= threshold

if __name__ == "__main__":
    # 定義圖片路徑
    base_path = Path('.')
    folder1 = base_path / "image1"  # 遊戲主要流程圖片
    
    # 測試兩個三星圖片
    paths = {
        "stars": str(folder1 / "stars.png"),
        "stars2": str(folder1 / "stars2.png")
    }
    
    window_name = "NBA 2K25"
    print(f"尋找遊戲視窗: {window_name}")
    hwnd = win32gui.FindWindow(None, window_name)
    
    if not hwnd:
        print("未找到遊戲視窗")
        exit(1)
    
    print(f"找到遊戲視窗，句柄: {hwnd}")
    print("操作說明:")
    print("- 按 'q' 鍵退出程式")
    print("- 按 'up' 鍵增加閥值 (+0.001)")
    print("- 按 'down' 鍵減少閥值 (-0.001)")
    print("- 按 'r' 鍵重置閥值為 0.990")
    
    # 設置閾值為 0.990
    threshold = 0.990
    
    try:
        while True:
            # 測試當前畫面
            print("\n" + "="*50)
            print(f"當前閥值: {threshold:.3f}")
            
            # 獲取截圖
            screenshot = get_window_screenshot(hwnd)
            if screenshot is None:
                print("截圖失敗")
                continue
            
            # 初始化最佳匹配結果
            best_match = {
                "image": None,
                "value": 0,
                "method": None,
                "loc": None
            }
            
            # 檢測所有圖片
            for image_name, path in paths.items():
                # 讀取模板圖片
                template = cv2.imread(path)
                if template is None:
                    print(f"無法讀取圖片：{path}")
                    continue
                
                # 圖像預處理
                screenshot_gray = cv2.cvtColor(screenshot, cv2.COLOR_BGR2GRAY)
                template_gray = cv2.cvtColor(template, cv2.COLOR_BGR2GRAY)
                
                # 使用不同的匹配方法
                methods = [cv2.TM_CCOEFF_NORMED, cv2.TM_CCORR_NORMED]
                for method in methods:
                    result = cv2.matchTemplate(screenshot_gray, template_gray, method)
                    _, val, _, loc = cv2.minMaxLoc(result)
                    if val > best_match["value"]:
                        best_match = {
                            "image": image_name,
                            "value": val,
                            "method": method,
                            "loc": loc
                        }
            
            # 顯示最佳匹配結果
            print(f"最佳匹配圖片: {best_match['image']}")
            print(f"最佳匹配值: {best_match['value']:.3f}")
            print(f"是否匹配: {'是' if best_match['value'] >= threshold else '否'}")
            
            # 如果有匹配結果，顯示匹配區域
            if best_match["value"] >= threshold:
                template = cv2.imread(paths[best_match["image"]])
                match_img = screenshot.copy()
                width = template.shape[1]
                height = template.shape[0]
                cv2.rectangle(match_img, best_match["loc"], 
                            (best_match["loc"][0] + width, best_match["loc"][1] + height),
                            (0, 0, 255), 2)
                
                # 顯示匹配區域的放大圖
                x, y = best_match["loc"]
                roi = screenshot[y:y+height, x:x+width]
                if roi.size > 0:
                    roi_resized = cv2.resize(roi, (roi.shape[1]*2, roi.shape[0]*2))
                    cv2.imshow("匹配區域", roi_resized)
                    
                    # 同時顯示調整後的模板以供比較
                    template_resized = cv2.resize(template, (template.shape[1]*2, template.shape[0]*2))
                    cv2.imshow("模板圖片", template_resized)
                
                cv2.imshow("匹配結果", match_img)
            
            # 檢查按鍵
            key = cv2.waitKey(1) & 0xFF
            if key == ord('q'):
                print("\n收到退出信號，程式結束")
                break
            elif key == 82:  # up arrow
                threshold = min(1.0, threshold + 0.001)
                print(f"提高閥值至: {threshold:.3f}")
            elif key == 84:  # down arrow
                threshold = max(0.0, threshold - 0.001)
                print(f"降低閥值至: {threshold:.3f}")
            elif key == ord('r'):
                threshold = 0.990
                print(f"重置閥值至: {threshold:.3f}")
            
            # 等待一段時間再進行下一次檢測
            time.sleep(0.5)
            
    except KeyboardInterrupt:
        print("\n使用者中斷程式")
    finally:
        cv2.destroyAllWindows()
        print("程式已結束") 