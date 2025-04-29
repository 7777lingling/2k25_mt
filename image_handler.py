import os
import time
import win32api
import win32con
import win32gui
import cv2
import numpy as np
from pathlib import Path
import logging
from datetime import datetime
import pyautogui
import json

class ImageHandler:
    def __init__(self):
        self.logger = self._setup_logging()
        self.window_name = "NBA 2K25"
        self.hwnd = None
        self.find_window()
        self.is_running = True
        
        # 讀取配置文件
        with open('config.json', 'r') as f:
            self.config = json.load(f)
        
        # 定義按鍵映射
        self.KEYS = {
            "RIGHT": ord('D'),
            "LEFT": ord('A'),
            "SPACE": win32con.VK_SPACE,
            "E": ord('E'),
            "S": ord('S'),
            "W": ord('W'),
            "ESC": win32con.VK_ESCAPE
        }
        
        # 遊戲狀態
        self.state = {
            'in_myteam': False,
            'in_domination': False
        }

    def _setup_logging(self):
        """設置日誌"""
        log_dir = Path("logs")
        log_dir.mkdir(exist_ok=True)
        
        current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = log_dir / f"image_handler_{current_time}.log"
        
        # 為文件使用完整格式（包含時間）
        file_formatter = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s')
        
        # 為控制台使用簡化格式（不含時間）
        console_formatter = logging.Formatter('[%(levelname)s] %(message)s')
        
        # 添加文件處理器
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setFormatter(file_formatter)
        
        # 添加控制台處理器
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(console_formatter)
        
        logger = logging.getLogger(__name__)
        logger.setLevel(logging.INFO)
        logger.addHandler(file_handler)
        logger.addHandler(console_handler)
        
        return logger

    def find_window(self):
        """查找遊戲視窗"""
        self.logger.info(f"正在查找遊戲視窗: {self.window_name}")
        self.hwnd = win32gui.FindWindow(None, self.window_name)
        if not self.hwnd:
            self.logger.error(f"未找到遊戲視窗: {self.window_name}")
            return False
        self.logger.info(f"找到遊戲視窗，句柄: {self.hwnd}")
        return True

    def press_key(self, key):
        """按下並釋放按鍵"""
        try:
            if not self.hwnd or not win32gui.IsWindow(self.hwnd):
                if not self.find_window():
                    return
            
            # 嘗試激活窗口
            foreground_success = False
            try:
                # 檢查窗口是否最小化
                if win32gui.IsIconic(self.hwnd):
                    # 如果最小化，先恢復窗口
                    win32gui.ShowWindow(self.hwnd, win32con.SW_RESTORE)
                    time.sleep(0.5)
                
                # 嘗試切換到前台
                win32gui.SetForegroundWindow(self.hwnd)
                foreground_success = True
            except Exception as e:
                self.logger.warning(f"無法將窗口設為前台，嘗試使用備用方法: {str(e)}")
            
            time.sleep(0.1)
            
            # 如果窗口激活成功，使用win32api發送按鍵
            if foreground_success:
                scan_code = win32api.MapVirtualKey(key, 0)
                win32api.keybd_event(key, scan_code, 0, 0)
                time.sleep(0.1)
                win32api.keybd_event(key, scan_code, win32con.KEYEVENTF_KEYUP, 0)
            else:
                # 備用方法：使用pyautogui發送按鍵
                self._send_key_with_pyautogui(key)
            
        except Exception as e:
            self.logger.error(f"按鍵操作出錯: {str(e)}")
            import traceback
            self.logger.debug(f"詳細錯誤: {traceback.format_exc()}")
            
    def _send_key_with_pyautogui(self, key):
        """使用pyautogui發送按鍵（備用方法）"""
        try:
            # 將win32 VK碼轉換為pyautogui可以理解的鍵名
            key_map = {
                self.KEYS["RIGHT"]: 'd',
                self.KEYS["LEFT"]: 'a',
                self.KEYS["SPACE"]: 'space',
                self.KEYS["E"]: 'e',
                self.KEYS["S"]: 's',
                self.KEYS["W"]: 'w',
                self.KEYS["ESC"]: 'esc'
            }
            
            if key in key_map:
                self.logger.info(f"使用pyautogui發送按鍵: {key_map[key]}")
                pyautogui.press(key_map[key])
            else:
                self.logger.warning(f"未知的按鍵代碼: {key}")
        except Exception as e:
            self.logger.error(f"pyautogui發送按鍵失敗: {str(e)}")

    def get_screenshot(self):
        """獲取遊戲窗口的截圖"""
        try:
            if not self.hwnd:
                self.logger.error("未找到遊戲視窗句柄")
                return None
                
            # 檢查窗口是否還存在
            if not win32gui.IsWindow(self.hwnd):
                self.logger.error("遊戲視窗已關閉或不可用")
                self.hwnd = None
                if not self.find_window():
                    return None
            
            # 檢查窗口是否最小化
            if win32gui.IsIconic(self.hwnd):
                self.logger.error("遊戲視窗已最小化")
                return None
                
            rect = win32gui.GetWindowRect(self.hwnd)
            left, top, right, bottom = rect
            width = right - left
            height = bottom - top
            
            self.logger.debug(f"遊戲視窗位置: 左={left}, 上={top}, 右={right}, 下={bottom}")
            self.logger.debug(f"遊戲視窗大小: 寬={width}, 高={height}")
            
            if width <= 0 or height <= 0:
                self.logger.error(f"視窗大小異常: {width}x{height}")
                return None
            
            try:
                screenshot = pyautogui.screenshot(region=(left, top, width, height))
                screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
                self.logger.debug(f"截圖大小: {screenshot.shape}")
                return screenshot
            except Exception as e:
                self.logger.error(f"pyautogui截圖失敗: {str(e)}")
                return None
            
        except Exception as e:
            import traceback
            self.logger.error(f"獲取截圖失敗: {str(e)}")
            self.logger.error(f"錯誤詳情: {traceback.format_exc()}")
            return None

    def detect_image(self, image_key, threshold=None):
        """檢測圖片"""
        try:
            # 從配置文件獲取圖片路徑
            image_path = self.config["image_paths"].get(image_key)
            if not image_path:
                self.logger.error(f"找不到圖片配置: {image_key}")
                return False, None
                
            # 從配置文件獲取閾值，如果沒有設定則使用默認值
            if threshold is None:
                threshold = self.config["thresholds"].get(image_key, 0.8)
            
            # 顯示當前查找的圖片和閾值
            self.logger.info(f"查找: {image_key} - 閾值: {threshold:.2f}")
                
            screenshot = self.get_screenshot()
            if screenshot is None:
                return False, None
                
            template = cv2.imread(image_path)
            if template is None:
                self.logger.error(f"無法讀取圖片: {image_path}")
                return False, None
                
            result = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
            
            # 顯示匹配結果
            self.logger.info(f"結果: {image_key} - 匹配值: {max_val:.2f}")
            
            if max_val >= threshold:  # 使用閾值而不是固定的0.99
                self.logger.info(f"匹配成功: {image_key} - 執行操作")
                return True, max_loc
            return False, None
            
        except Exception as e:
            self.logger.error(f"圖片檢測失敗: {str(e)}")
            return False, None

    def handle_domination_flow(self):
        """處理稱霸賽流程"""
        self.logger.info("開始稱霸賽流程")
        
        # 需要檢測的圖片順序
        flow_images = [
            "mycareer",
            "myteam",
            "daily_reward",
            "new_content",
            "domination_home",
            "domination_btn",
            "select"
        ]
        
        while self.is_running:
            # 按照流程順序檢測圖片
            for image_key in flow_images:
                if self.detect_image(image_key)[0]:
                    if image_key == "new_content":
                        self.logger.info("找到全新內容 → 按E鍵")
                        self.press_key(self.KEYS["E"])
                        time.sleep(0.5)
                        self.press_key(self.KEYS["E"])
                        time.sleep(0.5)
                        break
                        
                    elif image_key == "domination_btn":
                        self.logger.info("找到稱霸賽按鈕 → 按空格+D鍵")
                        self.press_key(self.KEYS["SPACE"])
                        time.sleep(1)
                        # 按D鍵8次
                        for _ in range(8):
                            self.press_key(self.KEYS["RIGHT"])
                            time.sleep(0.2)
                        break
                        
                    elif image_key == "domination_home":
                        self.logger.info("找到稱霸賽主頁 → 按S鍵")
                        for i in range(5):
                            self.press_key(self.KEYS["S"])
                            time.sleep(0.5)
                        break
                        
                    elif image_key == "mycareer":
                        self.logger.info("找到MyCAREER → 按D鍵")
                        self.press_key(self.KEYS["RIGHT"])
                        time.sleep(0.5)
                        break
                        
                    elif image_key == "myteam":
                        self.logger.info("找到MyTEAM → 按空格")
                        self.press_key(self.KEYS["SPACE"])
                        time.sleep(0.5)
                        break
                        
                    elif image_key == "daily_reward":
                        self.logger.info("找到每日獎勵 → 按空格")
                        self.press_key(self.KEYS["SPACE"])
                        time.sleep(3)
                        self.press_key(self.KEYS["SPACE"])
                        break
                        
                    elif image_key == "select":
                        self.logger.info("找到選擇按鈕 → 按空格")
                        self.press_key(self.KEYS["SPACE"])
                        time.sleep(0.5)
                        # 檢測三星並開始遊戲
                        self.handle_three_stars_game()
                        break
                        
            time.sleep(0.5)

    def handle_three_stars_game(self):
        """處理三星遊戲流程"""
        self.logger.info("檢測三星")
        
        while self.is_running:
            if self.detect_image("three_stars")[0]:
                self.logger.info("找到三星 → 開始遊戲")
                self.trigger_game_start()
                break
            time.sleep(0.5)

    def trigger_game_start(self):
        """開始遊戲流程"""
        self.logger.info("進入遊戲流程")
        
        # 按空格確認
        self.press_key(self.KEYS["SPACE"])
        time.sleep(1)
        
        # 選擇難度
        self.logger.info("選擇難度 → 按S鍵")
        for _ in range(2):
            self.press_key(self.KEYS["S"])
            time.sleep(0.5)
        
        # 確認難度
        self.press_key(self.KEYS["SPACE"])
        time.sleep(0.5)
        
        # 開始遊戲
        self.press_key(self.KEYS["SPACE"])
        time.sleep(0.5)
        
        # 進入遊戲循環
        self.logger.info("進入遊戲循環")
        while self.is_running:
            # 循環檢查遊戲中的按鈕
            if self.handle_game_buttons():
                # 找到並處理了一個按鈕，等待後繼續循環
                time.sleep(0.5)
            else:
                # 未找到任何按鈕，短暫等待
                time.sleep(0.2)

    def handle_game_buttons(self):
        """處理遊戲中的按鈕"""
        if self.detect_image("forward")[0]:
            self.logger.info("找到前進按鈕 → 按空格")
            self.press_key(self.KEYS["SPACE"])
            time.sleep(0.5)
            return True
            
        elif self.detect_image("pause")[0]:
            self.logger.info("找到暫停按鈕 → 按空格")
            self.press_key(self.KEYS["SPACE"])
            time.sleep(0.5)
            return True
            
        elif self.detect_image("continue")[0]:
            self.logger.info("找到繼續按鈕 → 按空格")
            self.press_key(self.KEYS["SPACE"])
            self.state['in_domination'] = False
            time.sleep(0.5)
            return True
            
        elif self.detect_image("three_stars")[0]:
            self.logger.info("找到三星圖片 → 開始遊戲")
            self.trigger_game_start()
            return True
            
        return False

def main():
    """主函數"""
    handler = ImageHandler()
    
    try:
        handler.handle_domination_flow()
    except KeyboardInterrupt:
        handler.logger.info("使用者中斷程式")
        handler.is_running = False
    except Exception as e:
        handler.logger.error(f"程式執行時發生錯誤: {str(e)}")
        handler.is_running = False

if __name__ == "__main__":
    main() 