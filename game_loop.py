import os
import time
import win32api
import win32con
import win32gui
import cv2
import numpy as np
import pyautogui
import logging
from datetime import datetime
import traceback
import json
from pathlib import Path

def setup_logging():   #日志設置
    log_dir = Path("logs")
    log_dir.mkdir(exist_ok=True)
    
    # 限制日志文件数量
    MAX_LOG_FILES = 5
    log_files = sorted(log_dir.glob("game_log_*.log"), key=lambda x: x.stat().st_mtime, reverse=True)
    
    # 删除旧的日志文件
    for old_log in log_files[MAX_LOG_FILES-1:]:
        try:
            old_log.unlink()
        except Exception as e:
            print(f"无法删除旧日志文件 {old_log}: {e}")
    
    current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = log_dir / f"game_log_{current_time}.log"
    
    log_format = '%(asctime)s [%(levelname)s] %(message)s'
    date_format = '%Y-%m-%d %H:%M:%S'
    
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)  # 将日志级别改为INFO
    
    for handler in root_logger.handlers[:]:
        root_logger.removeHandler(handler)
    
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(logging.INFO)  # 将文件处理器的级别改为INFO
    file_handler.setFormatter(logging.Formatter(log_format, date_format))
    root_logger.addHandler(file_handler)
    
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)  # 将控制台处理器的级别改为INFO
    console_handler.setFormatter(logging.Formatter('%(message)s'))
    root_logger.addHandler(console_handler)
    
    logging.info("遊戲啟動")
    return log_file

class GameWindow:   #遊戲視窗
    def __init__(self, window_name="NBA 2K25"):
        self.window_name = window_name
        self.hwnd = None
        self.find_window()

    def find_window(self):
        """查找遊戲窗口"""
        self.hwnd = win32gui.FindWindow(None, self.window_name)
        if not self.hwnd:
            logging.error("未找到遊戲視窗")
            return False
        return True

    def get_window_rect(self):
        """獲取遊戲窗口的位置和大小"""
        try:
            if not self.hwnd:
                return None
            return win32gui.GetWindowRect(self.hwnd)
        except Exception as e:
            return None

    def get_screenshot(self):
        """獲取遊戲窗口的截圖"""
        try:
            rect = win32gui.GetWindowRect(self.hwnd)
            left, top, right, bottom = rect
            width = right - left
            height = bottom - top
            
            screenshot = pyautogui.screenshot(region=(left, top, width, height))
            screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
            return screenshot
            
        except Exception as e:
            return None

class GameState:
    def __init__(self): #遊戲狀態
        self.states = {
            'in_myteam': False,
            'in_domination': False,
            'search_count': 0
        }
    
    def set(self, **kwargs):
        """設置一個或多個狀態"""
        for key, value in kwargs.items():
            if key in self.states:
                self.states[key] = value
            else:
                logging.warning(f"未知的狀態: {key}")
    
    def get(self, key):
        """獲取狀態值"""
        return self.states.get(key)
    
    def reset(self):
        """重置所有狀態為默認值"""
        self.states = {
            'in_myteam': False,
            'in_domination': False,
            'search_count': 0
        }

class GameLoop:
    def __init__(self): #遊戲循環
        self.is_running = False
        self.window_name = "NBA 2K25"
        self.logger = logging.getLogger(__name__)
        self.game_window = GameWindow(self.window_name)
        self.state = GameState()
        
        # 載入配置
        self.load_config()
        
        self.KEYS = {
            "RIGHT": ord('D'), "LEFT": ord('A'), "SPACE": win32con.VK_SPACE,
            "E": ord('E'), "S": ord('S'), "W": ord('W'), "ESC": win32con.VK_ESCAPE
        }
        
        # 定義模板匹配方法
        self.MATCH_METHODS = [
            cv2.TM_CCOEFF_NORMED,
            cv2.TM_CCORR_NORMED,
            cv2.TM_SQDIFF_NORMED
        ]

    def load_config(self):
        """載入配置文件"""
        try:
            with open('config.json', 'r', encoding='utf-8') as f:
                config = json.load(f)
            self.paths = config['image_paths']
            self.thresholds = config['thresholds']
            self.priority_order = config['priority']
            self.logger.info("成功載入配置文件")
        except Exception as e:
            self.logger.error(f"載入配置文件失敗: {str(e)}")
            raise

    def find_game_window(self):
        """查找遊戲窗口"""
        self.game_window.find_window()
        if not self.game_window.hwnd:
            self.logger.error("未找到遊戲視窗")
            return False
        return True

    def press_and_release(self, key):
        if not self.game_window.hwnd or not win32gui.IsWindow(self.game_window.hwnd):
            if not self.find_game_window(): 
                self.logger.error("找不到遊戲視窗")
                return
                
        if win32gui.GetWindowLong(self.game_window.hwnd, win32con.GWL_STYLE) & win32con.WS_MINIMIZE:
            self.logger.warning("視窗最小化")
            return
            
        try:
            win32gui.SetForegroundWindow(self.game_window.hwnd)
            time.sleep(0.1)  # 縮短設置前景窗口等待時間
            scan_code = win32api.MapVirtualKey(key, 0)
            win32api.keybd_event(key, scan_code, 0, 0)  # 按下鍵
            time.sleep(0.1)  # 縮短按住時間為0.1秒
            win32api.keybd_event(key, scan_code, win32con.KEYEVENTF_KEYUP, 0)  # 釋放鍵
        except Exception as e:
            self.logger.error(f"按鍵操作出錯: {str(e)}")

    def detect_image(self, image_name, threshold=None):
        """檢測圖片"""
        if not os.path.exists(image_name):
            self.logger.error(f"圖片不存在: {image_name}")
            return False, None
            
        try:
            rect = self.game_window.get_window_rect()
            if not rect:
                self.logger.error("無法獲取窗口區域")
                return False, None
                
            screenshot = self.game_window.get_screenshot()
            if screenshot is None:
                self.logger.error("無法獲取截圖")
                return False, None
                
            template = cv2.imread(image_name)
            if template is None:
                self.logger.error(f"無法讀取模板圖片: {image_name}")
                return False, None
            
            debug_dir = Path("debug")
            debug_dir.mkdir(exist_ok=True)
            
            # 限制调试图片数量
            MAX_DEBUG_FILES = 50
            debug_files = sorted(debug_dir.glob("*.png"), key=lambda x: x.stat().st_mtime, reverse=True)
            
            # 删除旧的调试图片
            for old_file in debug_files[MAX_DEBUG_FILES-1:]:
                try:
                    old_file.unlink()
                except Exception as e:
                    self.logger.warning(f"無法刪除舊調試圖片 {old_file}: {e}")
            
            # 只在匹配分数较高时保存调试图片
            best_val = -1
            max_loc = None
            
            for method in self.MATCH_METHODS:
                result = cv2.matchTemplate(screenshot, template, method)
                min_val, max_val, min_loc, curr_max_loc = cv2.minMaxLoc(result)
                curr_val = 1 - min_val if method == cv2.TM_SQDIFF_NORMED else max_val
                
                if curr_val > best_val:
                    best_val = curr_val
                    max_loc = curr_max_loc if method != cv2.TM_SQDIFF_NORMED else min_loc
            
            threshold = threshold if threshold is not None else self.thresholds.get(Path(image_name).stem, 0.8)
            
            # 只在匹配分数接近阈值时记录日志
            if abs(best_val - threshold) < 0.1:
                self.logger.info(f"檢測圖片 {Path(image_name).stem} - 分數: {best_val:.3f} - 閾值: {threshold:.3f}")
            
            if best_val >= (threshold - 0.001):
                # 只在成功匹配时保存调试图片
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                h, w = template.shape[:2]
                top_left = max_loc
                bottom_right = (top_left[0] + w, top_left[1] + h)
                match_img = screenshot.copy()
                cv2.rectangle(match_img, top_left, bottom_right, (0, 255, 0), 2)
                cv2.imwrite(str(debug_dir / f"{timestamp}_{Path(image_name).stem}_match.png"), match_img)
                return True, max_loc
            
            return False, None
                
        except Exception as e:
            self.logger.error(f"圖片匹配出錯: {str(e)}")
            return False, None

    def handle_matched_image(self, image_name):#遊戲處理
        image_name = Path(image_name).stem.lower()
        
        if image_name == "new_content":
            self.logger.info("檢測到全新內容")
            self.press_and_release(self.KEYS["E"])
            time.sleep(0.5)
            self.press_and_release(self.KEYS["E"])
            time.sleep(0.5)
            return True
            
        elif image_name == "domination_btn":
            self.logger.info("檢測到稱霸賽按鈕")
            self.press_and_release(self.KEYS["SPACE"])
            self.state.set(in_domination=True)
            time.sleep(0.5)
            
            # 先檢查是否有滿星
            self.check_full_stars()
            
            # 檢查選擇按鈕
            if self.detect_image(self.paths["select"], threshold=self.thresholds["select"])[0]:
                self.logger.info("檢測到選擇按鈕")
                self.press_and_release(self.KEYS["SPACE"])
                time.sleep(0.5)
                self.handle_three_stars_search()
            return True
                
        elif image_name == "domination_home":
            self.logger.info("檢測到稱霸賽主頁")
            for i in range(5):
                self.press_and_release(self.KEYS["S"])
                time.sleep(0.5)
            return True
                
        elif image_name == "mycareer":
            self.logger.info("檢測到MyCAREER")
            self.press_and_release(self.KEYS["RIGHT"])
            time.sleep(0.5)
            return True
            
        elif image_name == "myteam":
            self.logger.info("檢測到MyTEAM")
            self.press_and_release(self.KEYS["SPACE"])
            self.state.set(in_myteam=True)
            time.sleep(0.5)
            return True
            
        elif image_name == "daily_reward":
            self.logger.info("檢測到每日獎勵")
            self.press_and_release(self.KEYS["SPACE"])
            time.sleep(3)
            self.press_and_release(self.KEYS["SPACE"])
            return True
            
        return False

    def handle_main_images(self):#主圖片處理
        any_image_handled = False
        
        for image_name in self.priority_order:
            if not image_name in self.paths:
                continue
                
            detected, loc = self.detect_image(self.paths[image_name], threshold=self.thresholds[image_name])
            
            if detected:
                result = self.handle_matched_image(image_name)
                if result:
                    any_image_handled = True
                    if image_name in ["new_content", "daily_reward"]:
                        return True
                        
        return any_image_handled

    def check_three_stars(self):#三星檢查
        self.logger.info("檢查三星...")
        
        # 檢查所有三星圖片
        star_images = ["stars", "stars2", "stars3", "stars4"]
        for img_name in star_images:
            if img_name in self.paths and img_name in self.thresholds:
                result = self.detect_image(self.paths[img_name], threshold=self.thresholds[img_name])
                if result[0]:
                    self.logger.info(f"找到三星！({img_name})")
                    return True
            
        return False

    def press_key_and_check_stars(self, key, times):#三星按鍵檢查
        self.logger.info(f"按{times}次按鍵並檢查三星")
        for i in range(times):
            self.press_and_release(key)
            time.sleep(0.5)
            if self.check_three_stars():
                self.logger.info("按鍵檢查時找到三星")
                return True
        return False

    def check_and_trigger_game(self, message="找到三星！"):#檢查三星並開始遊戲
        """檢查是否有三星並觸發遊戲開始"""
        if self.check_three_stars():
            self.logger.info(f"{message}準備開始遊戲")
            self.state.set(search_count=0)
            self.trigger_game_start()
            return True
        return False
        
    def press_key_sequence(self, key, times, pause=0.3):#多次按鍵
        """按指定鍵多次"""
        for _ in range(times):
            self.press_and_release(key)
            time.sleep(pause)
    
    def navigate_and_check(self, direction_key, description):#導航檢查
        """導航並檢查三星"""
        self.logger.info(f"按{description}鍵")
        self.press_key_sequence(direction_key, 1)
        
        # 檢查三星
        if self.check_and_trigger_game():
            return True
            
        # 向上導航
        self.logger.info("按W鍵")
        self.press_key_sequence(self.KEYS["W"], 5)
        
        # 再次檢查
        if self.check_and_trigger_game():
            return True
            
        # 向下導航
        self.logger.info("按S鍵")
        if self.press_key_and_check_stars(self.KEYS["S"], 5):
            self.logger.info("找到三星！準備開始遊戲")
            self.state.set(search_count=0)                
            time.sleep(0.5)
            self.trigger_game_start()
            return True
            
        return False

    def handle_three_stars_search(self):#三星搜尋
        # 重置搜尋次數
        self.state.set(search_count=0)
        self.logger.info("=== 開始三星搜尋 ===")
        
        while self.is_running:
            search_count = self.state.get('search_count') + 1
            self.state.set(search_count=search_count)
            self.logger.info(f"\n第 {search_count} 次搜尋開始...")
            
            # 先向左導航檢查
            if self.navigate_and_check(self.KEYS["LEFT"], "左"):
                continue
                
            # 再向右導航檢查
            if self.navigate_and_check(self.KEYS["RIGHT"], "右"):
                continue
            
            # 最後檢查一次
            if self.check_and_trigger_game("最後檢查，"):
                continue
            
            # 本次搜索失敗，判斷是否需要換下一個
            self.logger.info("本次搜尋失敗")
            
            if search_count >= 2:
                self.logger.info("已搜尋2次，準備切換下一個")
                self.press_and_release(self.KEYS["ESC"])
                time.sleep(0.5)
                self.press_and_release(self.KEYS["RIGHT"])
                time.sleep(0.5)
                self.press_and_release(self.KEYS["SPACE"])
                self.state.set(search_count=0)
                self.logger.info("重置搜尋次數為0")
            
            time.sleep(0.5)

    def trigger_game_start(self):  # 進入遊戲流程
        self.logger.info("=== 開始進入遊戲流程 ===")
        self.logger.info("按空格確認")
        self.press_and_release(self.KEYS["SPACE"])
        time.sleep(1)
        
        self.logger.info("按兩次S鍵選擇難度")
        for i in range(2):
            self.press_and_release(self.KEYS["S"])
            time.sleep(0.5)        
        
        self.logger.info("按空格確認難度")
        self.press_and_release(self.KEYS["SPACE"])
        time.sleep(0.5)
        
        self.logger.info("按空格開始遊戲")
        self.press_and_release(self.KEYS["SPACE"])
        time.sleep(0.5)

        # 進入遊戲循環
        self.logger.info("=== 進入遊戲循環 ===")
        while self.is_running:
            if self.handle_game_buttons():
                continue
            time.sleep(0.5)

    def check_full_stars(self):#滿星檢查
        self.logger.info("檢查滿星...")
        
        # 初始化變數
        not_found_count = 0  # 連續未找到的次數
        found_any = False    # 是否找到過任何滿星
        max_not_found = 3    # 連續幾次未找到就退出
        
        # 只要連續未找到的次數不超過閾值，就繼續搜索
        while not_found_count < max_not_found:
            result = self.detect_image(self.paths["full_of_stars"], threshold=self.thresholds["full_of_stars"])
            
            if result[0]:
                self.logger.info(f"找到滿星！按D鍵切換")
                self.press_and_release(self.KEYS["RIGHT"])
                time.sleep(0.2)  # 給畫面更新時間
                found_any = True
                not_found_count = 0  # 重置未找到計數
            else:
                not_found_count += 1
                if not_found_count < max_not_found:
                    self.logger.info(f"第{not_found_count}次未找到滿星，繼續搜索...")
                    time.sleep(0.2)  # 短暫等待後再次檢測
        
        if found_any:
            self.logger.info("完成所有滿星處理")
            return True
        else:
            self.logger.info(f"連續{max_not_found}次未找到滿星，退出搜索")
            return False

    def handle_game_buttons(self): # 進入遊戲循環
        if self.detect_image(self.paths["forward"])[0]:
            self.logger.info("找到前進按鈕，按下空白鍵")
            self.press_and_release(self.KEYS["SPACE"])
            time.sleep(0.5)
            return True
        elif self.detect_image(self.paths["pause"])[0]:
            self.logger.info("找到暫停按鈕，按下空白鍵")
            self.press_and_release(self.KEYS["SPACE"])
            time.sleep(0.5)
            return True
        elif self.detect_image(self.paths["continue"])[0]:
            self.logger.info("找到繼續按鈕，按下空白鍵")
            self.press_and_release(self.KEYS["SPACE"])
            self.state.set(in_domination=False)
            time.sleep(0.5)
            return True
        elif self.check_three_stars():
            self.logger.info("找到三星按鈕，進入三星搜尋")
            self.handle_three_stars_search()
            return True
            
        self.logger.info("尋找圖片...")
        return False

    def start(self):#開始
        if not self.find_game_window():
            self.logger.error("找不到遊戲視窗，程式退出")
            return
        self.logger.info("開始執行自動化程序...")
        self.is_running = True
        try:
            self.main_loop()
        except KeyboardInterrupt:
            self.logger.info("\n使用者中止程式")
            self.is_running = False
        except Exception as e:
            self.logger.error(f"發生錯誤: {str(e)}")
            self.is_running = False

    def stop(self):
        self.is_running = False

    def main_loop(self):
        self.state.reset()  # 重置所有狀態

        while self.is_running:
            if not win32gui.IsWindow(self.game_window.hwnd):
                if not self.find_game_window(): break

            # 檢查主要圖片
            self.handle_main_images()
            time.sleep(0.5)

def main(): #主函數
    log_file = setup_logging()
    
    try:
        game = GameLoop()
        game.start()
    except KeyboardInterrupt:
        logging.info("使用者中斷程式")
    except Exception as e:
        logging.error(f"程式執行時發生錯誤: {str(e)}")
        logging.error(f"錯誤詳情:\n{traceback.format_exc()}")
    finally:
        logging.info("程式結束")

if __name__ == "__main__":
    main()
