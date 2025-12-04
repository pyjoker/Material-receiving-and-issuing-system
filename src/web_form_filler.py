"""網頁表單自動填寫模組"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import time
from typing import Optional


class WebFormFiller:
    """自動填寫網頁表單的類別"""
    
    def __init__(self, headless: bool = False):
        """
        初始化 WebFormFiller
        
        Args:
            headless: 是否使用無頭模式（不顯示瀏覽器視窗）
        """
        self.driver = None
        self.headless = headless
    
    def start_browser(self):
        """啟動瀏覽器"""
        options = webdriver.ChromeOptions()
        if self.headless:
            options.add_argument('--headless')
        options.add_argument('--disable-gpu')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        
        # 自動下載並使用最新的 ChromeDriver
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=options)
        print("✓ 瀏覽器已啟動")
    
    def open_url(self, url: str, wait_time: int = 10):
        """
        開啟指定網址
        
        Args:
            url: 網頁 URL
            wait_time: 等待頁面載入的時間（秒）
        """
        if not self.driver:
            self.start_browser()
        
        self.driver.get(url)
        print(f"✓ 已開啟網址: {url}")
        
        # 等待頁面載入
        time.sleep(wait_time)
    
    def find_item_index(self, item_value: str) -> Optional[int]:
        """
        查找項次對應的索引
        
        Args:
            item_value: 項次值（例如 "2-1"）
        
        Returns:
            找到的索引，若未找到則返回 None
        """
        try:
            # 移除項次值的空白
            item_value = str(item_value).strip()
            
            # 嘗試查找所有可能的 gvReceive_lblItem 元素
            index = 0
            while True:
                try:
                    element_id = f"gvReceive_lblItem_{index}"
                    element = self.driver.find_element(By.ID, element_id)
                    element_text = element.text.strip()
                    
                    if element_text == item_value:
                        print(f"  ✓ 找到項次 '{item_value}' 在索引 {index}")
                        return index
                    
                    index += 1
                    
                    # 設定最大搜尋範圍避免無限循環
                    if index > 1000:
                        break
                        
                except NoSuchElementException:
                    # 沒有更多元素了
                    break
            
            print(f"  ✗ 未找到項次 '{item_value}'")
            return None
            
        except Exception as e:
            print(f"  ✗ 查找項次時發生錯誤: {e}")
            return None
    
    def fill_quantity_and_amount(self, index: int, quantity: float, amount: float) -> bool:
        """
        填入數量和複價
        
        Args:
            index: 項次的索引
            quantity: 數量
            amount: 複價
        
        Returns:
            填寫成功返回 True，失敗返回 False
        """
        try:
            # 填入數量
            qty_element_id = f"gvReceive_txtRecvQty_{index}"
            qty_element = self.driver.find_element(By.ID, qty_element_id)
            qty_element.clear()
            qty_element.send_keys(str(quantity))
            print(f"    ✓ 已填入數量: {quantity}")
            
            # 填入複價
            amt_element_id = f"gvReceive_txtRecvAmt_{index}"
            amt_element = self.driver.find_element(By.ID, amt_element_id)
            amt_element.clear()
            amt_element.send_keys(str(amount))
            print(f"    ✓ 已填入複價: {amount}")
            
            # 短暫延遲，確保數據已輸入
            time.sleep(0.3)
            
            return True
            
        except Exception as e:
            print(f"    ✗ 填入數據時發生錯誤: {e}")
            return False
    
    def process_dataframe(self, df: pd.DataFrame, delay: float = 0.5) -> dict:
        """
        處理整個 DataFrame，自動填寫表單
        
        Args:
            df: 包含「項次」、「數量」、「複價」欄位的 DataFrame
            delay: 每筆資料之間的延遲時間（秒）
        
        Returns:
            包含處理結果的字典
        """
        if not self.driver:
            raise RuntimeError("瀏覽器尚未啟動，請先呼叫 start_browser() 或 open_url()")
        
        results = {
            'total': len(df),
            'success': 0,
            'failed': 0,
            'not_found': 0,
            'failed_items': []
        }
        
        print("\n" + "=" * 50)
        print("開始處理 DataFrame 資料...")
        print("=" * 50)
        
        for idx, row in df.iterrows():
            item = str(row['項次']).strip()
            quantity = row['數量']
            amount = row['複價']
            
            print(f"\n[{idx + 1}/{len(df)}] 處理項次: {item}")
            
            # 查找項次對應的索引
            web_index = self.find_item_index(item)
            
            if web_index is None:
                results['not_found'] += 1
                results['failed_items'].append({
                    'item': item,
                    'reason': '網頁中未找到此項次'
                })
                continue
            
            # 填入數量和複價
            success = self.fill_quantity_and_amount(web_index, quantity, amount)
            
            if success:
                results['success'] += 1
            else:
                results['failed'] += 1
                results['failed_items'].append({
                    'item': item,
                    'reason': '填入數據時發生錯誤'
                })
            
            # 延遲，避免操作過快
            time.sleep(delay)
        
        print("\n" + "=" * 50)
        print("處理完成！")
        print("=" * 50)
        print(f"總計: {results['total']} 筆")
        print(f"成功: {results['success']} 筆")
        print(f"失敗: {results['failed']} 筆")
        print(f"未找到: {results['not_found']} 筆")
        
        if results['failed_items']:
            print("\n失敗的項次:")
            for item in results['failed_items']:
                print(f"  - {item['item']}: {item['reason']}")
        
        return results
    
    def close_browser(self):
        """關閉瀏覽器"""
        if self.driver:
            self.driver.quit()
            print("\n✓ 瀏覽器已關閉")
    
    def __enter__(self):
        """支援 with 語句"""
        self.start_browser()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """支援 with 語句"""
        self.close_browser()


def fill_web_form_from_dataframe(df: pd.DataFrame, url: str, headless: bool = False, 
                                  wait_time: int = 10, delay: float = 0.5) -> dict:
    """
    便捷函式：從 DataFrame 自動填寫網頁表單
    
    Args:
        df: 包含「項次」、「數量」、「複價」欄位的 DataFrame
        url: 網頁 URL
        headless: 是否使用無頭模式
        wait_time: 等待頁面載入的時間（秒）
        delay: 每筆資料之間的延遲時間（秒）
    
    Returns:
        包含處理結果的字典
    """
    with WebFormFiller(headless=headless) as filler:
        filler.open_url(url, wait_time=wait_time)
        results = filler.process_dataframe(df, delay=delay)
    
    return results
