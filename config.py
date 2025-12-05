"""配置檔案 - 集中管理所有路徑設定"""

from pathlib import Path

# ============ 輸入檔案路徑 ============
# Excel 來源檔案路徑
INPUT_FILE_PATH = Path("C:\\Users\\03010430\\Documents\\16P2759A-M0008-003-伸泰-電線電纜(第10期次計價)-114.11.29 - 複製.xls")

# ============ 輸出目錄路徑 ============
# 處理後的檔案輸出目錄
OUTPUT_DIR = Path("C:\\Users\\03010430\\Documents")

# 已處理的資料檔案路徑（如果有現成的處理後資料）
PROCESSED_FILE_PATH = Path("C:\\Users\\03010430\\Documents\\processed_data_20251206_032835.xlsx")

# ============ 其他設定 ============
# 預設工作表索引或名稱
DEFAULT_SHEET_INDEX = 2  # 第三個工作表

# 要讀取的欄位（Excel 欄位字母）
COLUMNS_TO_READ = "C,T,U"

# 欄位重新命名對應
COLUMN_RENAME_MAP = {
    0: '項次',  # 第一欄
    1: '數量',  # 第二欄
    2: '複價'   # 第三欄
}
