#shopee-oms 5.5 完整版

import json
import sys
import tkinter as tk
from tkinter import ttk, messagebox, font
import pandas as pd
from datetime import datetime, timedelta  # 引入 timedelta 來處理時區加減
import os
import re
import pickle
import threading 
import hashlib
from ImportWizard import ImportWizard
from ShippingWizard import show_shipping_dialog
from decimal import Decimal, ROUND_HALF_UP
import platform
import uuid




# 1. 匯入敏感資料
try:
    from secrets_config import SECRET_SALT, AUTH_FILE, RESCUE_SALT, RESCUE_ACCOUNT
except ImportError:
    print("ℹ️ 系統提示：未偵測到自定義設定檔，將以『預覽模式』啟動 (使用預設安全參數)。")
    SECRET_SALT = "PUBLIC_DEMO_SALT_2026"
    AUTH_FILE = "sys_config.bin"
    RESCUE_SALT = "RESCUE_DEMO_SALT"
    RESCUE_ACCOUNT = "RESCUE_ADMIN"


def get_machine_id():
    """ 獲取電腦唯一識別碼 """
    node = uuid.getnode()
    return hashlib.sha256(str(node).encode()).hexdigest()[:12].upper()


# 2. 加入這段函式：用來處理打包後的資源路徑
def resource_path(relative_path):
    """ 獲取資源的絕對路徑，兼容 Dev 和 PyInstaller """
    try:
        # PyInstaller 創建臨時文件夾，路徑存儲在 _MEIPASS 中
        base_path = sys._MEIPASS
    except Exception:
        # 獲取「目前這個檔案」所在的資料夾絕對路徑
        base_path = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(base_path, relative_path)
    

def hash_password(password):
    """ 將密碼加上 Salt 後進行 SHA256 加密 """
    # 延用你之前的 SECRET_SALT，增加破解難度
    salt = SECRET_SALT 
    return hashlib.sha256((password + salt).encode()).hexdigest()


def secure_hash(text):
    """ 使用 Salt 進行 SHA256 雜湊，確保不可逆 """
    # 這裡的 Salt 建議與 License 使用不同的字串
    internal_salt = "ERP_INTERNAL_SECURITY_2026" 
    return hashlib.sha256((text + internal_salt).encode()).hexdigest()


def get_rescue_password():
    """ 
    動態救援密鑰：結合年月，密鑰每個月會自動改變
    """
    # 取得當前年月 (例如 "202602")
    dynamic_factor = datetime.now().strftime("%Y%m") 
    
    # 組合：Salt + 暗號 + 年月
    raw_string = SECRET_SALT + RESCUE_SALT + dynamic_factor
    
    return hashlib.sha256(raw_string.encode()).hexdigest()[:10].upper()

# --- Google Drive 相關套件 ---
try:
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    GOOGLE_LIB_INSTALLED = True
except ImportError:
    GOOGLE_LIB_INSTALLED = False

# 設定 Excel 檔案名稱
FILE_NAME = resource_path('sales_data.xlsx')
CREDENTIALS_FILE = resource_path('credentials.json')  
TOKEN_FILE =  resource_path('token.json')             
SCOPES = ['https://www.googleapis.com/auth/drive.file'] 

SHEET_PURCHASES = '進貨紀錄'
SHEET_PUR_TRACKING = '進貨追蹤'
SHEET_VENDORS = '進貨廠商管理'
SHEET_SALES = '銷售紀錄'      # 歷史已完成訂單
SHEET_TRACKING = '訂單追蹤'   # 未完成/出貨中 (緩衝區)
SHEET_RETURNS = '退貨紀錄'    # 退貨區
SHEET_PRODUCTS = '商品資料'
SHEET_FEES = '手續費設定'       # 原本的 '系統設定' 內容搬到這
SHEET_SYS_SETTINGS = '系統設定'  # 專門存店名、版本、權限等





# 設定雲端硬碟上的備份資料夾名稱
BACKUP_FOLDER_NAME = "蝦皮進銷存系統_備份"

TAIWAN_CITIES = [
    "基隆市", "臺北市", "新北市", "桃園市", "新竹市", "新竹縣", "苗栗縣",
    "臺中市", "彰化縣", "南投縣", "雲林縣", "嘉義市", "嘉義縣", "臺南市",
    "高雄市", "屏東縣", "宜蘭縣", "花蓮縣", "臺東縣", "澎湖縣", "金門縣", "連江縣",
    "海外", "面交", "未提供"
]


PLATFORM_OPTIONS = [
    "蝦皮購物", "賣貨便(7-11)", "好賣家(全家)", "旋轉拍賣", 
    "官方網站", "Facebook社團", "IG", "PChome", "Momo", "實體店面/面交"
]

SHIPPING_METHODS = [
    "7-11", "全家", "萊爾富", "OK超商", "蝦皮店到店", 
    "蝦皮店到店-隔日到貨", "蝦皮店到宅",
    "黑貓宅急便", "新竹物流", "郵局掛號", "賣家宅配", "面交/自取"
]

def thread_safe_file(func):
    """ 裝飾器：自動為檔案操作加上互斥鎖 (支援跨類別) """
    def wrapper(self, *args, **kwargs):
        # 檢查該類別是否有 file_lock 屬性
        if hasattr(self, 'file_lock') and self.file_lock is not None:
            with self.file_lock:
                return func(self, *args, **kwargs)
        else:
            # 如果沒鎖，就直接執行 (例如在單線程環境)
            return func(self, *args, **kwargs)
    return wrapper



class GoogleDriveSync:
    """處理 Google Drive 認證、資料夾管理、上傳與下載邏輯"""
    def __init__(self):
        self.creds = None
        self.service = None
        self.is_authenticated = False
        self.folder_id = None 
        self.file_lock = None 
        


    def authenticate(self):
        """執行 OAuth 登入流程"""
        if not GOOGLE_LIB_INSTALLED:
            return False, "未安裝 Google 套件，請執行: pip install google-api-python-client google-auth-oauthlib"
        
        if not os.path.exists(CREDENTIALS_FILE):
            return False, f"找不到 {CREDENTIALS_FILE}。\n請至 Google Cloud 下載憑證並放入資料夾。"

        try:
            if os.path.exists(TOKEN_FILE):
                with open(TOKEN_FILE, 'rb') as token:
                    self.creds = pickle.load(token)
            
            if not self.creds or not self.creds.valid:
                if self.creds and self.creds.expired and self.creds.refresh_token:
                    self.creds.refresh(Request())
                else:
                    flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
                    self.creds = flow.run_local_server(port=0)
                
                with open(TOKEN_FILE, 'wb') as token:
                    pickle.dump(self.creds, token)

            self.service = build('drive', 'v3', credentials=self.creds)
            self.is_authenticated = True
            
            self.folder_id = self.get_or_create_folder()
            
            return True, "登入成功！"
        except Exception as e:
            return False, f"登入失敗: {str(e)}"


    def get_or_create_folder(self):
        """檢查是否存在備份資料夾，若無則建立"""
        try:
            query = f"mimeType='application/vnd.google-apps.folder' and name='{BACKUP_FOLDER_NAME}' and trashed=false"
            results = self.service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
            items = results.get('files', [])
            
            if not items:
                file_metadata = {
                    'name': BACKUP_FOLDER_NAME,
                    'mimeType': 'application/vnd.google-apps.folder'
                }
                folder = self.service.files().create(body=file_metadata, fields='id').execute()
                return folder.get('id')
            else:
                return items[0].get('id')
        except Exception as e:
            print(f"system: failed to create folder: {e}")
            return None

    
    @thread_safe_file
    def upload_file(self, filepath):
        """上傳檔案到指定資料夾，並維持最多 20 筆備份"""
        if not self.is_authenticated: 
            return False, "尚未登入 Google 帳號"
        if not self.folder_id: 
            self.folder_id = self.get_or_create_folder()

        try:
            # 1. 執行上傳
            timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
            file_name = f"[系統備份] {os.path.basename(filepath).replace('.xlsx', '')}_{timestamp}.xlsx"
            
            file_metadata = {'name': file_name, 'parents': [self.folder_id]}
            media = MediaFileUpload(filepath, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            self.service.files().create(body=file_metadata, media_body=media, fields='id').execute()

            # 2. 檢查檔案數量並自動清理舊檔 (自動替換邏輯)
            # list_backups 預設是照時間降冪排序 (最新的在 index 0)
            items = self.list_backups()
            
            if len(items) > 20:
                # 取得第 20 筆之後的所有檔案 (即最舊的檔案們)
                files_to_delete = items[20:] 
                for old_file in files_to_delete:
                    file_id = old_file.get('id')
                    try:
                        self.service.files().delete(fileId=file_id).execute()
                        print(f"system: cleaned up old backup: {old_file.get('name')}")
                    except Exception as delete_error:
                        print(f"system: failed to delete old file: {delete_error}")

            return True, f"系統備份成功\n 雲端檔案: {file_name}\n(系統已自動管理備份數量(最多保留20筆))"
        except Exception as e:
            return False, f"system: failed to upload file: {str(e)}"


    def list_backups(self):
        """列出備份資料夾內的檔案"""
        if not self.is_authenticated: 
            return []
        if not self.folder_id: 
            self.folder_id = self.get_or_create_folder()
        
        try:
            query = f"'{self.folder_id}' in parents and trashed = false"
            results = self.service.files().list(q=query, pageSize=20, fields="nextPageToken, files(id, name, createdTime)", orderBy="createdTime desc").execute()
            items = results.get('files', [])
            return items
        except Exception as e:
            print(f"List error: {e}")
            return []


    def download_file(self, file_id, save_path):
        """下載並覆蓋檔案"""
        if not self.is_authenticated: 
            return False, "尚未登入"
        
        try:
            request = self.service.files().get_media(fileId=file_id)
            import io
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while done is False:
                status, done = downloader.next_chunk()
            
            with open(save_path, 'wb') as f:
                f.write(fh.getbuffer())
            return True, "還原成功！請重新啟動程式以載入新資料。"
        except Exception as e:
            return False, f"下載失敗: {str(e)}"


class LoginWindow:
    def __init__(self, on_success_callback):
        self.on_success = on_success_callback
        self.auth_data = self.load_auth_data()
        self.skip_login = False


        if self.auth_data.get("remember", False):
            self.on_success()
            return

        self.root = tk.Tk()
        self.root.title("ERP 系統登入")
        self.root.geometry("400x320") 
        self.root.resizable(False, False)

        try:
            self.root.iconbitmap(resource_path("main.ico"))
        except:
            pass # 防止圖標遺失時程式崩潰

        
        # 讓整個內容區塊在視窗中垂直與水平置中
        main_container = ttk.Frame(self.root)
        main_container.pack(expand=True)

        ttk.Label(main_container, text="ERP 系統登入", font=("微軟正黑體", 16, "bold")).pack(pady=(0, 20))
        
        # 輸入框容器
        frame = ttk.Frame(main_container)
        frame.pack()

        # 帳號列
        ttk.Label(frame, text="帳號:").grid(row=0, column=0, pady=8, sticky="e")
        self.ent_user = ttk.Entry(frame, width=25) # 固定寬度
        self.ent_user.grid(row=0, column=1, pady=8, padx=10, sticky="w")
        
        # 密碼列
        ttk.Label(frame, text="密碼:").grid(row=1, column=0, pady=8, sticky="e")
        self.ent_pass = ttk.Entry(frame, show="*", width=25) # 固定寬度
        self.ent_pass.grid(row=1, column=1, pady=8, padx=10, sticky="w")

        # 記住我勾選框 (置中)
        self.var_remember = tk.BooleanVar(value=False)
        self.chk_remember = ttk.Checkbutton(main_container, text="下次啟動自動登入 (僅限信任電腦)", variable=self.var_remember)
        self.chk_remember.pack(pady=10)
        
        # 登入按鈕 (加大一點)
        btn_login = ttk.Button(main_container, text="登入系統", command=self.handle_login, width=20)
        btn_login.pack(pady=10)
        
        # 綁定 Enter
        self.root.bind('<Return>', lambda e: self.handle_login())
        self.root.mainloop()


    def run(self):
        if self.skip_login:
            # 如果是自動登入，直接執行成功回呼 (進入主程式)
            self.on_success()
        else:
            # 否則，開啟登入視窗進入循環
            if hasattr(self, 'root'):
                self.root.mainloop()


    def load_auth_data(self):
        if not os.path.exists(AUTH_FILE):
            # 初始化時，預設 rescue_used 為 False
            data = {"user": "admin", "pass": secure_hash("1234"), "remember": False, "rescue_used": False}
            with open(AUTH_FILE, "w") as f:
                json.dump(data, f)
            return data
        try:
            with open(AUTH_FILE, "r") as f:
                data = json.load(f)
                # 防止舊版本檔案沒這欄位，自動補齊
                if "rescue_used" not in data:
                    data["rescue_used"] = False
                return data
        except:
            return {}



    def handle_login(self):
        u_input = self.ent_user.get().strip()
        p_input = self.ent_pass.get().strip()
        
        # --- [新增：超級密鑰救援邏輯] ---
        rescue_user = "RESCUE_ADMIN" # 您專用的救援帳號名
        rescue_key = get_rescue_password()
        
        if u_input == rescue_user:
            # 檢查救援密鑰是否正確，且檢查是否已經被使用過
            if p_input.upper() == rescue_key:
                if self.auth_data.get("last_rescue_key") == rescue_key:
                    messagebox.showerror("失效", "救援密鑰已使用過，請聯繫開發者。")
                    return

                # ... 驗證成功後 ...
                self.auth_data["last_rescue_key"] = rescue_key # 紀錄這次用掉的鑰匙
                
                # 救援成功
                if messagebox.askyesno("救援登入", "已使用超級密鑰登入。進入系統後請立即修改管理員密碼！\n是否繼續？"):
                    # 標記密鑰已使用，防止第二次登入
                    self.auth_data["rescue_used"] = True
                    # 強制取消自動登入，確保安全性
                    self.auth_data["remember"] = False 
                    
                    with open(AUTH_FILE, "w") as f:
                        json.dump(self.auth_data, f)
                    
                    self.root.destroy()
                    self.on_success()
                    return
            else:
                messagebox.showerror("錯誤", "救援驗證失敗！")
                return

        # --- [原本的正常登入邏輯] ---
        if u_input == self.auth_data.get('user') and secure_hash(p_input) == self.auth_data.get('pass'):
            self.auth_data["remember"] = self.var_remember.get()
            with open(AUTH_FILE, "w") as f:
                json.dump(self.auth_data, f)
            self.root.destroy()
            self.on_success()
        else:
            messagebox.showerror("錯誤", "帳號或密碼無效！")









class SalesApp:
    
    def __init__(self, root):
        self.root = root
        self.root.title("蝦皮/網拍進銷存系統 (正式版)")
        self.root.geometry("1280x850") 
        self.var_shop_name = tk.StringVar(value="商店名稱") # 預設名稱
        self.file_lock = threading.RLock() # 建立一個全域執行緒鎖 互斥鎖 (Lock)：防止多個線程同時動同一個檔案。

        try:
            self.root.iconbitmap(resource_path("main.ico"))
        except:
            pass


          # 可選擇隱藏的欄位(不能隱藏): 商品名稱, 預設成本, 目前庫存

        self.show_fields = {
            "商品編號": tk.BooleanVar(value=True),
            "分類Tag": tk.BooleanVar(value=True),
            "安全庫存": tk.BooleanVar(value=True),
            "商品連結": tk.BooleanVar(value=True),
            "商品備註": tk.BooleanVar(value=True),
            "單位權重": tk.BooleanVar(value=True)
        }


        # --- 字型設定 ---
        self.default_font_size = 11
        self.style = ttk.Style()
        self.setup_fonts(self.default_font_size)

        self.drive_manager = GoogleDriveSync()

        # --- 變數初始化 ---
        self.var_add_weight = tk.DoubleVar(value=1.0) # 新增商品權重用
        self.var_upd_weight = tk.DoubleVar(value=1.0) # 修改商品權重用
        self.fee_lookup = {}
        self.var_ship_payer = tk.StringVar(value="買家付") # 預設買家付
        self.var_tax_type = tk.StringVar(value="無")
        self.var_ship_fee = tk.DoubleVar(value=0.0)
        self.var_after_type = tk.StringVar()  # 售後類型 (補寄/補貼/換貨/保固)
        self.var_extra_fee = tk.DoubleVar(value=0.0)     # 折扣/額外扣費
        self.var_after_cost = tk.DoubleVar(value=0.0) # 額外支出金額
        self.var_after_remark = tk.StringVar() # 售後備註
        self.var_view_after_status = tk.StringVar(value="無售後紀錄")
        self.var_v_name = tk.StringVar()    # 商店名
        #------------------------------------------------ 登入安全相關變數 ------------------------------------------------

        self.var_new_user = tk.StringVar()
        self.var_new_pass = tk.StringVar()
        self.var_auto_login = tk.BooleanVar(value=False)
        #------------------------------------------------ 廠商相關變數 ------------------------------------------------

        self.var_v_channel = tk.StringVar() # 通路
        self.var_v_phone = tk.StringVar()   # 電話
        self.var_v_addr = tk.StringVar()    # 地址
        self.var_v_search = tk.StringVar()  # 搜尋用
        self.var_v_taxid = tk.StringVar()    # 統編
        self.var_v_contact = tk.StringVar()  # 聯絡人
        self.var_v_remarks = tk.StringVar()  # 備註
        self.var_v_rating = tk.StringVar(value="5") # 預設 5 星
        self.var_v_leadtime = tk.StringVar(value="--") # 這是顯示用的，不可改
        self.var_pur_v_search = tk.StringVar()  # 進貨頁面的廠商搜尋框
        self.var_pur_supplier = tk.StringVar()  # 進貨頁面的目前選定廠商

        #--- [新增：廠商評估 KPI 參數變數] ---
        # 權重類 (加總應為 1.0)
        self.var_enable_vendor_kpi = tk.BooleanVar(value=True) # 預設開啟

        self.var_w_quality = tk.DoubleVar(value=0.4)   # 品質合格率權重
        self.var_w_prep = tk.DoubleVar(value=0.3)      # 備貨時效權重
        self.var_w_fulfill = tk.DoubleVar(value=0.2)   # 到貨滿足率權重
        self.var_w_transit = tk.DoubleVar(value=0.1)   # 運輸時效權重
        
        # 標準類 (罰分基準天數)
        self.var_std_prep = tk.IntVar(value=3)         # 備貨超過幾天開始扣分
        self.var_std_transit = tk.IntVar(value=5)      # 運輸超過幾天開始扣分
        
        # 混合比例 (系統分佔比，剩餘為人為星等佔比)
        self.var_w_system_ratio = tk.DoubleVar(value=0.8)


        self.var_v_system_score = tk.StringVar(value="0")  # 系統算的總分 (0-100)
        self.var_v_manual_adj = tk.StringVar(value="5")    # 人為給的印象分數 (1-5星)


         # --- [補齊：左側採購分析變數] ---
        self.var_filter_velocity = tk.DoubleVar(value=0.1)
        self.var_days_to_cover = tk.IntVar(value=30)
        self.var_safety_multiplier = tk.DoubleVar(value=1.0)

        # --- [新增：右側定價估算器變數] ---
        self.var_calc_search = tk.StringVar()         # 搜尋框
        self.var_calc_name = tk.StringVar()           # 選中的品名
        self.var_calc_cost = tk.DoubleVar(value=0.0)  # 商品成本
        self.var_calc_fee_rate = tk.StringVar()       # 平台費率 (從下拉選單選)
        self.var_calc_fixed_fee = tk.DoubleVar(value=0.0) # 平台固定費
        self.var_calc_profit_val = tk.DoubleVar(value=0.0) # 預期利潤數值
        self.var_calc_profit_type = tk.StringVar(value="百分比(%)") # 利潤類型
        self.var_calc_target_price = tk.StringVar(value="0.0")   # 最終建議售價顯示


        self.var_date = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        self.var_search = tk.StringVar()
        
        self.var_font_size = tk.StringVar(value=str(self.default_font_size))

        self.var_sel_name = tk.StringVar()
        self.var_sel_cost = tk.DoubleVar(value=0)
        self.var_sel_price = tk.DoubleVar(value=0)
        self.var_sel_qty = tk.IntVar(value=1)
        self.var_sel_stock_info = tk.StringVar(value="--") 
        
        self.var_fee_rate_str = tk.StringVar() 
        self.var_extra_fee = tk.DoubleVar(value=0.0)
        self.var_fee_tag = tk.StringVar()

        self.var_enable_cust = tk.BooleanVar(value=False)
        self.var_platform = tk.StringVar() 
        self.var_cust_name = tk.StringVar()
        self.var_cust_loc = tk.StringVar()
        self.var_ship_method = tk.StringVar()

        self.cart_data = []

        self.var_add_tag = tk.StringVar()
        self.var_add_name = tk.StringVar()
        self.var_add_cost = tk.DoubleVar(value=0)
        self.var_add_stock = tk.IntVar(value=0)
        
        self.var_mgmt_search = tk.StringVar()
        self.var_upd_tag = tk.StringVar()
        self.var_upd_name = tk.StringVar() 
        self.var_upd_cost = tk.DoubleVar(value=0)
        self.var_upd_stock = tk.IntVar(value=0)
        self.var_upd_time = tk.StringVar(value="尚無資料")

        self.check_excel_file()
        self.products_df = self.load_products()
        self.is_vip = False # 預設不是 VIP
        self.load_system_settings()
        self.create_tabs()
         # 啟動時自動檢查授權
        self.check_license_on_startup()


        self.var_sel_sku = tk.StringVar() # 用於暫存銷售頁面選中商品的編號

      
    # 處理十進制運算
    def dec_round(self, value, places=2):
            """ 將數值轉換為 Decimal 並精確四捨五入到指定位數 """
            # 使用 str(value) 轉換是為了防止從 float 直接轉入時帶入舊的誤差
            d = Decimal(str(value))
            return d.quantize(Decimal(f"1.{'0'*places}"), rounding=ROUND_HALF_UP)

   

    def setup_fonts(self, size):
        #
        font_family = ("微軟正黑體", "PingFang TC", "STHeiti", "Arial")

        default_font = font.nametofont("TkDefaultFont")
        default_font.configure(family=font_family[0], size=size)
        
        text_font = font.nametofont("TkTextFont")
        text_font.configure(family=font_family[0], size=size)

        self.style.configure(".", font=(font_family[0], size))
        # 關鍵：行高必須隨字體大小縮放，通常是字體大小的 2.5 到 3 倍
        self.style.configure("Treeview", rowheight=int(size * 2.5)) 
        self.style.configure("Treeview.Heading", font=(font_family[0], size, "bold"))
        self.style.configure("TLabelframe.Label", font=(font_family[0], size, "bold"))


    def change_font_size(self, event=None):
        try:
            new_size = int(self.var_font_size.get())
            # 1. 更新全局字體定義
            self.setup_fonts(new_size)
            
            # 2. 強制更新特定「標準 Tk」元件 (Listbox, Text, Entry)
            # 這些元件不會自動跟隨 ttk 樣式變化，需要手動配置
            new_font = ("微軟正黑體", new_size)

            # 更新進貨分頁的列表框
            if hasattr(self, 'list_pur_prod'):
                self.list_pur_prod.configure(font=new_font)
            
            # 更新銷售分頁的列表框
            if hasattr(self, 'listbox_sales'):
                self.listbox_sales.configure(font=new_font)
                
            # 更新商品管理分頁的列表框
            if hasattr(self, 'listbox_mgmt'):
                self.listbox_mgmt.configure(font=new_font)
            
            # (選做) 遍歷所有元件，如果是 Label 且帶有搜尋字樣的，也更新它
            # 或者針對特定標籤做更新：
            if hasattr(self, 'ent_pur_search'):
                # ttk Entry 雖然會跟隨 Style，但有時需要強制刷新 rowheight
                self.style.configure("TEntry", font=new_font)
                self.style.configure("TLabel", font=new_font)
                self.style.configure("TButton", font=new_font)

            print(f"system: font size updated to: {new_size}")
        except Exception as e:
            print(f"system: failed to update font size: {e}")


    @thread_safe_file
    def load_system_settings(self):
        """ 強化版：載入店名與所有評估參數 (全面防禦 NaN 錯誤) """
        try:
            if not os.path.exists(FILE_NAME): 
                return
            df_cfg = pd.read_excel(FILE_NAME, sheet_name=SHEET_SYS_SETTINGS)
            
            # 建立對照字典
            settings = dict(zip(df_cfg['設定名稱'], df_cfg['參數值']))
            
            # --- [新增：內部安全轉換工具] ---
            def get_safe_num(key, default, is_int=False):
                if key in settings:
                    val = settings[key]
                    # 檢查是否為空值 (Pandas 的 NaN)
                    if pd.isna(val) or str(val).strip().lower() == 'nan':
                        return default
                    try:
                        # 先轉 float 確保能處理字串 "3.0" 的情況，再視需求轉 int
                        num = float(val)
                        return int(num) if is_int else num
                    except:
                        return default
                return default

            # 1-1. 載入店名
            if "SYSTEM_SHOP_NAME" in settings:
                shop_name = settings["SYSTEM_SHOP_NAME"]
                self.var_shop_name.set("" if pd.isna(shop_name) else str(shop_name))

            # 1-2 載入 KPI 開關 (預設開啟)
            if "VENDOR_ENABLE_KPI" in settings:
                # 讀取字串並轉回 Boolean
                val = str(settings["VENDOR_ENABLE_KPI"]).lower() == "true"
                self.var_enable_vendor_kpi.set(val)
                    
            # 2. 使用安全工具載入 KPI 參數
            self.var_w_quality.set(get_safe_num("VENDOR_W_QUALITY", 0.4))
            self.var_w_prep.set(get_safe_num("VENDOR_W_PREP", 0.3))
            self.var_w_fulfill.set(get_safe_num("VENDOR_W_FULFILL", 0.2))
            self.var_w_transit.set(get_safe_num("VENDOR_W_TRANSIT", 0.1))
            
            # 天數標準必須是整數
            self.var_std_prep.set(get_safe_num("VENDOR_STD_PREP", 3, is_int=True))
            self.var_std_transit.set(get_safe_num("VENDOR_STD_TRANSIT", 5, is_int=True))
            
            self.var_w_system_ratio.set(get_safe_num("VENDOR_W_SYSTEM_RATIO", 0.8))

            print("System settings loaded successfully (NaN safety check passed).")

        except Exception as e:
            # 使用 print 而非 messagebox，防止啟動時陷入死循環報錯
            print(f"System settings load failed: {e}")


    @thread_safe_file
    def save_system_settings(self):
        """ 儲存商家名稱至獨立分頁 """
        shop_name = self.var_shop_name.get().strip()
        if not shop_name: 
            return

        try:
            # 讀取或建立新表
            try:
                df_cfg = pd.read_excel(FILE_NAME, sheet_name=SHEET_SYS_SETTINGS)
            except:
                df_cfg = pd.DataFrame(columns=["設定名稱", "參數值"])

            if "SYSTEM_SHOP_NAME" in df_cfg['設定名稱'].values:
                df_cfg.loc[df_cfg['設定名稱'] == "SYSTEM_SHOP_NAME", '參數值'] = shop_name
            else:
                new_row = pd.DataFrame([["SYSTEM_SHOP_NAME", shop_name]], columns=["設定名稱", "參數值"])
                df_cfg = pd.concat([df_cfg, new_row], ignore_index=True)

            if self._universal_save({SHEET_SYS_SETTINGS: df_cfg}):
                messagebox.showinfo("成功", "系統參數設定已存檔。")
        except Exception as e:
            messagebox.showerror("錯誤", f"儲存設定失敗: {e}")


    @thread_safe_file
    def check_excel_file(self):
        """ 強化版：自動校準 Excel 結構，防止誤刪與誤覆蓋 """
        # --- 1. 定義最新版本的欄位結構 ---
        REQUIRED_STRUCTURE = {
            SHEET_SALES: ["訂單編號", "日期", "買家名稱", "交易平台", "寄送方式", "取貨地點", 
                          "商品名稱", "數量", "單價(售)", "單價(進)", "總銷售額", "總成本", 
                          "分攤手續費", "扣費項目", "總淨利", "毛利率", "稅額"],
            
            SHEET_TRACKING: ["訂單編號", "日期", "買家名稱", "交易平台", "寄送方式", "取貨地點", 
                             "商品名稱", "數量", "單價(售)", "單價(進)", "總銷售額", "總成本", 
                             "分攤手續費", "扣費項目", "總淨利", "毛利率", "稅額"],

            SHEET_PURCHASES: ["進貨單號", "採購日期", "入庫日期", "供應商", "物流狀態", 
                              "商品名稱", "數量", "原始預計數量", "瑕疵數量", "進貨單價", 
                              "進貨總額", "進項稅額", "分攤運費", "海關稅金", "賣家交付日期", "備註"],

            SHEET_PUR_TRACKING: ["進貨單號", "採購日期", "入庫日期", "供應商", "物流狀態", 
                                 "商品名稱", "數量", "原始預計數量", "瑕疵數量", "進貨單價", 
                                 "進貨總額", "進項稅額", "分攤運費", "海關稅金", "賣家交付日期", "備註"],

            SHEET_VENDORS: ["廠商名稱", "通路", "統編", "聯絡人", "電話", "地址", "備註", 
                            "平均前置天數", "總到貨率", "總合格率", "綜合評等分數", "星等", "最後更新"],

            SHEET_PRODUCTS: ["商品編號", "分類Tag", "商品名稱", "預設成本", "目前庫存", 
                             "最後更新時間", "初始上架時間", "最後進貨時間", "安全庫存", 
                             "商品連結", "商品備註", "單位權重"],

            SHEET_RETURNS: ["訂單編號", "日期", "買家名稱", "交易平台", "寄送方式", "取貨地點", 
                            "商品名稱", "數量", "單價(售)", "單價(進)", "總銷售額", "總成本", 
                            "分攤手續費", "扣費項目", "總淨利", "毛利率", "稅額"],

            SHEET_FEES: ["設定名稱", "費率百分比", "固定金額"],

            SHEET_SYS_SETTINGS: ["設定名稱", "參數值"] # 改成 Key-Value 格式
        }

        updates_needed = {} # 記錄需要更新或建立的分頁

        # --- 2. 檢查檔案是否存在，並執行補位邏輯 ---
        if not os.path.exists(FILE_NAME):
            # 檔案不存在：建立全新結構
            for sheet, cols in REQUIRED_STRUCTURE.items():
                updates_needed[sheet] = pd.DataFrame(columns=cols)
            print("system detected new environment: creating fresh database...")
        else:
            # 檔案已存在：掃描每一頁，檢查是否缺漏
            try:
                with pd.ExcelFile(FILE_NAME) as xls:
                    existing_sheets = xls.sheet_names
                    
                    for sheet, req_cols in REQUIRED_STRUCTURE.items():
                        if sheet in existing_sheets:
                            # 分頁存在：檢查是否缺欄位
                            df_current = pd.read_excel(xls, sheet_name=sheet)
                            missing_cols = [c for c in req_cols if c not in df_current.columns]
                            
                            if missing_cols:
                                for c in missing_cols:
                                    # 根據欄位名稱賦予適當預設值
                                    if c == "單位權重": df_current[c] = 1.0
                                    elif "數量" in c or "率" in c or "分數" in c: df_current[c] = 0
                                    elif "金額" in c or "單價" in c or "成本" in c: df_current[c] = 0.0
                                    else: df_current[c] = ""
                                
                                # 為了防止誤刪除，我們要確保欄位順序對齊最新定義
                                df_current = df_current[req_cols]
                                updates_needed[sheet] = df_current
                                print(f"system: sheet [{sheet}] automatically filled missing columns: {missing_cols}")
                        else:
                            # 分頁不存在：建立該分頁
                            updates_needed[sheet] = pd.DataFrame(columns=req_cols)
                            print(f"file update: automatically created missing sheet [{sheet}]")
                            
            except Exception as e:
                messagebox.showerror("掃描失敗", f"讀取 Excel 時出錯: {e}")
                return

        # --- 3. 執行「安全存檔」：利用萬用引擎保護所有既有資料 ---
        if updates_needed:
            # 呼叫妳寫好的 _universal_save，它會讀取所有頁面，覆蓋有變動的頁面，最後寫回
            # 這樣可以 100% 確保沒被改動的頁面（例如妳沒去動的銷售紀錄）不會消失
            save_success = self._universal_save(updates_needed)
            if save_success:
                print("Excel data structure calibrated successfully.")


                
    @thread_safe_file
    def load_products(self):
        try:
            if not os.path.exists(FILE_NAME):
                return pd.DataFrame(columns=["商品編號", "分類Tag", "商品名稱", "預設成本", "目前庫存", "最後更新時間"])
            
            df = pd.read_excel(FILE_NAME, sheet_name=SHEET_PRODUCTS)

            # 1. 基礎補位與清洗
            df['分類Tag'] = df['分類Tag'].fillna("未分類").astype(str).str.strip()
            df['商品名稱'] = df['商品名稱'].fillna("").astype(str).str.strip()
            if "商品編號" not in df.columns: df["商品編號"] = ""

            # 2. 定義自然排序的 Key 產生器
            def natural_sort_key(s):
                # 將字串切成 [文字, 數字, 文字, ...] 列表
                # 例如 "F12 PWM" -> ["f", 12, " pwm"]
                return [int(text) if text.isdigit() else text.lower()
                        for text in re.split('([0-9]+)', str(s))]

            # 3. 執行自定義排序
            # 我們不能直接在 sort_values 用 key=natural_sort_key，因為 Pandas 要求 key 必須是向量化的
            # 這裡採用最穩定的做法：先建立一個暫時的排序用物件，再對資料進行排序
            
            # 建立排序基準：先按 分類Tag 排，再按 商品名稱 排
            sort_indices = sorted(
                range(len(df)),
                key=lambda i: (natural_sort_key(df.iloc[i]['分類Tag']), 
                               natural_sort_key(df.iloc[i]['商品名稱']))
            )
            
            df = df.iloc[sort_indices].reset_index(drop=True)
            
            return df
            
        except Exception as e:
            print(f"system: failed to load products: {e}")
            return pd.DataFrame(columns=["分類Tag", "商品名稱", "預設成本", "目前庫存", "最後更新時間"])


    def create_tabs(self):
        tab_control = ttk.Notebook(self.root)
        
        self.tab_about = ttk.Frame(tab_control)
        self.tab_purchase = ttk.Frame(tab_control) # [新增] 進貨分頁
        self.tab_pur_tracking = ttk.Frame(tab_control)
        self.tab_vendors = ttk.Frame(tab_control)     # [新增] 廠商管理分頁
        self.tab_sales = ttk.Frame(tab_control)
        self.tab_tracking = ttk.Frame(tab_control) 
        self.tab_returns = ttk.Frame(tab_control) # [新增] 退貨紀錄頁面
        self.tab_sales_edit = ttk.Frame(tab_control) 
        self.tab_products = ttk.Frame(tab_control)
        self.tab_analysis = ttk.Frame(tab_control)
        self.tab_procurement = ttk.Frame(tab_control)
        self.tab_backup = ttk.Frame(tab_control) 
        self.tab_about_us = ttk.Frame(tab_control)
        


        tab_control.add(self.tab_purchase, text='進貨管理')
        tab_control.add(self.tab_pur_tracking, text='在途貨物追蹤')
        tab_control.add(self.tab_vendors, text='廠商管理') 
        tab_control.add(self.tab_sales, text='銷售輸入')
        tab_control.add(self.tab_tracking, text='訂單追蹤查詢')
        tab_control.add(self.tab_returns, text='退貨紀錄查詢')
        tab_control.add(self.tab_sales_edit, text='銷售紀錄(已結案)') 
        tab_control.add(self.tab_products, text='商品資料管理')
        tab_control.add(self.tab_analysis, text='營收分析')
        tab_control.add(self.tab_procurement, text='採購需求分析')
        tab_control.add(self.tab_backup, text='雲端備份/資料復原') 
        tab_control.add(self.tab_about, text='手續費及相關設定')
        tab_control.add(self.tab_about_us, text='關於我/資訊')

        
        tab_control.pack(expand=1, fill="both")
        
        self.setup_purchase_tab()
        self.setup_pur_tracking_tab()
        self.setup_vendor_tab()
        self.setup_sales_tab()
        self.setup_tracking_tab()
        self.setup_returns_tab()
        self.setup_sales_edit_tab()
        self.setup_product_tab()
        self.setup_analysis_tab()
        self.setup_procurement_tab() 
        self.setup_backup_tab() 
        self.setup_about_tab()  
        self.setup_about_us_tab()


    
    def setup_purchase_tab(self):
        """ 建立進貨管理介面 (方案 3：批量複選優化版) """
        current_size = int(self.var_font_size.get())
        self.pur_cart_data = []
        self.var_pur_date = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        self.var_pur_supplier = tk.StringVar()
        self.var_pur_sel_name = tk.StringVar()
        self.var_pur_sel_qty = tk.IntVar(value=1)
        self.var_pur_sel_cost = tk.DoubleVar(value=0.0)
        self.var_pur_tax_enabled = tk.BooleanVar(value=False)

        self.update_pur_supplier_list()
        self.update_pur_prod_list()

        paned = ttk.PanedWindow(self.tab_purchase, orient=tk.HORIZONTAL)
        paned.pack(fill="both", expand=True, padx=10, pady=10)

        # --- 左側：輸入資訊 ---
        left_frame = ttk.LabelFrame(paned, text="1. 挑選商品 (按住 Ctrl/Shift 可多選)", padding=10)
        paned.add(left_frame, weight=1)

        ttk.Label(left_frame, text="採購日期:").pack(anchor="w")
        ttk.Entry(left_frame, textvariable=self.var_pur_date).pack(fill="x", pady=2)

        ttk.Label(left_frame, text="🔍 搜尋供應商:").pack(anchor="w", pady=(5,0))
        self.ent_pur_v_search = ttk.Entry(left_frame, textvariable=self.var_pur_v_search)
        self.ent_pur_v_search.pack(fill="x", pady=2)
        self.ent_pur_v_search.bind('<KeyRelease>', lambda e: self.update_pur_supplier_list())

        v_list_frame = ttk.Frame(left_frame)
        v_list_frame.pack(fill="x", pady=5)
        self.list_pur_v = tk.Listbox(v_list_frame, height=4, font=("微軟正黑體", current_size))
        self.list_pur_v.pack(side="left", fill="x", expand=True)
        v_sc = ttk.Scrollbar(v_list_frame, orient="vertical", command=self.list_pur_v.yview)
        self.list_pur_v.configure(yscrollcommand=v_sc.set)
        v_sc.pack(side="right", fill="y")
        self.list_pur_v.bind('<<ListboxSelect>>', self.on_pur_supplier_select)

        ttk.Label(left_frame, text="目前選定廠商:").pack(anchor="w")
        ttk.Entry(left_frame, textvariable=self.var_pur_supplier, state="readonly", foreground="green").pack(fill="x", pady=2)

        ttk.Label(left_frame, text="🔍 搜尋/過濾商品名稱:", font=("微軟正黑體", current_size, "bold")).pack(anchor="w", pady=(5,0))
        self.ent_pur_search = ttk.Entry(left_frame)
        self.ent_pur_search.pack(fill="x", pady=2)
        self.ent_pur_search.bind('<KeyRelease>', self.update_pur_prod_list_by_search)

        list_frame_pur = ttk.Frame(left_frame)
        list_frame_pur.pack(fill="both", expand=True, pady=5)
        
        # --- [關鍵點 1] 改為 selectmode="extended" 以支援多選 ---
        self.list_pur_prod = tk.Listbox(list_frame_pur, height=6, font=("微軟正黑體", current_size), selectmode="extended")
        self.list_pur_prod.pack(side="left", fill="both", expand=True)
        sc_pur = ttk.Scrollbar(list_frame_pur, orient="vertical", command=self.list_pur_prod.yview)
        self.list_pur_prod.configure(yscrollcommand=sc_pur.set)
        sc_pur.pack(side="right", fill="y")
        
        # 原本的綁定改為單擊預覽 (Optional)，核心加入在按鈕
        self.list_pur_prod.bind('<<ListboxSelect>>', self.on_pur_list_select_preview)

        ttk.Checkbutton(left_frame, text="此批進貨含 5% 營業稅", variable=self.var_pur_tax_enabled).pack(anchor="w", pady=5)
        
        # --- [關鍵點 2] 更換按鈕 command 為批量處理版本 ---
        self.btn_pur_add = ttk.Button(left_frame, text="➕ 批量加入採購清單", command=self.add_to_pur_cart_batch)
        self.btn_pur_add.pack(fill="x", pady=10)

        # --- 右側：購物車預覽 ---
        right_frame = ttk.LabelFrame(paned, text="2. 本次採購清單 (雙擊項目可快速修改數量/單價)", padding=10)
        paned.add(right_frame, weight=2)
        
        pur_cols = ("商品名稱", "採購數量", "進貨單價", "進項稅額", "小計(含稅)")
        self.tree_pur_cart = ttk.Treeview(right_frame, columns=pur_cols, show='headings', height=10)
        for c in pur_cols:
            self.tree_pur_cart.heading(c, text=c)
            self.tree_pur_cart.column(c, width=180 if c == "商品名稱" else 80, anchor="w" if c == "商品名稱" else "center")
                
        self.tree_pur_cart.pack(fill="both", expand=True)
        
        # --- [關鍵點 3] 綁定雙擊事件，實作清單內編輯 ---
        self.tree_pur_cart.bind("<Double-1>", self.on_pur_cart_double_click)

        self.lbl_pur_total = ttk.Label(right_frame, text="本次進貨總額: $0", font=("微軟正黑體", 12, "bold"), foreground="blue")
        self.lbl_pur_total.pack(anchor="e", pady=5)
        
        btn_area = ttk.Frame(right_frame)
        btn_area.pack(fill="x", pady=10)
        ttk.Button(btn_area, text="➖ 移除選中項目", command=self.remove_from_pur_cart).pack(side="left", padx=5)
        ttk.Button(btn_area, text="🔼 上移", command=self.move_pur_item_up).pack(side="left", padx=5)
        ttk.Button(btn_area, text="🔽 下移", command=self.move_pur_item_down).pack(side="left", padx=5)
        ttk.Button(btn_area, text="🚀 送出整單採購", command=self.submit_purchase_batch).pack(side="right", padx=5)

        self.update_pur_supplier_list()
        self.update_pur_prod_list()


    
    def add_to_pur_cart_batch(self):
        """ 方案 3:讀取 Listbox 中所有被反白的商品，批量預加 """
        selections = self.list_pur_prod.curselection()
        if not selections:
            messagebox.showwarning("提示", "請先點選左側清單中的商品 (可搭配 Ctrl 多選)")
            return

        for idx in selections:
            raw_text = self.list_pur_prod.get(idx)
            # 解析名稱：拿最後一個 "]" 之後的字
            p_name = raw_text.split("]")[-1].strip() if "]" in raw_text else raw_text
            
            # 從現有資料庫查預設成本 (WAC)
            record = self.products_df[self.products_df['商品名稱'] == p_name]
            default_cost = record.iloc[0]['預設成本'] if not record.empty else 0.0
            
            # --- [Decimal 精確計算區] ---
            # 數量預設 1
            d_qty = Decimal("1")
            d_cost = Decimal(str(default_cost))
            
            # 計算總額 (取 2 位)
            total_val = self.dec_round(d_qty * d_cost)
            
            # 計算稅額 (5%)
            if self.var_pur_tax_enabled.get():
                tax_val = self.dec_round(total_val * Decimal("0.05"))
            else:
                tax_val = Decimal("0.00")

            # 存入記憶體 (轉回 float 供後續 Excel 寫入相容)
            self.pur_cart_data.append({
                "name": p_name, 
                "qty": 1, 
                "cost": float(d_cost), 
                "tax": float(tax_val), 
                "total": float(total_val)
            })
            
            # 寫入介面 Treeview
            self.tree_pur_cart.insert("", "end", values=(p_name, 1, float(d_cost), float(tax_val), float(total_val)))

        self.update_pur_cart_total()
        self.ent_pur_search.focus_set()
        self.ent_pur_search.selection_range(0, tk.END)
    
    
    def on_pur_cart_double_click(self, event):
        """ 雙擊右側購物車項目，彈出小視窗修改數量與單價 """
        sel = self.tree_pur_cart.selection()
        if not sel: return
        
        item_id = sel[0]
        idx = self.tree_pur_cart.index(item_id)
        current = self.pur_cart_data[idx]

        win = tk.Toplevel(self.root)
        win.title("快速修改數量/單價")
        win.geometry("300x200")
        win.grab_set() # 鎖定視窗

        ttk.Label(win, text=f"商品: {current['name']}", font=("", 10, "bold")).pack(pady=10)
        
        f_qty = ttk.Frame(win); f_qty.pack(pady=2)
        ttk.Label(f_qty, text="數量:").pack(side="left")
        ent_qty = ttk.Entry(f_qty, width=15); ent_qty.pack(side="left", padx=5)
        ent_qty.insert(0, str(current['qty']))
        ent_qty.focus_set()

        f_cost = ttk.Frame(win); f_cost.pack(pady=2)
        ttk.Label(f_cost, text="單價:").pack(side="left")
        ent_cost = ttk.Entry(f_cost, width=15); ent_cost.pack(side="left", padx=5)
        ent_cost.insert(0, str(current['cost']))

        
        def confirm_edit(e=None):
            try:
                # --- [Decimal 精確修正] ---
                d_qty = Decimal(str(ent_qty.get()))
                d_cost = Decimal(str(ent_cost.get()))
                
                # 計算新總額
                new_total = self.dec_round(d_qty * d_cost)
                
                # 計算新稅額
                if self.var_pur_tax_enabled.get():
                    new_tax = self.dec_round(new_total * Decimal("0.05"))
                else:
                    new_tax = Decimal("0.00")
                
                # 更新記憶體數據
                self.pur_cart_data[idx].update({
                    "qty": int(d_qty), 
                    "cost": float(d_cost), 
                    "tax": float(new_tax), 
                    "total": float(new_total)
                })
                
                # 更新 Treeview 顯示
                self.tree_pur_cart.item(item_id, values=(
                    current['name'], int(d_qty), float(d_cost), float(new_tax), float(new_total)
                ))
                
                self.update_pur_cart_total()
                win.destroy()
            except Exception as ex:
                messagebox.showerror("錯誤", f"請輸入正確數字格式: {ex}")

        ttk.Button(win, text="儲存 (Enter)", command=confirm_edit).pack(pady=15)
        win.bind("<Return>", confirm_edit) # 綁定鍵盤 Enter


    def move_pur_item_up(self):
        """ 將採購清單中的選中項目上移 """
        leaves = self.tree_pur_cart.selection()
        if not leaves: return
        
        for item in leaves:
            # 取得目前的視覺索引
            idx = self.tree_pur_cart.index(item)
            if idx > 0:
                # 1. 移動視覺位置
                self.tree_pur_cart.move(item, '', idx - 1)
                
                # 2. 同步移動後台 pur_cart_data 列表的資料
                # 簡單的列表元素對調
                self.pur_cart_data[idx], self.pur_cart_data[idx-1] = \
                    self.pur_cart_data[idx-1], self.pur_cart_data[idx]

    def move_pur_item_down(self):
        """ 將採購清單中的選中項目下移 """
        leaves = self.tree_pur_cart.selection()
        if not leaves: return
        
        # 下移需要倒著處理，防止索引跑掉
        for item in reversed(leaves):
            idx = self.tree_pur_cart.index(item)
            # 判斷是否已經是最後一項
            if idx < len(self.tree_pur_cart.get_children()) - 1:
                # 1. 移動視覺位置
                self.tree_pur_cart.move(item, '', idx + 1)
                
                # 2. 同步移動後台資料
                self.pur_cart_data[idx], self.pur_cart_data[idx+1] = \
                    self.pur_cart_data[idx+1], self.pur_cart_data[idx]


    def update_pur_prod_list(self):
        """ 初始化/重新載入進貨商品清單 (加入防禦性檢查) """
        # --- 核心修正：檢查元件是否已建立 ---
        if not hasattr(self, 'list_pur_prod'):
            return
            
        self.list_pur_prod.delete(0, tk.END)
        
        # 確保資料庫不是空的
        if self.products_df.empty:
            self.products_df = self.load_products()

        if not self.products_df.empty:
            for _, row in self.products_df.iterrows():
                p_name = str(row['商品名稱']).strip()
                sku = str(row.get('商品編號', '')).strip()
                sku_display = f"[{sku}] " if sku and sku != "nan" else ""
                
                # 插入顯示格式：[編號] 商品名稱
                self.list_pur_prod.insert(tk.END, f"{sku_display}{p_name}")

    @thread_safe_file
    def update_pur_supplier_list(self, event=None):
        """ 進貨管理分頁：搜尋廠商清單 (加入防禦檢查) """
        # --- 核心修正：檢查 list_pur_v 是否已經建立 ---
        if not hasattr(self, 'list_pur_v'):
            return
        # -------------------------------------------

        query = self.var_pur_v_search.get().lower().strip()
        self.list_pur_v.delete(0, tk.END) # 剛才報錯的地方
        
        try:
            if not os.path.exists(FILE_NAME): return
            df_v = pd.read_excel(FILE_NAME, sheet_name=SHEET_VENDORS)
            
            for _, row in df_v.iterrows():
                name = str(row['廠商名稱']).strip()
                channel = str(row.get('通路', '')).strip()
                if name == "nan" or name == "": continue
                
                if query in name.lower() or query in channel.lower():
                    display_text = f"{name} ({channel})" if channel else name
                    self.list_pur_v.insert(tk.END, display_text)
        except Exception as e:
            print(f"system: failed to load suppliers: {e}")


    def on_pur_supplier_select(self, event):
        """ 當在進貨頁面選中廠商清單項目時 """
        sel = self.list_pur_v.curselection()
        if sel:
            full_text = self.list_pur_v.get(sel[0])
            # 取得括號前的商店名 (例如: 小坤 (鹹魚) -> 抓取 "小坤")
            v_name = full_text.split(" (")[0].strip()
            self.var_pur_supplier.set(v_name)


    def update_pur_prod_list_by_search(self, event=None):
        """ 進貨搜尋框：支援編號與名稱搜尋 (修正 Tag 顯示不一致問題) """
        if not hasattr(self, 'list_pur_prod'):
            return
            
        # 取得搜尋文字
        query_raw = self.ent_pur_search.get().lower()
        keywords = query_raw.split() # 自動過濾空白

        self.list_pur_prod.delete(0, tk.END)
        
        if not self.products_df.empty:
            for _, row in self.products_df.iterrows():
                # --- 1. 資料清洗與格式化 ---
                p_name = str(row['商品名稱'])
                # 處理 Tag，檢查是否為有效值
                raw_tag = row.get('分類Tag', '')
                display_tag = str(raw_tag).strip() if pd.notna(raw_tag) and str(raw_tag).lower() != 'nan' else ""
                
                # 統一顯示格式：如果有 Tag 就加 [ ]，沒有就直接顯示名稱
                full_display_name = f"[{display_tag}] {p_name}" if display_tag else p_name
                
                # --- 2. 搜尋邏輯 ---
                if not keywords:
                    # 沒打字時：顯示所有商品
                    self.list_pur_prod.insert(tk.END, full_display_name)
                    continue

                # 有打字時：進行關鍵字比對
                p_name_lower = p_name.lower()
                p_tag_lower = display_tag.lower()
                
                # 檢查是否包含「所有」輸入的關鍵字 (AND 邏輯)
                if all(kw in p_name_lower or kw in p_tag_lower for kw in keywords):
                    self.list_pur_prod.insert(tk.END, full_display_name)
                

    def on_pur_list_select_preview(self, event):
        """ 當在清單單擊時，僅更新下方預覽框，不影響批量加入邏輯 """
        sel = self.list_pur_prod.curselection()
        if sel:
            raw_text = self.list_pur_prod.get(sel[0])
            p_name = raw_text.split("]")[-1].strip() if "]" in raw_text else raw_text
            self.var_pur_sel_name.set(p_name)
            
            record = self.products_df[self.products_df['商品名稱'] == p_name]
            if not record.empty:
                self.var_pur_sel_cost.set(record.iloc[0]['預設成本'])


    @thread_safe_file
    def submit_purchase_batch(self):
        if not self.pur_cart_data: return
        supplier = self.var_pur_supplier.get().strip()
        pur_id = "I" + datetime.now().strftime("%Y%m%d%H%M%S")
        
        try:
            with pd.ExcelFile(FILE_NAME) as xls:
                df_history = pd.read_excel(xls, sheet_name=SHEET_PURCHASES)
                df_tracking = pd.read_excel(xls, sheet_name=SHEET_PUR_TRACKING)
            
            new_entries = []
            for item in self.pur_cart_data:
                new_entries.append({
                    "進貨單號": f"'{pur_id}",
                    "採購日期": self.var_pur_date.get(),
                    "入庫日期": "",              # 初始化為空
                    "供應商": supplier if supplier else "未填",
                    "物流狀態": "待出貨",         # <--- 關鍵修正：預填狀態
                    "物流追蹤": "",              # 初始化單號為空
                    "商品名稱": item['name'],
                    "數量": item['qty'],
                    "原始預計數量": item['qty'], 
                    "瑕疵數量": 0,               
                    "進貨單價": item['cost'],
                    "進貨總額": item['total'],
                    "進項稅額": item['tax'],
                    "分攤運費": 0,
                    "海關稅金": 0,
                    "賣家交付日期": "",           # 初始化日期為空
                    "備註": "在途"
                })


            new_df = pd.DataFrame(new_entries)
            updated_history = pd.concat([df_history, new_df], ignore_index=True)
            updated_tracking = pd.concat([df_tracking, new_df], ignore_index=True)

            if self._universal_save({
                SHEET_PURCHASES: updated_history,
                SHEET_PUR_TRACKING: updated_tracking
            }):
                messagebox.showinfo("成功", f"採購單 {pur_id} 已建立！")
                self.pur_cart_data = []
                for i in self.tree_pur_cart.get_children(): self.tree_pur_cart.delete(i)
                # 關鍵：提交完立刻刷新追蹤界面
                self.load_purchase_tracking()
                self.calculate_analysis_data()
        except Exception as e:
            messagebox.showerror("錯誤", f"建立採購單失敗: {str(e)}")



    def remove_from_pur_cart(self):
        """ 移除進貨購物車中的選定單項商品 """
        sel = self.tree_pur_cart.selection()
        if not sel:
            messagebox.showwarning("提示", "請先點選要移除的商品項目")
            return
        
        # 因為一次可能選多筆，我們倒著刪除，防止索引跑掉
        for item in sel:
            # 1. 取得該項目在 Treeview 裡的索引
            idx = self.tree_pur_cart.index(item)
            
            # 2. 從記憶體資料清單中移除
            if 0 <= idx < len(self.pur_cart_data):
                del self.pur_cart_data[idx]
            
            # 3. 從介面列表中移除
            self.tree_pur_cart.delete(item)
        
        # 4. 重新計算並更新介面上的總額顯示
        total_sum = sum(item['total'] for item in self.pur_cart_data)
        # 如果您有 self.lbl_pur_total，請更新它
        if hasattr(self, 'lbl_pur_total'):
            self.lbl_pur_total.config(text=f"本次進貨總額: ${total_sum:,.0f}")
            
        print("system: removed item from temporary list")

    @thread_safe_file
    def load_purchase_tracking(self):
        """ 
        載入追蹤清單：
        1. 強力防止 NaN 導致的轉型崩潰
        2. 自動清理物流單號的 .0 
        3. 統一數值顯示格式 (預防未來格式錯誤)
        """
        for i in self.tree_pur_track.get_children(): 
            self.tree_pur_track.delete(i)
            
        try:
            if not os.path.exists(FILE_NAME): return
            df = pd.read_excel(FILE_NAME, sheet_name=SHEET_PUR_TRACKING)
            
            if df.empty: return

            # --- [預防性數據清洗] ---
            # 將常見的空值表示符統一轉換為 Pandas 可辨識的空值，再統一填補
            df = df.replace(['nan', 'NaN', 'None', 'None', 'null'], pd.NA)

            for idx, row in df.iterrows():
                
                # A. 數量安全處理 (解決 NaN to Integer 報錯的關鍵)
                # 使用 pd.to_numeric 強制轉換，不成功的會變成 NaN，再用 fillna(0) 補 0
                raw_qty = row.get('數量', 0)
                safe_qty = int(pd.to_numeric(raw_qty, errors='coerce')) if pd.notna(pd.to_numeric(raw_qty, errors='coerce')) else 0

                # B. 財務數值安全處理 (單價、稅額、運費)
                def to_f_clean(val):
                    n = pd.to_numeric(val, errors='coerce')
                    return f"{float(n):.1f}" if pd.notna(n) else "0.0"

                # C. 物流單號安全處理 (徹底消滅 .0 與 nan)
                track_no = str(row.get('物流追蹤', '')).strip()
                if track_no.endswith('.0'):
                    track_no = track_no[:-2]
                if track_no.lower() in ['nan', 'none', '', 'nat']:
                    track_no = ""

                # D. 物流狀態處理
                status = str(row.get('物流狀態', '')).strip()
                if status.lower() in ['nan', 'none', '']:
                    status = "待發貨"

                # E. 進貨單號處理 (移除單引號顯示)
                pur_id = str(row.get('進貨單號', '')).replace("'", "").strip()

                # 寫入介面
                self.tree_pur_track.insert("", "end", text=str(idx), values=(
                    pur_id,
                    row.get('供應商', ''),
                    row.get('商品名稱', ''),
                    safe_qty,                  # 已修正的數量
                    to_f_clean(row.get('進貨單價', 0)),
                    to_f_clean(row.get('海關稅金', 0)),
                    to_f_clean(row.get('分攤運費', 0)),
                    status,                    # 已修正的狀態
                    track_no                   # 已修正的單號
                ))
        except Exception as e:
            import traceback
            traceback.print_exc()
            print(f"system: failed to load purchase tracking: {e}")

    def setup_pur_tracking_tab(self):
        """ 建立在途貨物追蹤：將狀態與單號拆分為獨立欄位 """
        frame = self.tab_pur_tracking
        top_frame = ttk.Frame(frame, padding=5)
        top_frame.pack(fill="x")
        ttk.Label(top_frame, text="🚚 運輸中貨物管理 (狀態與單號已拆分)", foreground="blue").pack(side="left")
        ttk.Button(top_frame, text="🔄 刷新列表", command=self.load_purchase_tracking).pack(side="right")

        # --- 關鍵修正：將最後一欄拆分為「狀態」與「物流單號」 ---
        cols_pur_track = ("進貨單號", "供應商", "商品名稱", "數量", "單價", "稅額", "運費", "物流狀態", "物流單號")
        
        self.tree_pur_track = ttk.Treeview(frame, columns=cols_pur_track, show='headings', height=15, selectmode="extended")
        
        # 設定欄位標題與寬度
        widths = {
            "進貨單號": 120, "供應商": 100, "商品名稱": 180, "數量": 60, 
            "單價": 70, "稅額": 60, "運費": 60, "物流狀態": 100, "物流單號": 150
        }
        for c in cols_pur_track:
            self.tree_pur_track.heading(c, text=c)
            self.tree_pur_track.column(c, width=widths.get(c, 100), anchor="center" if c != "商品名稱" else "w")
            
        self.tree_pur_track.pack(fill="both", expand=True, padx=10)

        # 按鈕區
        btn_ctrl = ttk.Frame(frame, padding=10)
        btn_ctrl.pack(fill="x")
        ttk.Button(btn_ctrl, text="✏️ 補充修改單項資訊", command=self.action_update_pur_logistics).pack(side="left", padx=5)
        ttk.Button(btn_ctrl, text="⚖️ 整單運費稅金自動分攤", command=self.action_batch_distribute_shipping).pack(side="left", padx=5)
        ttk.Button(btn_ctrl, text="✅ 確認收貨入庫", command=self.action_confirm_inbound).pack(side="left", padx=5)
        ttk.Button(btn_ctrl, text="❌ 標記遺失/取消進貨", command=self.action_cancel_purchase).pack(side="left", padx=5)

        
        self.load_purchase_tracking()


    @thread_safe_file
    def action_batch_distribute_shipping(self):
        """ 彈出視窗：輸入整筆單據的總運費/稅金，並依重量權重自動分攤 """
        sel = self.tree_pur_track.selection()
        if not sel:
            messagebox.showwarning("提示", "請先選擇該訂單中的任意一個商品項目")
            return
            
        item = self.tree_pur_track.item(sel[0])
        target_pur_id = str(item['values'][0]).replace("'", "").strip()

        win = tk.Toplevel(self.root)
        win.title(f"整單運費分攤 - {target_pur_id}")
        win.geometry("400x300")
        win.grab_set()

        ttk.Label(win, text=f"正在處理單號: {target_pur_id}", foreground="blue").pack(pady=10)
        
        body = ttk.Frame(win, padding=20)
        body.pack(fill="both")

        ttk.Label(body, text="這箱貨物的「總運費」($):").pack(anchor="w")
        var_total_ship = tk.DoubleVar(value=0.0)
        ttk.Entry(body, textvariable=var_total_ship).pack(fill="x", pady=5)

        ttk.Label(body, text="這箱貨物的「總稅金」($):").pack(anchor="w", pady=(10,0))
        var_total_tax = tk.DoubleVar(value=0.0)
        ttk.Entry(body, textvariable=var_total_tax).pack(fill="x", pady=5)


        
        def calculate_and_save():
            try:
                total_s = var_total_ship.get()
                total_t = var_total_tax.get()
                
                with pd.ExcelFile(FILE_NAME) as xls:
                    df_track = pd.read_excel(xls, sheet_name=SHEET_PUR_TRACKING)
                    df_hist = pd.read_excel(xls, sheet_name=SHEET_PURCHASES)
                    df_prods = pd.read_excel(xls, sheet_name=SHEET_PRODUCTS)

                # --- [核心修正 1：強制轉型為浮點數] ---
                # 這樣就不會再出現 "Invalid value for dtype int64" 的錯誤
                for df in [df_track, df_hist]:
                    for col in ['分攤運費', '海關稅金']:
                        if col not in df.columns: df[col] = 0.0
                        # 關鍵：先轉成 numeric (處理 NaN)，再強制轉成 float 類型
                        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0).astype(float)
                
                # 商品資料的權重也要確保是數字
                if '單位權重' not in df_prods.columns: df_prods['單位權重'] = 1.0
                df_prods['單位權重'] = pd.to_numeric(df_prods['單位權重'], errors='coerce').fillna(1.0).astype(float)

                # 建立重量地圖
                weight_map = df_prods.set_index('商品名稱')['單位權重'].to_dict()

                # 清理匹配用的 ID
                df_track['tmp_id'] = df_track['進貨單號'].astype(str).str.replace("'", "").str.strip()
                mask = df_track['tmp_id'] == target_pur_id
                
                # 1. 計算總權重
                order_total_weight = 0
                for _, row in df_track[mask].iterrows():
                    p_name = str(row['商品名稱']).strip()
                    qty = float(row.get('數量', 1))
                    u_weight = weight_map.get(p_name, 1.0)
                    order_total_weight += (qty * u_weight)

                if order_total_weight == 0: order_total_weight = 1

                # 2. 開始分配 (小數點保留至第一位)
                for idx in df_track[mask].index:
                    p_name = str(df_track.at[idx, '商品名稱']).strip()
                    qty = float(df_track.at[idx, '數量'])
                    u_weight = weight_map.get(p_name, 1.0)
                    
                    share_ratio = (qty * u_weight) / order_total_weight
                    
                    # --- [核心修正 2：保留小數點第一位] ---
                    alloc_s = round(total_s * share_ratio, 1)
                    alloc_t = round(total_t * share_ratio, 1)

                    df_track.at[idx, '分攤運費'] = alloc_s
                    df_track.at[idx, '海關稅金'] = alloc_t
                    
                    # 同步歷史總表
                    df_hist['tmp_id'] = df_hist['進貨單號'].astype(str).str.replace("'", "").str.strip()
                    h_mask = (df_hist['tmp_id'] == target_pur_id) & (df_hist['商品名稱'].astype(str).str.strip() == p_name)
                    
                    # 確保歷史總表的欄位也接受浮點數
                    df_hist.loc[h_mask, '分攤運費'] = alloc_s
                    df_hist.loc[h_mask, '海關稅金'] = alloc_t

                # 存檔前清理
                df_track.drop(columns=['tmp_id'], inplace=True, errors='ignore')
                df_hist.drop(columns=['tmp_id'], inplace=True, errors='ignore')
                
                if self._universal_save({SHEET_PUR_TRACKING: df_track, SHEET_PURCHASES: df_hist}):
                    messagebox.showinfo("成功", f"分攤完畢！\n已將運費與稅金分配至各品項 (保留1位小數)")
                    self.load_purchase_tracking()
                    win.destroy()

            except Exception as e:
                import traceback
                traceback.print_exc()
                messagebox.showerror("計算失敗", f"發生型別錯誤: {str(e)}\n請檢查 Excel 欄位格式。")

        ttk.Button(win, text="🚀 開始自動分攤並儲存", command=calculate_and_save).pack(pady=20)



    def setup_vendor_tab(self):
        """ 建立進貨廠商管理介面 (UI 優化排版版) """
        paned = ttk.PanedWindow(self.tab_vendors, orient=tk.HORIZONTAL)
        paned.pack(fill="both", expand=True, padx=10, pady=10)

        # --- 左側：新增/編輯區 ---
        left_f = ttk.LabelFrame(paned, text="🆕 廠商資料維護", padding=15)
        paned.add(left_f, weight=1)

        # 設定 Grid 權重，讓輸入框自動拉長
        left_f.columnconfigure(1, weight=1)

        grid_opts = {'sticky': 'w', 'pady': 8} # 增加垂直間距
        e_opts = {'sticky': 'ew', 'pady': 8, 'padx': (10, 0)}
        
        curr = 0
        ttk.Label(left_f, text="* 廠商名稱:").grid(row=curr, column=0, **grid_opts)
        ttk.Entry(left_f, textvariable=self.var_v_name).grid(row=curr, column=1, **e_opts)
        curr += 1

        ttk.Label(left_f, text="採購通路:").grid(row=curr, column=0, **grid_opts)
        ttk.Entry(left_f, textvariable=self.var_v_channel).grid(row=curr, column=1, **e_opts)
        curr += 1

        ttk.Label(left_f, text="統一編號:").grid(row=curr, column=0, **grid_opts)
        ttk.Entry(left_f, textvariable=self.var_v_taxid).grid(row=curr, column=1, **e_opts)
        curr += 1

        ttk.Label(left_f, text="聯絡人:").grid(row=curr, column=0, **grid_opts)
        ttk.Entry(left_f, textvariable=self.var_v_contact).grid(row=curr, column=1, **e_opts)
        curr += 1

        ttk.Label(left_f, text="聯絡電話:").grid(row=curr, column=0, **grid_opts)
        ttk.Entry(left_f, textvariable=self.var_v_phone).grid(row=curr, column=1, **e_opts)
        curr += 1

        ttk.Label(left_f, text="廠商地址:").grid(row=curr, column=0, **grid_opts)
        ttk.Entry(left_f, textvariable=self.var_v_addr).grid(row=curr, column=1, **e_opts)
        curr += 1

        ttk.Label(left_f, text="備註事項:").grid(row=curr, column=0, **grid_opts)
        ttk.Entry(left_f, textvariable=self.var_v_remarks).grid(row=curr, column=1, **e_opts)
        curr += 1

        # --- 將績效相關元件封裝在一個專屬 Frame 中 ---
        self.perf_frame = ttk.Frame(left_f)
        
        # 內部的績效評分 Label
        ttk.Label(self.perf_frame, text="系統績效評分:").grid(row=0, column=0, sticky="w", pady=8)
        ttk.Label(self.perf_frame, textvariable=self.var_v_system_score, font=("", 10, "bold"), foreground="blue").grid(row=0, column=1, sticky="w", padx=10)
        
        # 內部的星等 Combo
        ttk.Label(left_f, text="主觀印象星等:").grid(row=curr, column=0, **grid_opts)
        
        # 建立一個容器來水平放置 下拉選單、提示、與按鈕
        action_row_f = ttk.Frame(left_f)
        action_row_f.grid(row=curr, column=1, sticky="ew", padx=10)
        
        # 1. 下拉選單
        self.combo_v_manual = ttk.Combobox(action_row_f, textvariable=self.var_v_manual_adj, values=["5","4","3","2","1"], width=3, state="readonly")
        self.combo_v_manual.pack(side="left")
        self.combo_v_manual.bind("<<ComboboxSelected>>", lambda e: self.refresh_vendor_live_score(self.var_v_name.get()))
        ttk.Label(action_row_f, text="星", foreground="gray").pack(side="left", padx=(2, 10))

        #評分提示小字
        ttk.Label(action_row_f, text="(針對溝通、包裝打分)", foreground="gray", font=("微軟正黑體", 9)).pack(side="left", padx=10)

        # 2. [儲存] 按鈕 (緊跟在星等後)
        ttk.Button(action_row_f, text="💾 儲存廠商", command=self.submit_vendor, width=12).pack(side="left", padx=5)

        # 3. [刪除] 按鈕
        ttk.Button(action_row_f, text="🗑️ 刪除", command=self.delete_vendor, width=8).pack(side="left", padx=5)

       
        
        curr += 1

        self.perf_frame.grid(row=curr, column=0, columnspan=2, sticky="ew")
        self.perf_frame.columnconfigure(1, weight=1) # 讓內容對齊

        curr += 1


        # --- 進入批次處理區塊 ---
        # 加入橫線分隔
        ttk.Separator(left_f, orient="horizontal").grid(row=curr, column=0, columnspan=2, sticky="ew", pady=15)
        curr += 1

        # 標題
        ttk.Label(left_f, text="📂 外部資料批次處理", font=("微軟正黑體", 10, "bold")).grid(row=curr, column=0, columnspan=2, sticky="w", padx=5)
        curr += 1

        # 啟動精靈按鈕
        btn_v_import = ttk.Button(left_f, text="📥 啟動廠商批次匯入精靈", command=self.open_vendor_import_wizard)
        btn_v_import.grid(row=curr, column=0, columnspan=2, sticky="ew", padx=10, pady=(5, 0))
        curr += 1

        # 底部提示小字
        ttk.Label(left_f, text="* 支援舊檔 Excel 欄位匹配匯入", foreground="gray", font=("微軟正黑體", 9)).grid(row=curr, column=0, columnspan=2, sticky="w", padx=5)

        # --- 右側清單 ---
        right_f = ttk.LabelFrame(paned, text="🔍 廠商清單", padding=15)
        paned.add(right_f, weight=1)

        ent_search = ttk.Entry(right_f, textvariable=self.var_v_search)
        ent_search.pack(fill="x", pady=(0, 10))
        ent_search.bind('<KeyRelease>', lambda e: self.update_vendor_list())

        self.list_vendors = tk.Listbox(right_f, font=("微軟正黑體", int(self.var_font_size.get())), relief="flat", borderwidth=1)
        self.list_vendors.pack(fill="both", expand=True)
        
        # 加上滾動條讓清單專業點
        sc_v = ttk.Scrollbar(self.list_vendors, orient="vertical", command=self.list_vendors.yview)
        self.list_vendors.configure(yscrollcommand=sc_v.set)
        sc_v.pack(side="right", fill="y")
        
        self.list_vendors.bind('<<ListboxSelect>>', self.on_vendor_select)

        self.update_vendor_list()


    def refresh_vendor_management_ui(self):
        """ 根據開關，動態隱藏或顯示廠商管理的績效區塊 """
        if not hasattr(self, 'perf_frame'): return
        
        if self.var_enable_vendor_kpi.get():
            # 重新顯示
            self.perf_frame.grid()
        else:
            # 隱藏 (grid_remove 會保留位置但看不到，grid_forget 會徹底移除空間)
            self.perf_frame.grid_remove()


    @thread_safe_file
    def update_vendor_list(self):
        """ 刷新廠商清單 """
        self.list_vendors.delete(0, tk.END)
        query = self.var_v_search.get().lower().strip()
        try:
            df = pd.read_excel(FILE_NAME, sheet_name=SHEET_VENDORS)
            for _, row in df.iterrows():
                name = str(row['廠商名稱'])
                channel = str(row.get('通路', ''))
                if query in name.lower() or query in channel.lower():
                    self.list_vendors.insert(tk.END, f"{name} ({channel})")
        except: pass

    @thread_safe_file
    def on_vendor_select(self, event):
        """ 當點選廠商清單時，將詳情填入左側，並即時運算績效評分 """
        sel = self.list_vendors.curselection()
        if not sel: return
        
        display_str = self.list_vendors.get(sel[0])
        v_name_selected = display_str.split(" (")[0].strip()

        try:
            # 1. 讀取廠商基本資料 (地址、電話等)
            df_v = pd.read_excel(FILE_NAME, sheet_name=SHEET_VENDORS)
            df_v.columns = [str(c).strip() for c in df_v.columns]
            
            row = df_v[df_v['廠商名稱'].astype(str).str.strip() == v_name_selected].iloc[0]
            
            def c(val, default=""):
                if pd.isna(val) or str(val).lower() == "nan": return default
                return str(val)

            self.var_v_name.set(c(row['廠商名稱']))
            self.var_v_channel.set(c(row.get('通路', '')))
            self.var_v_taxid.set(c(row.get('統編', '')))
            self.var_v_contact.set(c(row.get('聯絡人', '')))
            self.var_v_phone.set(c(row.get('電話', '')))
            self.var_v_addr.set(c(row.get('地址', '')))
            self.var_v_remarks.set(c(row.get('備註', '')))
            self.var_v_manual_adj.set(str(int(pd.to_numeric(row.get('星等', 5), errors='coerce'))))

            # 2. 【核心新增】即時運算進貨紀錄並顯示結果
            self.refresh_vendor_live_score(v_name_selected)

        except Exception as e:
            print(f"system: failed to load vendor details: {e}")


    @thread_safe_file
    def refresh_vendor_live_score(self, vendor_name):
        """ 
        V5.2 修正版：對接新版物流節點欄位 (解決 0d 問題)
        """
        if not vendor_name or vendor_name == "" or vendor_name == "nan":
            self.var_v_system_score.set("請選擇廠商")
            return

        try:
            df_h = pd.read_excel(FILE_NAME, sheet_name=SHEET_PURCHASES)
            v_mask = (df_h['供應商'].astype(str).str.strip() == vendor_name)
            v_all_data = df_h[v_mask].copy()
            
            # 只抓有『入庫日期』的結案單據
            v_data = v_all_data[v_all_data['入庫日期'].notna() & (v_all_data['入庫日期'] != "")].copy()

            if v_data.empty:
                pending_count = len(v_all_data)
                self.var_v_system_score.set(f"評估中 (尚無結案紀錄，目前有 {pending_count} 筆在途)")
                self.var_v_leadtime.set("等待首次到貨")
                return

            # --- [核心參數導入] ---
            w_q = self.var_w_quality.get()
            w_p = self.var_w_prep.get()
            w_f = self.var_w_fulfill.get()
            w_t = self.var_w_transit.get()
            std_p = self.var_std_prep.get()
            std_t = self.var_std_transit.get()
            sys_ratio = self.var_w_system_ratio.get()

            # --- [數據清洗：對接新版欄位] ---
            # 採購日
            pur_dt = pd.to_datetime(v_data['採購日期'], errors='coerce')
            
            # 出貨日：優先抓『時間_廠商出貨』，如果沒填才抓舊的『賣家交付日期』
            if '時間_廠商出貨' in v_data.columns:
                ship_dt = pd.to_datetime(v_data['時間_廠商出貨'], errors='coerce')
            else:
                ship_dt = pd.to_datetime(v_data.get('賣家交付日期'), errors='coerce')
            
            # 入庫日
            in_dt = pd.to_datetime(v_data['入庫日期'], errors='coerce')

            # --- [指標計算] ---
            # 1. 備貨天數 (出貨 - 採購)
            prep_series = (ship_dt - pur_dt).dt.days.dropna()
            avg_prep = round(prep_series.mean(), 1) if not prep_series.empty else 0

            # 2. 運輸天數 (入庫 - 出貨)
            transit_series = (in_dt - ship_dt).dt.days.dropna()
            avg_transit = round(transit_series.mean(), 1) if not transit_series.empty else 0

            # 3. 品質合格率 (實收數量 vs 瑕疵數量)
            qty_s = pd.to_numeric(v_data['數量'], errors='coerce').fillna(0)
            defect_s = pd.to_numeric(v_data.get('瑕疵數量', 0), errors='coerce').fillna(0)
            q_rate = round((1 - (defect_s.sum() / qty_s.sum())) * 100, 1) if qty_s.sum() > 0 else 0

            # 4. 到貨滿足率 (實收數量 vs 原始預計)
            orig_s = pd.to_numeric(v_data.get('原始預計數量', qty_s), errors='coerce').fillna(0)
            f_rate = round((qty_s.sum() / orig_s.sum()) * 100, 1) if orig_s.sum() > 0 else 0

            # --- [評分引擎] ---
            score_prep = max(100 - (max(avg_prep - std_p, 0) * 10), 0)
            score_transit = max(100 - (max(avg_transit - std_t, 0) * 10), 0)
            
            system_score = (q_rate * w_q) + (score_prep * w_p) + (f_rate * w_f) + (score_transit * w_t)

            # 混合星等
            try:
                manual_stars = int(self.var_v_manual_adj.get())
                manual_score = manual_stars * 20
            except:
                manual_score = 100
                
            final_mixed_score = (system_score * sys_ratio) + (manual_score * (1 - sys_ratio))

            # --- [介面呈現] ---
            display_text = f"{round(final_mixed_score, 1)} (質:{int(q_rate)}% / 備:{avg_prep}d / 運:{avg_transit}d)"
            self.var_v_system_score.set(display_text)
            
            total_days_series = (in_dt - pur_dt).dt.days.dropna()
            avg_total = round(total_days_series.mean(), 1) if not total_days_series.empty else 0
            self.var_v_leadtime.set(f"平均總耗時: {avg_total} 天")

        except Exception as e:
            print(f"評分更新失敗: {e}")
            self.var_v_system_score.set("計算中...")


    @thread_safe_file
    def initialize_kpi_defaults(self):
        """ 確保系統設定分頁具備所有必要的 KPI 參數 """
        try:
            df_sys = pd.read_excel(FILE_NAME, sheet_name=SHEET_SYS_SETTINGS)
            
            # 定義預設清單 (Key: Value)
            defaults = {
                "VENDOR_ENABLE_KPI": "True", # 加入這行以啟用 KPI 功能
                "VENDOR_W_QUALITY": "0.4",
                "VENDOR_W_PREP": "0.3",
                "VENDOR_W_FULFILL": "0.2",
                "VENDOR_W_TRANSIT": "0.1",
                "VENDOR_STD_PREP": "3",
                "VENDOR_STD_TRANSIT": "5",
                "VENDOR_W_SYSTEM_RATIO": "0.8"
            }
            
            changed = False
            for key, val in defaults.items():
                if key not in df_sys['設定名稱'].values:
                    new_row = pd.DataFrame([{"設定名稱": key, "參數值": val}])
                    df_sys = pd.concat([df_sys, new_row], ignore_index=True)
                    changed = True
            
            if changed:
                self._universal_save({SHEET_SYS_SETTINGS: df_sys})
                print("system: KPI default parameters initialized.")
        except Exception as e:
            print(f"system: failed to initialize KPI parameters: {e}")


    @thread_safe_file
    def submit_vendor(self):
        """ 儲存或更新廠商資料 (修正版：修復提取函數與通路存檔問題) """
        name = self.var_v_name.get().strip()
        channel = self.var_v_channel.get().strip() # 抓取通路

        if not name:
            messagebox.showwarning("警告", "「廠商名稱」為必填項目！")
            return

        try:
            # 1. 讀取資料
            try:
                df = pd.read_excel(FILE_NAME, sheet_name=SHEET_VENDORS)
                df = df.fillna("") # 讀取後立刻清掉 nan
            except:
                # 若讀取失敗，建立符合結構的空表
                df = pd.DataFrame(columns=[
                    "廠商名稱", "通路", "統編", "聯絡人", "電話", "地址", "備註", 
                    "平均前置天數", "總到貨率", "總合格率", "綜合評等分數", "星等", "最後更新"
                ])

            # 強制清理標題空白
            df.columns = [str(c).strip() for c in df.columns]

            # 2. 定義輔助函數 (放在內部確保能讀取變數)
            def get_extracted_value(source_str, key):
                """ 提取 '92.5 (質:100% / 備:0.5d / 運:3.0d)' 中的數值 """
                try:
                    if key in source_str:
                        # 找到 key 之後，取到下一個斜線、天(d)或百分比(%)為止
                        part = source_str.split(key)[1].split('/')[0].split(')')[0].strip()
                        return part
                    return "0%" if "質" in key or "滿" in key else "0"
                except:
                    return "0"

            now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
            score_raw = self.var_v_system_score.get()
            
            # 3. 建立要寫入的資料包
            new_entry = {
                "廠商名稱": name,
                "通路": channel if channel else "",  
                "統編": self.var_v_taxid.get().strip(),
                "聯絡人": self.var_v_contact.get().strip(),
                "電話": self.var_v_phone.get().strip(),
                "地址": self.var_v_addr.get().strip(),
                "備註": self.var_v_remarks.get().strip(),
                # 這裡對應您畫面上顯示的標籤 (質、備、運)
                "平均前置天數": get_extracted_value(score_raw, "備:").replace("d", ""), 
                "總到貨率": get_extracted_value(score_raw, "滿:"), # 備註：若無「滿:」則會傳回 0%
                "總合格率": get_extracted_value(score_raw, "質:"),
                "綜合評等分數": score_raw.split('(')[0].strip() if '(' in score_raw else "0.0",
                "最後更新": now_str
            }
            
            try: 
                new_entry["星等"] = int(self.var_v_manual_adj.get())
            except: 
                new_entry["星等"] = 5

            # 4. 準備型別轉換 (防止 float64 衝突)
            for col in df.columns:
                df[col] = df[col].astype(object)

            # 5. 執行更新或新增
            df['廠商名稱_clean'] = df['廠商名稱'].astype(str).str.strip()
            
            if name in df['廠商名稱_clean'].values:
                idx = df[df['廠商名稱_clean'] == name].index[0]
                for key, val in new_entry.items():
                    if key in df.columns: # 確保欄位存在才寫入
                        df.at[idx, key] = val
            else:
                df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)

            df = df.dropna(subset=['廠商名稱'])
            df = df[df['廠商名稱'] != ""]

            # 移除臨時欄位
            if '廠商名稱_clean' in df.columns:
                df = df.drop(columns=['廠商名稱_clean'])

            # 6. 寫回前強制校準數值欄位
            numeric_cols = ["平均前置天數", "綜合評等分數", "星等"]
            for col in numeric_cols:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

            # 7. 呼叫萬用引擎存檔
            if self._universal_save({SHEET_VENDORS: df}):
                messagebox.showinfo("成功", f"廠商 [{name}] 資料與績效評分已更新。")
                self.update_vendor_list()
                self.update_pur_supplier_list()
                
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("儲存失敗", f"錯誤詳情：{str(e)}")

    @thread_safe_file
    def delete_vendor(self):
        name = self.var_v_name.get().strip()
        if not name or not messagebox.askyesno("確認", f"確定刪除廠商 [{name}]？"): return
        try:
            df = pd.read_excel(FILE_NAME, sheet_name=SHEET_VENDORS)
            df = df[df['廠商名稱'] != name]
            if self._universal_save({SHEET_VENDORS: df}):
                self.update_vendor_list()
                self.update_pur_supplier_list()
                self.var_v_name.set(""); self.var_v_channel.set("")
        except: pass


    def open_vendor_import_wizard(self):
        """ 開啟廠商匯入精靈 """
        from VendorImportWizard import VendorImportWizard # 假設檔案名
        VendorImportWizard(self.root, self.callback_vendor_import)


    @thread_safe_file
    def callback_vendor_import(self, new_data_list):
        """ 處理匯入後的廠商資料合併 """
        try:
            df_new = pd.DataFrame(new_data_list)
            
            # 讀取現有廠商
            with pd.ExcelFile(FILE_NAME) as xls:
                df_old = pd.read_excel(xls, sheet_name=SHEET_VENDORS)

            # 合併，以「廠商名稱」為準去重
            df_combined = pd.concat([df_old, df_new], ignore_index=True)
            df_combined.drop_duplicates(subset=['廠商名稱'], keep='last', inplace=True)
            
            # 存檔
            if self._universal_save({SHEET_VENDORS: df_combined}):
                self.update_vendor_list()
                return True
            return False
        except Exception as e:
            messagebox.showerror("匯入失敗", f"錯誤: {e}")
            return False




       # ================= 營收與商品分析 (新功能) =================
    def setup_analysis_tab(self):
        # 主框架：左右分割
        paned = ttk.PanedWindow(self.tab_analysis, orient=tk.HORIZONTAL)
        paned.pack(fill="both", expand=True, padx=10, pady=10)

        # --- 左側：時間維度收益分析 ---
        left_frame = ttk.LabelFrame(paned, text="📅 週期收益報表 (月/週/日)", padding=10)
        paned.add(left_frame, weight=1)

        # 1. 摘要看板 (Summary)
        summary_frame = ttk.Frame(left_frame, relief="groove", borderwidth=2)
        summary_frame.pack(fill="x", pady=(0, 10))
        
        self.lbl_month_sales = ttk.Label(summary_frame, text="本月營收: $0", font=("微軟正黑體", 12, "bold"), foreground="blue")
        self.lbl_month_sales.pack(anchor="w", padx=5, pady=2)
        self.lbl_month_profit = ttk.Label(summary_frame, text="本月淨利: $0", font=("微軟正黑體", 12, "bold"), foreground="green")
        self.lbl_month_profit.pack(anchor="w", padx=5, pady=2)

        # 2. 詳細列表 (Treeview)
        cols_time = ("時間區間", "總營收", "總淨利", "訂單數")
        self.tree_time_stats = ttk.Treeview(left_frame, columns=cols_time, show='headings', height=15)
        
        self.tree_time_stats.heading("時間區間", text="時間區間 (月/日)")
        self.tree_time_stats.column("時間區間", width=120)
        self.tree_time_stats.heading("總營收", text="總營收")
        self.tree_time_stats.column("總營收", width=80, anchor="e")
        self.tree_time_stats.heading("總淨利", text="總淨利")
        self.tree_time_stats.column("總淨利", width=80, anchor="e")
        self.tree_time_stats.heading("訂單數", text="訂單數")
        self.tree_time_stats.column("訂單數", width=50, anchor="center")
        
        self.tree_time_stats.pack(fill="both", expand=True)

# --- 右側：商品銷售排行榜 ---
        right_frame = ttk.LabelFrame(paned, text="🏆 商品銷售排行榜", padding=10)
        paned.add(right_frame, weight=1)

        # 排序控制區
        sort_frame = ttk.Frame(right_frame)
        sort_frame.pack(fill="x", pady=(0, 5))
        ttk.Label(sort_frame, text="排序依據:").pack(side="left")
        
        self.var_prod_sort_by = tk.StringVar(value="平均毛利率")
        sort_options = ["平均毛利率", "總銷量排行", "總獲利排行", "銷售速度排行"]
        self.combo_prod_sort = ttk.Combobox(sort_frame, textvariable=self.var_prod_sort_by, values=sort_options, state="readonly", width=12)
        self.combo_prod_sort.pack(side="left", padx=5)
        self.combo_prod_sort.bind("<<ComboboxSelected>>", lambda e: self.calculate_analysis_data())

        cols_prod_ids = ("p_name", "p_margin", "p_profit", "p_qty", "p_velocity")

        self.tree_prod_stats = ttk.Treeview(right_frame, columns=cols_prod_ids, show='headings', height=15)
        
        # 設定各欄位
        self.tree_prod_stats.heading("p_name", text="商品名稱")
        self.tree_prod_stats.column("p_name", width=150)
        self.tree_prod_stats.heading("p_margin", text="平均毛利", command=lambda: self.sort_tree_column(self.tree_prod_stats, "p_margin", False))
        self.tree_prod_stats.column("p_margin", width=80, anchor="e")
        self.tree_prod_stats.heading("p_profit", text="總獲利", command=lambda: self.sort_tree_column(self.tree_prod_stats, "p_profit", False))
        self.tree_prod_stats.column("p_profit", width=80, anchor="e")
        self.tree_prod_stats.heading("p_qty", text="總銷量", command=lambda: self.sort_tree_column(self.tree_prod_stats, "p_qty", False))
        self.tree_prod_stats.column("p_qty", width=60, anchor="center")
        self.tree_prod_stats.heading("p_velocity", text="銷售速度", command=lambda: self.sort_tree_column(self.tree_prod_stats, "p_velocity", False))
        self.tree_prod_stats.column("p_velocity", width=100, anchor="e")

        sb = ttk.Scrollbar(right_frame, orient="vertical", command=self.tree_prod_stats.yview)
        self.tree_prod_stats.configure(yscrollcommand=sb.set)
        self.tree_prod_stats.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        btn_refresh = ttk.Button(self.tab_analysis, text="🔄 重新計算分析數據", command=self.calculate_analysis_data)
        btn_refresh.pack(fill="x", pady=10, padx=10)
        
        self.calculate_analysis_data()


    @thread_safe_file
    def calculate_analysis_data(self):
        """ 核心分析邏輯 V5.2: 使用 Decimal 高精度運算，修正消失問題 """
        if not hasattr(self, 'tree_time_stats') or not hasattr(self, 'tree_prod_stats'): return
        
        for i in self.tree_time_stats.get_children(): self.tree_time_stats.delete(i)
        for i in self.tree_prod_stats.get_children(): self.tree_prod_stats.delete(i)
        
        if not os.path.exists(FILE_NAME): return

        try:
            with pd.ExcelFile(FILE_NAME) as xls:
                df_sales = pd.read_excel(xls, sheet_name=SHEET_SALES)
                df_prods = pd.read_excel(xls, sheet_name=SHEET_PRODUCTS)

            if df_sales.empty: return

            # --- [數據清洗] ---
            df_sales = df_sales.replace(r'^\s*$', pd.NA, regex=True)
            fill_cols = ['訂單編號', '日期', '買家名稱', '交易平台']
            for col in fill_cols:
                if col in df_sales.columns:
                    df_sales[col] = df_sales[col].ffill()

            df_sales = df_sales.dropna(subset=['商品名稱'])
            df_sales['日期'] = pd.to_datetime(df_sales['日期'], errors='coerce')
            
            # --- [左側：時間維度統計] ---
            df_time = df_sales.dropna(subset=['日期']).copy()

            if not df_time.empty:
                df_time['月份'] = df_time['日期'].dt.strftime('%Y-%m')
                
                # 按月份分組並使用 Decimal 計算
                monthly_list = []
                for month, group in df_time.groupby('月份', sort=False):
                    m_sales = Decimal("0.00")
                    m_profit = Decimal("0.00")
                    for _, r in group.iterrows():
                        m_sales += Decimal(str(r.get('總銷售額', 0)))
                        m_profit += Decimal(str(r.get('總淨利', 0)))
                    
                    monthly_list.append({
                        'month': month,
                        'sales': m_sales,
                        'profit': m_profit,
                        'count': group['訂單編號'].nunique()
                    })
                
                # 排序 (最新月份在前)
                monthly_list.sort(key=lambda x: x['month'], reverse=True)

                # 更新頂部看板 (本月)
                if monthly_list:
                    curr = monthly_list[0]
                    self.lbl_month_sales.config(text=f"本月({curr['month']}) 營收: ${float(curr['sales']):,.2f}")
                    self.lbl_month_profit.config(text=f"本月({curr['month']}) 淨利: ${float(curr['profit']):,.2f}")

                # 填入月份表格
                for m in monthly_list:
                    self.tree_time_stats.insert("", "end", values=(
                        f"{m['month']} (月)", 
                        f"${float(m['sales']):,.2f}", 
                        f"${float(m['profit']):,.2f}", 
                        f"{int(m['count'])} 單"
                    ))

                self.tree_time_stats.insert("", "end", values=("--- 近10日明細 ---", "", "", ""))

                # 每日明細計算 (Decimal)
                df_time['日期字串'] = df_time['日期'].dt.strftime('%Y-%m-%d')
                daily_groups = df_time.groupby('日期字串', sort=False)
                
                daily_list = []
                for d_str, group in daily_groups:
                    d_sales = Decimal("0.00")
                    d_profit = Decimal("0.00")
                    for _, r in group.iterrows():
                        d_sales += Decimal(str(r.get('總銷售額', 0)))
                        d_profit += Decimal(str(r.get('總淨利', 0)))
                    daily_list.append((d_str, d_sales, d_profit, group['訂單編號'].nunique()))

                daily_list.sort(key=lambda x: x[0], reverse=True)
                for d in daily_list[:10]:
                    self.tree_time_stats.insert("", "end", values=(
                        d[0], f"${float(d[1]):,.2f}", f"${float(d[2]):,.2f}", f"{int(d[3])} 單"
                    ))

           # --- [2. 右側：商品排行榜核心修正] ---
            # A. 預處理數值欄位，轉為 float 方便 Pandas 內部排序運算
            df_sales['S_F'] = df_sales['總銷售額'].astype(float)
            df_sales['P_F'] = df_sales['總淨利'].astype(float)
            df_sales['Q_F'] = df_sales['數量'].astype(float)

            # B. 執行聚合
            prod_group = df_sales.groupby('商品名稱').agg({
                'S_F': 'sum',
                'P_F': 'sum',
                'Q_F': 'sum'
            }).reset_index()

            # C. 計算毛利率欄位 (這就是排序失效的主因)
            prod_group['Margin_F'] = (prod_group['P_F'] / prod_group['S_F'] * 100).fillna(0)

            # D. 計算銷售速度
            now = pd.Timestamp.now()
            start_date_map = df_prods.set_index('商品名稱')['初始上架時間'].to_dict()
            first_sale_map = df_sales.groupby('商品名稱')['日期'].min().to_dict()

            def get_velocity(name, qty):
                st = pd.to_datetime(start_date_map.get(name), errors='coerce')
                if pd.isna(st): st = first_sale_map.get(name)
                if pd.isna(st): st = now
                days = max((now - st).days, 1)
                return round(float(qty) / days, 2)

            prod_group['velocity'] = prod_group.apply(lambda r: get_velocity(r['商品名稱'], r['Q_F']), axis=1)

            # E. 【關鍵：執行排序】
            sort_mode = self.var_prod_sort_by.get()
            sort_map = {
                "平均毛利率": 'Margin_F',
                "總銷量排行": 'Q_F',
                "總獲利排行": 'P_F',
                "銷售速度排行": 'velocity'
            }
            target_col = sort_map.get(sort_mode, 'P_F')
            prod_group = prod_group.sort_values(by=target_col, ascending=False)

            # F. 填入 Treeview (顯示時使用 Decimal 確保美觀，內部計算已完成)
            for _, row in prod_group.iterrows():
                # 轉回 Decimal 只是為了呼叫 float 轉字串時格式漂亮
                self.tree_prod_stats.insert("", "end", values=(
                    row['商品名稱'], 
                    f"{row['Margin_F']:.1f}%", 
                    f"${float(row['P_F']):,.2f}", 
                    int(row['Q_F']), 
                    f"{row['velocity']} 件/日"
                ))

        except Exception as e:
            import traceback
            traceback.print_exc()

            
    def sort_tree_column(self, tree, col, reverse):
        """(進階功能) 點擊標題可以排序"""
        l = [(tree.set(k, col), k) for k in tree.get_children('')]
        
        # 嘗試將字串轉數字進行排序 (去除 $ 和 % 符號)
        try:
            l.sort(key=lambda t: float(t[0].replace('$', '').replace(',', '').replace('%', '')), reverse=reverse)
        except ValueError:
            l.sort(reverse=reverse)

        # 重新排列
        for index, (val, k) in enumerate(l):
            tree.move(k, '', index)

        # 切換下次排序順序
        tree.heading(col, command=lambda: self.sort_tree_column(tree, col, not reverse))



    def setup_procurement_tab(self):
        """ 採購需求分析 V5.4: 修復顏色標籤與佈局優化 """
        frame = self.tab_procurement
        
        # 主容器：左右分割
        self.procure_paned = ttk.PanedWindow(frame, orient=tk.HORIZONTAL)
        self.procure_paned.pack(fill="both", expand=True, padx=5, pady=5)

        # ================= [左側：採購建議區] =================
        left_main_f = ttk.Frame(self.procure_paned)
        self.procure_paned.add(left_main_f, weight=1)

        # 參數控制區
        ctrl_frame = ttk.LabelFrame(left_main_f, text="⚙️ 採購評估參數", padding=10)
        ctrl_frame.pack(fill="x", pady=5)
        
        ttk.Label(ctrl_frame, text="銷售速度>").grid(row=0, column=0)
        ttk.Entry(ctrl_frame, textvariable=self.var_filter_velocity, width=5).grid(row=0, column=1)
        ttk.Label(ctrl_frame, text="備貨時間(天)").grid(row=0, column=2, padx=5)
        ttk.Entry(ctrl_frame, textvariable=self.var_days_to_cover, width=5).grid(row=0, column=3)
        ttk.Button(ctrl_frame, text="🔄 刷新", command=self.generate_procurement_report).grid(row=0, column=4, padx=10)

        # 建議清單
        list_frame = ttk.LabelFrame(left_main_f, text="📋 建議採購商品清單", padding=10)
        list_frame.pack(fill="both", expand=True, pady=5)
        
        cols = ("品名", "目前庫存", "安全值", "銷售速度", "缺貨狀態", "建議採購量")
        self.tree_procure = ttk.Treeview(list_frame, columns=cols, show='headings', height=20)
        widths = {"品名": 150, "目前庫存": 60, "安全值": 60, "銷售速度": 80, "缺貨狀態": 80, "建議採購量": 80}
        for c in cols:
            self.tree_procure.heading(c, text=c)
            self.tree_procure.column(c, width=widths[c], anchor="center")
        
        # --- [關鍵修正：解決主題覆蓋標籤顏色的問題] ---
        # 很多 Windows 主題會強制覆蓋 Treeview 的顏色，這段代碼能強制將自定義顏色刷上去
        def fixed_map(option):
            return [elm for elm in self.style.map('Treeview', query_opt=option) if elm[:2] != ('!disabled', '!selected')]
        self.style.map('Treeview', foreground=fixed_map('foreground'), background=fixed_map('background'))

        # 定義顏色標籤 (確保顏色在 white 背景下清晰)
        self.tree_procure.tag_configure('urgent', foreground='#CC0000', font=("微軟正黑體", 11))  # 深紅
        self.tree_procure.tag_configure('warning', foreground='#FF8C00', font=("微軟正黑體", 11)) # 橘色
        
        self.tree_procure.pack(fill="both", expand=True)

        # ================= [右側：定價估算器] =================
        right_main_f = ttk.LabelFrame(self.procure_paned, text="💰 商品上架價格估算器", padding=15)
        self.procure_paned.add(right_main_f, weight=1)

        # --- A. 搜尋區 ---
        search_f = ttk.Frame(right_main_f)
        search_f.pack(fill="x", pady=(0, 10))
        ttk.Label(search_f, text="🔍 搜尋商品:").pack(side="left")
        self.var_calc_search = tk.StringVar()
        ent_calc_search = ttk.Entry(search_f, textvariable=self.var_calc_search)
        ent_calc_search.pack(side="left", fill="x", expand=True, padx=5)
        ent_calc_search.bind('<KeyRelease>', self.filter_calc_prod_list)

        # --- B. 商品選取列表 ---
        self.list_calc_prod = tk.Listbox(right_main_f, height=10, font=("微軟正黑體", 10))
        self.list_calc_prod.pack(fill="x", pady=5)
        self.list_calc_prod.bind('<<ListboxSelect>>', self.on_calc_prod_select)

       # --- C. 計算面板 ---
        self.calc_grid_f = ttk.LabelFrame(right_main_f, text="🧮 定價試算參數", padding=15)
        self.calc_grid_f.pack(fill="both", expand=True, pady=10)

        c_opts = {'pady': 5, 'sticky': 'w'}
        ttk.Label(self.calc_grid_f, text="選中商品:").grid(row=0, column=0, **c_opts)
        ttk.Label(self.calc_grid_f, textvariable=self.var_calc_name, foreground="blue", font=("", 10, "bold")).grid(row=0, column=1, columnspan=3, **c_opts)

        ttk.Label(self.calc_grid_f, text="商品成本($):").grid(row=1, column=0, **c_opts)
        ttk.Entry(self.calc_grid_f, textvariable=self.var_calc_cost, width=15).grid(row=1, column=1, **c_opts)

        ttk.Label(self.calc_grid_f, text="平台費率:").grid(row=2, column=0, **c_opts)
        self.combo_calc_fee = ttk.Combobox(self.calc_grid_f, textvariable=self.var_calc_fee_rate, state="readonly", width=18)
        self.combo_calc_fee.grid(row=2, column=1, **c_opts)
        if hasattr(self, 'fee_lookup'):
            self.combo_calc_fee['values'] = list(self.fee_lookup.keys())

        ttk.Label(self.calc_grid_f, text="預期利潤:").grid(row=3, column=0, **c_opts)
        ttk.Entry(self.calc_grid_f, textvariable=self.var_calc_profit_val, width=10).grid(row=3, column=1, **c_opts)
        self.combo_profit_type = ttk.Combobox(self.calc_grid_f, textvariable=self.var_calc_profit_type, 
                                             values=["百分比(%)", "固定金額($)"], state="readonly", width=10)
        self.combo_profit_type.grid(row=3, column=2, padx=5, **c_opts)

        ttk.Separator(self.calc_grid_f, orient="horizontal").grid(row=4, column=0, columnspan=4, sticky="ew", pady=15)

        result_f = ttk.Frame(self.calc_grid_f)
        result_f.grid(row=5, column=0, columnspan=4, sticky="ew")
        ttk.Label(result_f, text="💡 建議上架售價:", font=("", 12, "bold")).pack(side="left")
        ttk.Label(result_f, textvariable=self.var_calc_target_price, font=("Arial", 18, "bold"), foreground="#d9534f").pack(side="left", padx=10)
        ttk.Label(result_f, text="元").pack(side="left")

        # 綁定即時計算
        self.var_calc_cost.trace_add("write", lambda *args: self.run_pricing_calc())
        self.var_calc_profit_val.trace_add("write", lambda *args: self.run_pricing_calc())
        self.combo_calc_fee.bind("<<ComboboxSelected>>", lambda e: self.run_pricing_calc())
        self.combo_profit_type.bind("<<ComboboxSelected>>", lambda e: self.run_pricing_calc())

        self.generate_procurement_report()
        self.update_calc_prod_list()


    @thread_safe_file
    def generate_procurement_report(self):
        """ 採購需求分析 V5.5：精確 ROP 模型，只顯示需要採購的商品 """
        if not hasattr(self, 'tree_procure'): return
        for i in self.tree_procure.get_children(): self.tree_procure.delete(i)
        
        try:
            if not os.path.exists(FILE_NAME): return
            with pd.ExcelFile(FILE_NAME) as xls:
                df_sales = pd.read_excel(xls, sheet_name=SHEET_SALES)
                df_prods = pd.read_excel(xls, sheet_name=SHEET_PRODUCTS)
            
            if df_prods.empty: return

            # 資料清洗與型別轉換
            for col in ['目前庫存', '安全庫存']:
                df_prods[col] = pd.to_numeric(df_prods[col], errors='coerce').fillna(0)
            df_sales['數量'] = pd.to_numeric(df_sales['數量'], errors='coerce').fillna(0)
            df_sales['日期'] = pd.to_datetime(df_sales['日期'], errors='coerce')

            now = pd.Timestamp.now()
            # 建立首賣日地圖 (與營收分析同步)
            first_sale_map = df_sales.groupby('商品名稱')['日期'].min().to_dict()
            qty_sum = df_sales.groupby('商品名稱')['數量'].sum()

            try:
                v_threshold = float(self.var_filter_velocity.get())
                cover_days = int(self.var_days_to_cover.get()) # 您設定的備貨天數 (例如 30)
            except:
                v_threshold, cover_days = 0.01, 30

            for _, row in df_prods.iterrows():
                p_name = str(row['商品名稱'])
                curr_stock = float(row['目前庫存'])
                base_safety = float(row['安全庫存'])
                
                # --- [1. 統一計算銷售速度] ---
                st_date = pd.to_datetime(row.get('初始上架時間'), errors='coerce')
                if pd.isna(st_date): st_date = first_sale_map.get(p_name, now)
                
                days_diff = max((now - st_date).days, 1)
                velocity = float(qty_sum.get(p_name, 0)) / days_diff
                
                # --- [2. 補貨點判定 (Reorder Point)] ---
                # 預期在等待收貨期間會賣出的量
                demand_during_leadtime = velocity * cover_days
                # 補貨臨界值 = 等待期間需求 + 安全庫存
                reorder_point = demand_during_leadtime + base_safety

                # --- [3. 過濾與顯示邏輯] ---
                # 只有符合以下任一條件才顯示：
                # A. 庫存已經低於補貨點且有在賣 (活躍商品)
                # B. 已經超賣 (負數)
                # C. 手動設定了安全庫存但現貨不足
                
                is_needed = False
                status = ""
                row_tag = ""
                
                if curr_stock < 0:
                    is_needed = True
                    status = "⚠️ 帳面超賣"
                    row_tag = "urgent"
                elif curr_stock <= reorder_point and velocity >= v_threshold:
                    is_needed = True
                    status = "🔴 需補貨"
                    row_tag = "urgent"
                elif curr_stock <= base_safety and base_safety > 0:
                    is_needed = True
                    status = "🟡 庫存偏低"
                    row_tag = "warning"
                
                # 若不需要補貨，直接跳過 (這能解決「庫存充足」佔用清單的問題)
                if not is_needed:
                    continue

                # 計算建議採購量 (補到 ROP 以上)
                import math
                suggest_qty = math.ceil(max(reorder_point - curr_stock, 0))

                self.tree_procure.insert("", "end", values=(
                    p_name, 
                    int(curr_stock), 
                    round(reorder_point, 1), # 這裡顯示補貨警戒值
                    f"{round(velocity, 2)}件/日", 
                    status, 
                    int(suggest_qty)
                ), tags=(row_tag,))
                
        except Exception as e:
            print(f"採購建議刷新失敗: {e}")


    def update_calc_prod_list(self):
        """ 初始化估算器的商品清單 """
        if not hasattr(self, 'list_calc_prod'): return
        self.list_calc_prod.delete(0, tk.END)
        if not self.products_df.empty:
            for _, row in self.products_df.iterrows():
                self.list_calc_prod.insert(tk.END, str(row['商品名稱']))

    def filter_calc_prod_list(self, event=None):
        """ 右側估算器專用的關鍵字過濾 """
        query = self.var_calc_search.get().lower().strip()
        self.list_calc_prod.delete(0, tk.END)
        
        if not self.products_df.empty:
            for name in self.products_df['商品名稱']:
                if query in str(name).lower():
                    self.list_calc_prod.insert(tk.END, name)

    def on_calc_prod_select(self, event):
        """ 當選取列表商品時，自動填入成本與預設資訊 """
        sel = self.list_calc_prod.curselection()
        if not sel: return
        
        p_name = self.list_calc_prod.get(sel[0])
        self.var_calc_name.set(p_name)
        
        # 從資料庫找成本
        record = self.products_df[self.products_df['商品名稱'] == p_name]
        if not record.empty:
            # 取得 WAC 成本
            cost = float(record.iloc[0].get('預設成本', 0.0))
            self.var_calc_cost.set(cost)
            
            # 自動觸發一次計算
            self.run_pricing_calc()


    def run_pricing_calc(self):
        """ 核心定價引擎：導入無條件進位 (Ceiling) 確保利潤空間 """
        try:
            from decimal import ROUND_CEILING # 確保匯入進位模式
            
            cost = Decimal(str(self.var_calc_cost.get()))
            profit_val = Decimal(str(self.var_calc_profit_val.get()))
            profit_type = self.var_calc_profit_type.get()
            
            fee_selection = self.var_calc_fee_rate.get()
            fee_perc = Decimal("0")
            fee_fixed = Decimal("0")
            
            if fee_selection in self.fee_lookup:
                p, f = self.fee_lookup[fee_selection]
                fee_perc = Decimal(str(p))
                fee_fixed = Decimal(str(f))
            
            r_fee = fee_perc / Decimal("100")
            
            if "百分比" in profit_type:
                r_profit = profit_val / Decimal("100")
                denominator = Decimal("1") - r_fee - r_profit
                if denominator <= 0:
                    self.var_calc_target_price.set("無法計算")
                    return
                target_price = (cost + fee_fixed) / denominator
            else:
                denominator = Decimal("1") - r_fee
                if denominator <= 0:
                    self.var_calc_target_price.set("錯誤")
                    return
                target_price = (cost + profit_val + fee_fixed) / denominator

            # --- [關鍵修正：無條件進位到整數] ---
            # .quantize(Decimal("1"), ...) 代表不保留小數點，且無條件進位
            final_price = target_price.quantize(Decimal("1"), rounding=ROUND_CEILING)
            self.var_calc_target_price.set(f"{float(final_price):,.0f}")

        except Exception:
            self.var_calc_target_price.set("0")

    # ================= 備份還原頁面 =================
    def setup_backup_tab(self):
        frame = ttk.Frame(self.tab_backup, padding=20)
        frame.pack(fill="both", expand=True)

           # ... (VIP 輸入區塊不用動) ...

        # 1. Google 帳號連結
        auth_frame = ttk.LabelFrame(frame, text="1. Google 帳號連結 (VIP 限定)", padding=15)
        auth_frame.pack(fill="x", pady=10)
        
        # 預設顯示：請先解鎖
        self.lbl_auth_status = ttk.Label(auth_frame, text="狀態: 🔒 請先輸入啟用碼解鎖", foreground="gray")
        self.lbl_auth_status.pack(side="left", padx=10)
        
        # 【修正點 1】這裡加上 state="disabled"
        self.btn_login = ttk.Button(auth_frame, text="登入 Google 帳號", command=self.start_login_thread, state="disabled")
        self.btn_login.pack(side="right")

        # 2. 備份操作區塊
        op_frame = ttk.LabelFrame(frame, text="2. 檔案備份與還原 (自動存入「蝦皮進銷存系統_備份」)", padding=15)
        op_frame.pack(fill="both", expand=True, pady=10)

        up_frame = ttk.Frame(op_frame)
        up_frame.pack(fill="x", pady=5)
        ttk.Label(up_frame, text="將目前的 Excel 檔案備份到雲端 (建議每日執行):").pack(side="left")
        
        # 【修正點 2】這裡加上 state="disabled"
        self.btn_upload = ttk.Button(up_frame, text="⬆️ 上傳備份", command=self.start_upload_thread, state="disabled")
        self.btn_upload.pack(side="right")

        ttk.Separator(op_frame, orient="horizontal").pack(fill="x", pady=15)

        ttk.Label(op_frame, text="3. 歷史備份紀錄 (雙擊項目可還原):").pack(anchor="w")
        
        cols = ("檔名", "備份時間")
        self.tree_backup = ttk.Treeview(op_frame, columns=cols, show='headings', height=10)
        # ... (Treeview 設定略) ...
        self.tree_backup.pack(fill="both", expand=True, pady=5)
        self.tree_backup.bind("<Double-1>", self.action_restore_backup)

        # 【修正點 3】這裡加上 state="disabled"
        self.btn_refresh = ttk.Button(op_frame, text="🔄 重新整理列表", command=self.start_list_thread, state="disabled")
        self.btn_refresh.pack(fill="x", pady=5)

        # ... (VIP 輸入框建立程式碼略) ...


        # === VIP 驗證區塊 ===
        

        vip_frame = ttk.LabelFrame(frame, text="🔒 進階功能解鎖", padding=15)
        vip_frame.pack(fill="x", pady=10)


        # 新增欄位：讓客戶輸入他的帳號
        ttk.Label(vip_frame, text="授權帳號(Email):").pack(side="left")
        self.var_vip_user = tk.StringVar()
        ttk.Entry(vip_frame, textvariable=self.var_vip_user, width=20).pack(side="left", padx=5)

        ttk.Label(vip_frame, text="啟用碼:").pack(side="left")
        self.var_vip_code = tk.StringVar()
        ttk.Entry(vip_frame, textvariable=self.var_vip_code, width=15).pack(side="left", padx=5)
        

        btn_unlock = ttk.Button(vip_frame, text="解鎖", command=self.unlock_vip_features)
        btn_unlock.pack(side="left", padx=10)

        # ... (後面的按鈕預設 disabled 邏輯同上)

    def unlock_vip_features(self):
        user_id = self.var_vip_user.get().strip()
        input_code = self.var_vip_code.get().strip().upper()
        
        if not user_id or not input_code:
            messagebox.showwarning("提示", "請輸入授權帳號與啟用碼")
            return

        # --- [核心邏輯 A：驗證啟用碼是否屬於該 Email] ---
        # 這裡的算法與你的 KeyGenApp 必須一致：SHA256(Email + Salt)
        salt = globals().get('SECRET_SALT', "redmaple")
        raw_string = user_id + salt
        expected_code = hashlib.sha256(raw_string.encode()).hexdigest()[:10].upper()

        if input_code == expected_code:
            mid = get_machine_id() # 抓取這台電腦 ID
            self.is_vip = True
            
            try:
                save_data = {
                    "user_id": user_id,
                    "license_key": input_code,
                    "machine_history": [mid]  # 初始化紀錄清單，這台是第一台
                }
                with open("license.json", "w", encoding="utf-8") as f:
                    json.dump(save_data, f)
                
                messagebox.showinfo("成功", "VIP 功能已解鎖！授權已與此電腦綁定。")
                self.refresh_backup_ui_status() # 更新 UI 按鈕狀態
            except Exception as e:
                messagebox.showerror("錯誤", f"授權存檔失敗: {e}")
        else:
            messagebox.showerror("錯誤", "啟用碼錯誤！")
    

    def check_license_on_startup(self):
        """
        程式啟動時檢查授權：不檢查路徑，只檢查機器指紋。
        支援自動換機登記 (最多 3 台)。
        """
        if not os.path.exists("license.json"):
            return 
            
        try:
            with open("license.json", "r", encoding="utf-8") as f:
                data = json.load(f)
            
            saved_user = data.get("user_id", "")
            saved_key = data.get("license_key", "")
            # 讀取已授權的機器清單
            history = data.get("machine_history", [])
            
            # 1. 重新驗證金鑰金鑰本身的正確性 (防人為修改 JSON)
            salt = globals().get('SECRET_SALT', "redmaple")
            raw_string = saved_user + salt
            expected_code = hashlib.sha256(raw_string.encode()).hexdigest()[:10].upper()
            
            if saved_key != expected_code:
                print("system: license key tampered.")
                return

            # 2. 機器指紋檢查
            current_mid = get_machine_id()
            
            if current_mid in history:
                # 情況一：這台電腦本來就在授權名單內
                self.is_vip = True
            else:
                # 情況二：這是一台「新電腦」
                if len(history) < 3: # 假設上限為 3 次換機機會
                    history.append(current_mid)
                    data["machine_history"] = history
                    
                    # 更新 JSON (下次啟動就不用再加了)
                    with open("license.json", "w", encoding="utf-8") as f:
                        json.dump(data, f)
                    
                    self.is_vip = True
                    messagebox.showinfo("授權更新", f"偵測到新環境。已自動綁定（剩餘更換次數：{3 - len(history)}）")
                else:
                    # 情況三：超過換機上限
                    messagebox.showerror("授權超限", "此授權已在超過 3 台電腦上使用過，請連繫開發者。")
                    self.is_vip = False
                    return

            # 3. 如果通過驗證，更新 UI
            if self.is_vip:
                self.var_vip_user.set(saved_user)
                self.var_vip_code.set(saved_key)
                self.btn_login.config(state="normal")
                self.lbl_auth_status.config(text="狀態: 🔒 VIP 授權有效 (自動登入)", foreground="green")
                # 如果 Google 已登入則解鎖備份按鈕
                if self.drive_manager.is_authenticated:
                    self.btn_upload.config(state="normal")
                    self.btn_refresh.config(state="normal")

        except Exception as e:
            print(f"system: failed to read license: {e}")

    def refresh_backup_ui_status(self):
        """ 解鎖成功後，立即啟用相關按鈕 """
        self.btn_login.config(state="normal")
        self.lbl_auth_status.config(text="狀態: ✅ VIP 已解鎖 (尚未連結 Google)", foreground="blue")

    # --- 執行緒相關函數 ---
    def start_login_thread(self):
        self.btn_login.config(state="disabled")
        self.lbl_auth_status.config(text="狀態: 正在開啟瀏覽器...請稍候", foreground="orange")
        threading.Thread(target=self._run_login, daemon=True).start()

    def _run_login(self):
        success, msg = self.drive_manager.authenticate()
        self.root.after(0, lambda: self._login_callback(success, msg))

    def _login_callback(self, success, msg):
        self.btn_login.config(state="normal")
        if success:
            self.lbl_auth_status.config(text=f"狀態: 登入成功", foreground="green")
            
            # 【修正點 5】登入成功後，解鎖功能按鈕
            self.btn_upload.config(state="normal")
            self.btn_refresh.config(state="normal")
            
            self.start_list_thread() 
        else:
            self.lbl_auth_status.config(text=f"狀態: {msg}", foreground="red")
            messagebox.showerror("登入錯誤", msg)


    @thread_safe_file
    def start_upload_thread(self):
        if not self.drive_manager.is_authenticated:
            messagebox.showwarning("警告", "請先登入 Google 帳號！")
            return
        if not os.path.exists(FILE_NAME):
            messagebox.showerror("錯誤", "找不到 Excel 檔案！")
            return
            
        self.btn_upload.config(state="disabled", text="上傳中...")
        threading.Thread(target=self._run_upload, daemon=True).start()

    def _run_upload(self):
        success, msg = self.drive_manager.upload_file(FILE_NAME)
        self.root.after(0, lambda: self._upload_callback(success, msg))

    def _upload_callback(self, success, msg):
        self.btn_upload.config(state="normal", text="⬆️ 上傳備份")
        if success:
            messagebox.showinfo("成功", msg)
            self.start_list_thread()
        else:
            messagebox.showerror("失敗", msg)

    def start_list_thread(self):
        if not self.drive_manager.is_authenticated: return
        self.btn_refresh.config(state="disabled", text="讀取中...")
        threading.Thread(target=self._run_list, daemon=True).start()

    def _run_list(self):
        files = self.drive_manager.list_backups()
        self.root.after(0, lambda: self._list_callback(files))

    def _list_callback(self, files):
        self.btn_refresh.config(state="normal", text="🔄 重新整理列表")
        for item in self.tree_backup.get_children():
            self.tree_backup.delete(item)
            
        if not files: return

        for f in files:
            raw_time = f.get('createdTime', '')
            try:
                # 1. 讀取 Google 回傳的 UTC 時間
                dt = datetime.strptime(raw_time, "%Y-%m-%dT%H:%M:%S.%fZ")
                # 2. 自動加 8 小時 (修正為台灣時間)
                dt = dt + timedelta(hours=8)
                nice_time = dt.strftime("%Y-%m-%d %H:%M")
            except:
                nice_time = raw_time
            
            self.tree_backup.insert("", "end", values=(f['name'], nice_time), tags=(f['id'],))

    def action_restore_backup(self, event):
        item_id = self.tree_backup.selection()
        if not item_id: return
        
        item = self.tree_backup.item(item_id)
        file_name = item['values'][0]
        file_id = self.tree_backup.item(item_id, "tags")[0]

        confirm = messagebox.askyesno("⚠️ 危險操作：確認還原？", 
                                      f"您確定要將資料還原成：\n{file_name}\n\n注意：這將會「覆蓋」目前電腦上所有的銷售與庫存紀錄！")
        if confirm:
            success, msg = self.drive_manager.download_file(file_id, FILE_NAME)
            if success:
                messagebox.showinfo("還原完成", msg)
                self.products_df = self.load_products()
                self.update_sales_prod_list()
                self.update_mgmt_prod_list()
            else:
                messagebox.showerror("還原失敗", msg)

    # ================= 銷售輸入頁面 (不變) =================
    def setup_sales_tab(self):
        top_frame = ttk.LabelFrame(self.tab_sales, text="訂單基本資料", padding=10)
        top_frame.pack(fill="x", padx=10, pady=5)

        r1 = ttk.Frame(top_frame)
        r1.pack(fill="x", pady=2)
        ttk.Label(r1, text="訂單日期:").pack(side="left")
        ttk.Entry(r1, textvariable=self.var_date, width=12).pack(side="left", padx=5)
        
        chk = ttk.Checkbutton(r1, text="填寫來源與顧客", variable=self.var_enable_cust, command=self.toggle_cust_info)
        chk.pack(side="left", padx=20)

        self.cust_frame = ttk.Frame(top_frame)
        self.cust_frame.pack(fill="x", pady=5)
        
        ttk.Label(self.cust_frame, text="交易平台:").grid(row=0, column=0, sticky="w", padx=2)
        self.combo_platform = ttk.Combobox(self.cust_frame, textvariable=self.var_platform, values=PLATFORM_OPTIONS, state="readonly", width=14)
        self.combo_platform.grid(row=0, column=1, padx=5)
        self.combo_platform.set("蝦皮購物")

        ttk.Label(self.cust_frame, text="買家名稱(ID):").grid(row=0, column=2, sticky="w", padx=10)
        self.entry_cust_name = ttk.Entry(self.cust_frame, textvariable=self.var_cust_name, width=15)
        self.entry_cust_name.grid(row=0, column=3, padx=5)

        ttk.Label(self.cust_frame, text="物流方式:").grid(row=1, column=0, sticky="w", padx=2, pady=5)
        self.combo_ship = ttk.Combobox(self.cust_frame, textvariable=self.var_ship_method, values=SHIPPING_METHODS, state="readonly", width=14)
        self.combo_ship.grid(row=1, column=1, padx=5, pady=5)
        self.combo_ship.bind("<<ComboboxSelected>>", self.on_ship_method_change)

        ttk.Label(self.cust_frame, text="取貨縣市:").grid(row=1, column=2, sticky="w", padx=10, pady=5)
        self.combo_loc = ttk.Combobox(self.cust_frame, textvariable=self.var_cust_loc, values=TAIWAN_CITIES, width=13)
        self.combo_loc.grid(row=1, column=3, padx=5, pady=5)
        self.combo_loc.bind('<KeyRelease>', self.filter_cities)

        self.toggle_cust_info()

        paned = ttk.PanedWindow(self.tab_sales, orient=tk.HORIZONTAL)
        paned.pack(fill="both", expand=True, padx=10, pady=5)

        left_frame = ttk.LabelFrame(paned, text="加入商品", padding=10)
        paned.add(left_frame, weight=1)

        ttk.Label(left_frame, text="搜尋:").pack(anchor="w")
        entry_search = ttk.Entry(left_frame, textvariable=self.var_search)
        entry_search.pack(fill="x", pady=5)
        entry_search.bind('<KeyRelease>', self.update_sales_prod_list)

        list_frame = ttk.Frame(left_frame)
        list_frame.pack(fill="both", expand=True, pady=5)
        self.listbox_sales = tk.Listbox(list_frame, height=10)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.listbox_sales.yview)
        self.listbox_sales.configure(yscrollcommand=scrollbar.set)
        self.listbox_sales.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.listbox_sales.bind('<<ListboxSelect>>', self.on_sales_prod_select)
        
        self.update_sales_prod_list()

        detail_frame = ttk.Frame(left_frame)
        detail_frame.pack(fill="x", pady=5)
        
        grid_opts = {'sticky': 'w', 'padx': 2, 'pady': 2}
        ttk.Label(detail_frame, text="已選:").grid(row=0, column=0, **grid_opts)
        ttk.Entry(detail_frame, textvariable=self.var_sel_name, state='readonly').grid(row=0, column=1, sticky="ew")
        
        ttk.Label(detail_frame, text="庫存:").grid(row=1, column=0, **grid_opts)
        lbl_stock = ttk.Label(detail_frame, textvariable=self.var_sel_stock_info, foreground="blue")
        lbl_stock.grid(row=1, column=1, sticky="w", padx=2)

        ttk.Label(detail_frame, text="售價:").grid(row=2, column=0, **grid_opts)
        ttk.Entry(detail_frame, textvariable=self.var_sel_price).grid(row=2, column=1, sticky="ew")

        ttk.Label(detail_frame, text="數量:").grid(row=3, column=0, **grid_opts)
        ttk.Entry(detail_frame, textvariable=self.var_sel_qty).grid(row=3, column=1, sticky="ew")

        ttk.Label(detail_frame, text="成本:").grid(row=4, column=0, **grid_opts)
        ttk.Entry(detail_frame, textvariable=self.var_sel_cost).grid(row=4, column=1, sticky="ew")

        ttk.Button(detail_frame, text="加入清單 ->", command=self.add_to_cart).grid(row=5, column=0, columnspan=2, pady=10, sticky="ew")

        right_frame = ttk.LabelFrame(paned, text="訂單內容", padding=10)
        paned.add(right_frame, weight=2)

        cols = ("編號", "商品名稱", "數量", "單價", "總計")
        self.tree = ttk.Treeview(right_frame, columns=cols, show='headings', height=8)
        self.tree.heading("編號", text="編號/位置",anchor="w")
        self.tree.column("編號", width=80) 
        self.tree.heading("商品名稱", text="商品名稱",anchor="w")
        self.tree.column("商品名稱", width=120)
        self.tree.heading("單價", text="售價")
        self.tree.column("單價", width=80, anchor="e")
        self.tree.heading("數量", text="數量")
        self.tree.column("數量", width=60, anchor="center")
        self.tree.heading("總計", text="小計")
        self.tree.column("總計", width=70, anchor="e")
        self.tree.pack(fill="both", expand=True)


        sales_btn_f = ttk.Frame(right_frame)
        sales_btn_f.pack(fill="x", pady=2)

        ttk.Button(sales_btn_f, text="🔼 排序上移", command=self.move_sales_item_up).pack(side="left", padx=2)
        ttk.Button(sales_btn_f, text="🔽 排序下移", command=self.move_sales_item_down).pack(side="left", padx=2)
        ttk.Button(sales_btn_f, text="(x) 移除", command=self.remove_from_cart).pack(side="right", padx=10)

        fee_frame = ttk.LabelFrame(right_frame, text="費用與折扣", padding=10)
        fee_frame.pack(fill="x", pady=5)
        
        # 第一排：平台費率
        f1 = ttk.Frame(fee_frame)
        f1.pack(fill="x")
        ttk.Label(f1, text="平台費率:").pack(side="left")
        self.combo_fee_rate = ttk.Combobox(f1, textvariable=self.var_fee_rate_str, state="readonly", width=28)
        self.combo_fee_rate.pack(side="left", padx=5)
        self.combo_fee_rate.bind('<<ComboboxSelected>>', self.on_fee_option_selected)

        # [新增這行]：讓使用者手動輸入數字時，也能即時觸發計算
        self.combo_fee_rate.bind('<KeyRelease>', self.update_totals_event)


        # 第二排：物流運費 (新增)
        f_ship = ttk.Frame(fee_frame)
        f_ship.pack(fill="x", pady=5)
        
        ttk.Label(f_ship, text="物流運費:").pack(side="left")
        ent_ship = ttk.Entry(f_ship, textvariable=self.var_ship_fee, width=8)
        ent_ship.pack(side="left", padx=5)
        ent_ship.bind('<KeyRelease>', self.update_totals_event)
        
        # 加入支付方選擇
        self.combo_payer = ttk.Combobox(f_ship, textvariable=self.var_ship_payer, 
                                        values=["買家付", "賣家付"], state="readonly", width=7)
        self.combo_payer.pack(side="left", padx=5)
        self.combo_payer.bind('<<ComboboxSelected>>', lambda e: self.update_totals())
        
        ttk.Label(f_ship, text="(影響出貨單總額與利潤)", foreground="gray", font=("", 9)).pack(side="left")

        # 第三排：扣費與折扣 (移除運費補貼，加入折扣券)
        f2 = ttk.Frame(fee_frame)
        f2.pack(fill="x", pady=5)
        
        ttk.Label(f2, text="折扣/扣費:").pack(side="left")

    

        # 移除 "運費補貼" 選項，改為更精確的標籤
        tag_opts = ["", "折扣券", "蝦幣折抵", "活動費", "補償金額", "私人預定", "補寄補貼", "固定成本"]
        self.combo_tag = ttk.Combobox(f2, textvariable=self.var_fee_tag, values=tag_opts, state="readonly", width=12)
        self.combo_tag.pack(side="left", padx=5)
        self.var_fee_tag.set("")
        self.combo_tag.set("扣費原因")

        ttk.Label(f2, text=" 金額$").pack(side="left", padx=2)
        e_extra = ttk.Entry(f2, textvariable=self.var_extra_fee, width=8)
        e_extra.pack(side="left")
        e_extra.bind('<KeyRelease>', self.update_totals_event)

        btn_print = ttk.Button(f2, text="📄 產生出貨單(預覽)", command=self.export_shipping_note)
        btn_print.pack(side="right", padx=10) # 加上 padx 讓按鈕與標籤有間距

        
        sum_frame = ttk.Frame(right_frame, relief="groove", padding=5)
        sum_frame.pack(fill="x", side="bottom")
        
        self.lbl_gross = ttk.Label(sum_frame, text="總金額: $0")
        self.lbl_gross.pack(anchor="w")
        self.lbl_fee = ttk.Label(sum_frame, text="扣費: $0", foreground="blue")
        self.lbl_fee.pack(anchor="w")
        self.lbl_profit = ttk.Label(sum_frame, text="實收淨利: $0", foreground="green")
        self.lbl_profit.pack(anchor="w")
        self.lbl_income = ttk.Label(sum_frame, text="預估入帳: $0", foreground="#ff0800")
        self.lbl_income.pack(anchor="w")
        

        btn_area = ttk.Frame(sum_frame)
        btn_area.pack(fill="x", pady=5)
        
        ttk.Button(sum_frame, text="✔ 送出訂單", command=self.submit_order).pack(fill="x", pady=5)

        self.refresh_fee_tree()

    def move_sales_item_up(self):
        """ 將銷售清單中的選中項目上移 """
        selected = self.tree.selection()
        if not selected: return
        
        for item in selected:
            idx = self.tree.index(item)
            if idx > 0:
                # 1. 移動 Treeview 視覺位置
                self.tree.move(item, '', idx - 1)
                
                # 2. 同步移動後台 cart_data 列表
                self.cart_data[idx], self.cart_data[idx-1] = \
                    self.cart_data[idx-1], self.cart_data[idx]

    def move_sales_item_down(self):
        """ 將銷售清單中的選中項目下移 """
        selected = self.tree.selection()
        if not selected: return
        
        # 下移要倒著處理，防止索引位移問題
        for item in reversed(selected):
            idx = self.tree.index(item)
            if idx < len(self.tree.get_children()) - 1:
                # 1. 移動 Treeview 視覺位置
                self.tree.move(item, '', idx + 1)
                
                # 2. 同步移動後台資料
                self.cart_data[idx], self.cart_data[idx+1] = \
                    self.cart_data[idx+1], self.cart_data[idx]



    def export_shipping_note(self):
        """ 呼叫外部模組產生出貨單 """
        if not self.cart_data:
            messagebox.showwarning("提示", "購物車內沒有商品")
            return

        # 彙整目前畫面的資料包
        order_info = {
            "shop_name": self.var_shop_name.get(), # 抓取設定頁面的店名
            "buyer": self.var_cust_name.get() if self.var_enable_cust.get() else "一般零售",
            "date": self.var_date.get(),
            "platform": self.var_platform.get(),
            "ship_method": self.var_ship_method.get(),
            "ship_fee": self.var_ship_fee.get(),
            "payer": self.var_ship_payer.get(),
            "discount_tag": self.var_fee_tag.get() if self.var_fee_tag.get() != "扣費原因" else "優惠折抵",
            "discount_amount": self.var_extra_fee.get()
        }

        # 呼叫彈窗讓賣家選尺寸，選完後會自動執行後續列印
        show_shipping_dialog(self.root, order_info, self.cart_data)



    def setup_product_tab(self):
        """ [修正版] 建立商品資料管理：修正 Tag 讀取與及時搜尋功能 """
        # --- 1. 初始化變數 ---
        self.var_add_sku = tk.StringVar() # 新增用的編號
        self.var_upd_sku = tk.StringVar() # 修改用的編號
        self.var_add_tag = tk.StringVar()
        self.var_add_name = tk.StringVar()
        self.var_add_url = tk.StringVar()
        self.var_add_remarks = tk.StringVar()
        self.var_add_safety = tk.IntVar(value=0)

        self.var_upd_tag = tk.StringVar()
        self.var_upd_name = tk.StringVar()
        self.var_upd_url = tk.StringVar()
        self.var_upd_remarks = tk.StringVar()
        self.var_upd_safety = tk.IntVar(value=0)
        self.var_upd_stock = tk.IntVar(value=0)
        self.var_upd_cost = tk.DoubleVar(value=0.0)
        self.var_upd_time = tk.StringVar(value="尚未選擇商品")

        # 主容器

        if hasattr(self, 'product_main_container'):
            self.product_main_container.destroy()
        
        self.product_main_container = ttk.Frame(self.tab_products)
        self.product_main_container.pack(fill="both", expand=True)

        paned = ttk.PanedWindow(self.product_main_container, orient=tk.HORIZONTAL)
        paned.pack(fill="both", expand=True, padx=10, pady=10)

        
         # --- 左側：新商品建檔 ---
        self.frame_left = ttk.LabelFrame(paned, text="🆕 新商品建檔", padding=15)
        paned.add(self.frame_left, weight=1)
        
        self.render_add_area() # 渲染左側輸入區

        # --- 右側：資料查詢與維護 ---
        self.frame_right = ttk.LabelFrame(paned, text="🔍 商品資料維護", padding=15)
        paned.add(self.frame_right, weight=1)
        
        # 搜尋與列表 (這部分固定顯示)
        ent_search = ttk.Entry(self.frame_right, textvariable=self.var_mgmt_search)
        ent_search.pack(fill="x")
        ent_search.bind('<KeyRelease>', lambda e: self.update_mgmt_prod_list())

        self.listbox_mgmt = tk.Listbox(self.frame_right, height=8)
        self.listbox_mgmt.pack(fill="both", expand=True, pady=5)
        self.listbox_mgmt.bind('<<ListboxSelect>>', self.on_mgmt_prod_select)

        self.edit_frame = ttk.LabelFrame(self.frame_right, text="✏️ 快速編輯資料", padding=10)
        self.edit_frame.pack(fill="x")
        
        self.render_edit_area() # 渲染右側編輯區

        # 底部按鈕 (固定顯示)
        btn_f = ttk.Frame(self.edit_frame)
        btn_f.grid(row=20, column=0, columnspan=4, pady=10) # row給大一點確保在底部
        ttk.Button(btn_f, text="💾 儲存修改", command=self.submit_update_product).pack(side="left", padx=5)
        ttk.Button(btn_f, text="🗑️ 刪除商品", command=self.delete_product).pack(side="left", padx=5)

        self.update_mgmt_prod_list()

        # 初始載入清單

    def render_add_area(self):
        """ 動態渲染左側建檔區 """
        for w in self.frame_left.winfo_children(): w.destroy()
        
        # 1. 商品編號 (可選)
        if self.show_fields["商品編號"].get():
            ttk.Label(self.frame_left, text="商品編號 (位置):").pack(anchor="w")
            ttk.Entry(self.frame_left, textvariable=self.var_add_sku).pack(fill="x", pady=2)
        
        # 2. 分類Tag (可選)
        if self.show_fields["分類Tag"].get():
            ttk.Label(self.frame_left, text="分類 Tag:").pack(anchor="w")
            self.combo_add_tag = ttk.Combobox(self.frame_left, textvariable=self.var_add_tag)
            self.combo_add_tag.pack(fill="x", pady=2)
            self.combo_add_tag.bind('<Button-1>', self.load_existing_tags)

        # 3. 商品名稱 (必填)
        ttk.Label(self.frame_left, text="* 商品名稱:").pack(anchor="w")
        ttk.Entry(self.frame_left, textvariable=self.var_add_name).pack(fill="x", pady=2)

        #4. 單位權重 
        if self.show_fields["單位權重"].get():
            ttk.Label(self.frame_left, text="單位權重 (g):").pack(anchor="w")
            ttk.Entry(self.frame_left, textvariable=self.var_add_weight).pack(fill="x", pady=2)

        # 4. 安全庫存 (可選)
        if self.show_fields["安全庫存"].get():
            ttk.Label(self.frame_left, text="安全庫存量:").pack(anchor="w")
            ttk.Entry(self.frame_left, textvariable=self.var_add_safety).pack(fill="x", pady=2)

        # 5. 連結與備註 (可選)
        if self.show_fields["商品連結"].get():
            ttk.Label(self.frame_left, text="採購連結 (URL):").pack(anchor="w")
            ttk.Entry(self.frame_left, textvariable=self.var_add_url).pack(fill="x", pady=2)
        
        if self.show_fields["商品備註"].get():
            ttk.Label(self.frame_left, text="商品備註:").pack(anchor="w")
            ttk.Entry(self.frame_left, textvariable=self.var_add_remarks).pack(fill="x", pady=2)

        ttk.Button(self.frame_left, text="✅ 完成建檔", command=self.submit_new_product).pack(fill="x", pady=15)


        ttk.Separator(self.frame_left, orient="horizontal").pack(fill="x", pady=10)
        
        ttk.Label(self.frame_left, text="📂 外部資料批次處理", font=("", 10, "bold")).pack(anchor="w")
        
        btn_wizard = ttk.Button(self.frame_left, text="📥 啟動商品批次匯入精靈", 
                                command=self.open_import_wizard)
        btn_wizard.pack(fill="x", pady=(5, 0))
        
        ttk.Label(self.frame_left, text="* 支援舊檔 Excel 欄位匹配匯入", 
                  foreground="gray", font=("", 9)).pack(anchor="w")

    def render_edit_area(self):
        """ 動態渲染右側編輯區 (使用 Grid) """
        for w in self.edit_frame.winfo_children(): 
            if w.winfo_class() != "TFrame": w.destroy() # 保留按鈕 Frame

        curr_row = 0
        e_opts = {'padx': 5, 'pady': 2, 'sticky': 'w'}

        # 必選欄位
        ttk.Label(self.edit_frame, text="名稱:").grid(row=curr_row, column=0, **e_opts)
        ttk.Entry(self.edit_frame, textvariable=self.var_upd_name, state="readonly").grid(row=curr_row, column=1, sticky="ew")
        

        if self.show_fields["商品編號"].get():
            ttk.Label(self.edit_frame, text="編號:").grid(row=curr_row, column=2, **e_opts)
            ttk.Entry(self.edit_frame, textvariable=self.var_upd_sku).grid(row=curr_row, column=3, sticky="ew")
        curr_row += 1

        if self.show_fields["單位權重"].get():
            ttk.Label(self.edit_frame, text="單位權重 (g/支):").grid(row=curr_row, column=0, **e_opts)
            ttk.Entry(self.edit_frame, textvariable=self.var_upd_weight).grid(row=curr_row, column=1, sticky="ew")
        curr_row += 1

        if self.show_fields["分類Tag"].get():
            ttk.Label(self.edit_frame, text="Tag:").grid(row=curr_row, column=0, **e_opts)
            self.combo_upd_tag = ttk.Combobox(self.edit_frame, textvariable=self.var_upd_tag)
            self.combo_upd_tag.grid(row=curr_row, column=1, sticky="ew")
        curr_row += 1

        # 庫存與成本 (必選)
        ttk.Label(self.edit_frame, text="庫存:").grid(row=curr_row, column=0, **e_opts)
        ttk.Entry(self.edit_frame, textvariable=self.var_upd_stock).grid(row=curr_row, column=1, sticky="ew")
        ttk.Label(self.edit_frame, text="成本:").grid(row=curr_row, column=2, **e_opts)
        ttk.Entry(self.edit_frame, textvariable=self.var_upd_cost).grid(row=curr_row, column=3, sticky="ew")
        curr_row += 1

        if self.show_fields["安全庫存"].get():
            ttk.Label(self.edit_frame, text="安全量:").grid(row=curr_row, column=0, **e_opts)
            ttk.Entry(self.edit_frame, textvariable=self.var_upd_safety).grid(row=curr_row, column=1, sticky="ew")
            curr_row += 1

        if self.show_fields["單位權重"].get():
            ttk.Label(self.edit_frame, text="單位權重(g):").grid(row=curr_row, column=0, **e_opts)
            ttk.Entry(self.edit_frame, textvariable=self.var_upd_weight).grid(row=curr_row, column=1, sticky="ew")
            curr_row += 1

        if self.show_fields["商品連結"].get():
            ttk.Label(self.edit_frame, text="連結:").grid(row=curr_row, column=0, **e_opts)
            ttk.Entry(self.edit_frame, textvariable=self.var_upd_url).grid(row=curr_row, column=1, columnspan=3, sticky="ew")
            curr_row += 1

        if self.show_fields["商品備註"].get():
            ttk.Label(self.edit_frame, text="備註:").grid(row=curr_row, column=0, **e_opts)
            ttk.Entry(self.edit_frame, textvariable=self.var_upd_remarks).grid(row=curr_row, column=1, columnspan=3, sticky="ew")

    def refresh_product_ui_layout(self):
        """ 當勾選設定改變時，重新繪製商品管理頁面 """
        self.setup_product_tab()

    

    def open_import_wizard(self):
        """ 開啟外部匯入精靈視窗 """
        # 這裡的 ImportWizard 是我們剛剛更新過支援「商品編號」的版本
        ImportWizard(self.root, self.callback_from_wizard)



    @thread_safe_file
    def callback_from_wizard(self, new_data_list):
        """ 當精靈完成匹配並按下確認時，接收資料並存入 Excel """
        if not new_data_list: return False
        
        try:
            df_new = pd.DataFrame(new_data_list)
            
            # 1. 讀取目前現有的商品資料
            with pd.ExcelFile(FILE_NAME) as xls:
                df_old = pd.read_excel(xls, sheet_name=SHEET_PRODUCTS)

            # 2. 合併資料
            # 將新舊資料合併，並根據「商品名稱」去重
            # keep='last' 代表如果名稱重複，以新匯入的資料為準
            df_combined = pd.concat([df_old, df_new], ignore_index=True)
            df_combined.drop_duplicates(subset=['商品名稱'], keep='last', inplace=True)
            
            # 3. 呼叫萬用引擎存檔 (確保分頁不消失)
            save_success = self._universal_save({SHEET_PRODUCTS: df_combined})
            
            if save_success:
                # 4. 成功後刷新介面資料
                self.products_df = self.load_products()
                self.update_mgmt_prod_list() # 刷新管理列表
                self.update_sales_prod_list() # 刷新銷售選單
                self.update_pur_prod_list()  # 刷新進貨列表
                return True
            return False
            
        except Exception as e:
            messagebox.showerror("匯入存檔失敗", f"錯誤原因: {str(e)}")
            return False



    def setup_tracking_tab(self):
        """ 建立訂單追蹤區 (緩衝區) """
        frame = self.tab_tracking
        # --- 1. 頂部操作與搜尋區 ---
        top_frame = ttk.Frame(frame, padding=10)
        top_frame.pack(fill="x")

        # 搜尋功能
        search_box = ttk.LabelFrame(top_frame, text="🔍 快速篩選訂單", padding=5)
        search_box.pack(side="left", fill="x", expand=True, padx=(0, 10))

        ttk.Label(search_box, text="關鍵字 (買家/商品):").pack(side="left", padx=5)
        self.var_track_search = tk.StringVar()
        # 綁定 KeyRelease 事件，達成「邊打字邊過濾」的效果
        ent_search = ttk.Entry(search_box, textvariable=self.var_track_search, width=30)
        ent_search.pack(side="left", padx=5)
        ent_search.bind("<KeyRelease>", lambda e: self.load_tracking_data())

        ttk.Button(top_frame, text="🔄 重新整理", command=self.load_tracking_data).pack(side="right", pady=10)


        # 2. 中間：列表
        tree_frame = ttk.Frame(frame)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=5)
        cols = ("訂單編號", "日期", "平台", "買家", "商品名稱", "數量", "售價")
        self.tree_track = ttk.Treeview(tree_frame, columns=cols, show='headings', height=15)
        for c in cols:
            self.tree_track.heading(c, text=c)
            self.tree_track.column(c, width=100 if "商品" not in c else 200)
        
        sb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree_track.yview)
        self.tree_track.configure(yscrollcommand=sb.set)
        self.tree_track.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        # 3. 下方：兩行按鈕區
        btn_main_frame = ttk.LabelFrame(frame, text="訂單操作面板", padding=10)
        btn_main_frame.pack(fill="x", padx=10, pady=10)

        # 第一行：修改與刪除
        row1 = ttk.Frame(btn_main_frame)
        row1.pack(fill="x", pady=2)
        ttk.Button(row1, text="✏️ 修改數量/售價", command=self.action_track_modify).pack(side="left", fill="x", expand=True, padx=2)
        ttk.Button(row1, text="➖ 刪除單一商品 (補位)", command=self.action_track_delete_item).pack(side="left", fill="x", expand=True, padx=2)
        ttk.Button(row1, text="🗑️ 刪除整筆訂單", command=self.action_track_delete_order).pack(side="left", fill="x", expand=True, padx=2)

        # 第二行：結案與退貨
        row2 = ttk.Frame(btn_main_frame)
        row2.pack(fill="x", pady=2)
        ttk.Button(row2, text="↩️ 退貨單一商品", command=self.action_track_return_item).pack(side="left", fill="x", expand=True, padx=2)
        ttk.Button(row2, text="⏪ 退貨整筆訂單", command=self.action_track_return_order).pack(side="left", fill="x", expand=True, padx=2)
        ttk.Button(row2, text="✅ 完成訂單 (整筆結案)", command=self.action_track_complete_order).pack(side="left", fill="x", expand=True, padx=2)

        self.load_tracking_data()


    @thread_safe_file
    def load_tracking_data(self):
        """ 讀取『訂單追蹤』分頁：使用分組填充，防止買家名稱錯誤繼承 """
        for i in self.tree_track.get_children():
            self.tree_track.delete(i)
            
        try:
            if not os.path.exists(FILE_NAME): return
            
            # 1. 讀取 Excel 原始資料
            df = pd.read_excel(FILE_NAME, sheet_name=SHEET_TRACKING)
            if df.empty: return

            # 2. 統一格式化訂單編號 (這是我們的分組依據)
            df['訂單編號'] = df['訂單編號'].astype(str).str.replace(r'^\'', '', regex=True).str.replace(r'\.0$', '', regex=True).str.strip()

            # 3. --- [核心修正：分組向下填充] ---
            # 建立副本進行顯示處理
            df_display = df.copy()
            
            # 定義需要補齊資訊的欄位
            fill_cols = ['日期', '買家名稱', '交易平台', '寄送方式', '取貨地點']
            
            # 【關鍵點】：按『訂單編號』分組後再執行 ffill
            # 這樣「訂單 A」的買家名稱絕對不會流到「訂單 B」
            df_display[fill_cols] = df_display.groupby('訂單編號')[fill_cols].ffill()
            
            # 如果分組填充完後還是 NaN (代表該訂單編號的第一行本來就沒寫買家)，則填入預設值
            df_display[fill_cols] = df_display[fill_cols].fillna("資訊缺失")

            # 4. 取得搜尋關鍵字
            query = self.var_track_search.get().strip().lower()

            # 5. 執行過濾 (在補齊資料後的副本上搜尋)
            if query:
                mask = (
                    df_display['買家名稱'].astype(str).str.lower().str.contains(query) |
                    df_display['商品名稱'].astype(str).str.lower().str.contains(query) |
                    df_display['訂單編號'].astype(str).str.lower().str.contains(query)
                )
                df_filtered = df_display[mask]
            else:
                df_filtered = df_display

            # 6. 填入 Treeview
            for idx, row in df_filtered.iterrows():
                # 使用 text=str(idx) 確保我們修改時能對應回 Excel 的原始列號
                self.tree_track.insert("", "end", text=str(idx), values=(
                    row.get('訂單編號', ''),
                    row.get('日期', ''),
                    row.get('交易平台', ''),
                    row.get('買家名稱', ''),
                    row.get('商品名稱', ''),
                    int(row.get('數量', 0)),
                    float(row.get('單價(售)', 0))
                ))
                
        except Exception as e:
            print(f"system: failed to load tracking list: {e}")


    @thread_safe_file
    def action_track_modify(self):
        """ 修改資料: 跳出視窗修改數量與價格 """
        sel = self.tree_track.selection()
        if not sel:
            messagebox.showwarning("提示", "請先選擇要修改的商品項目")
            return
        item = self.tree_track.item(sel[0]); idx = int(item['text']); vals = item['values']
        prod_name = vals[4]; old_qty = vals[5]; old_price = vals[6]
        win = tk.Toplevel(self.root); win.title(f"修改: {prod_name}"); win.geometry("300x200")
        tk.Label(win, text="數量:").pack(pady=5)
        var_qty = tk.IntVar(value=old_qty); tk.Entry(win, textvariable=var_qty).pack()
        tk.Label(win, text="售價:").pack(pady=5)
        var_price = tk.DoubleVar(value=old_price); tk.Entry(win, textvariable=var_price).pack()
        def save_mod():
            try:
                df = pd.read_excel(FILE_NAME, sheet_name=SHEET_TRACKING)
                new_qty = var_qty.get(); new_price = var_price.get()
                df.at[idx, '數量'] = new_qty; df.at[idx, '單價(售)'] = new_price
                cost = df.at[idx, '單價(進)']; fee = df.at[idx, '分攤手續費']
                df.at[idx, '總銷售額'] = new_qty * new_price
                df.at[idx, '總成本'] = new_qty * cost
                df.at[idx, '總淨利'] = (new_qty * new_price) - (new_qty * cost) - fee
                self._universal_save({ SHEET_TRACKING: df })
                messagebox.showinfo("成功", "資料已更新"); self.load_tracking_data(); win.destroy()
            except Exception as e: messagebox.showerror("錯誤", f"存檔失敗: {e}")
        tk.Button(win, text="確認修改", command=save_mod).pack(pady=15)



    @thread_safe_file
    def action_track_delete_item(self):
        """ 刪除單一商品 (含表頭自動遞補邏輯) """
        sel = self.tree_track.selection()
        if not sel: return
        item = self.tree_track.item(sel[0]); idx = int(item['text'])
        order_id = str(item['values'][0]); prod_name = str(item['values'][4])
        if not messagebox.askyesno("刪除商品", f"確定要從訂單 [{order_id}] 中\n刪除商品「{prod_name}」嗎？"): return
        try:
            df = pd.read_excel(FILE_NAME, sheet_name=SHEET_TRACKING)
            df['訂單編號'] = df['訂單編號'].astype(str).str.replace(r'\.0$', '', regex=True)
            is_header = pd.notna(df.at[idx, '日期']) and str(df.at[idx, '日期']) != ""
            if is_header:
                mask_others = (df['訂單編號'] == order_id) & (df.index != idx)
                others_indices = df[mask_others].index.tolist()
                if others_indices:
                    new_header_idx = others_indices[0]
                    cols_to_inherit = ['日期', '交易平台', '買家名稱', '寄送方式', '取貨地點', '扣費項目']
                    for col in cols_to_inherit: df.at[new_header_idx, col] = df.at[idx, col]
            df.drop(idx, inplace=True)
            self._universal_save({ SHEET_TRACKING: df })
            messagebox.showinfo("成功", "商品已刪除"); self.load_tracking_data()
        except Exception as e: messagebox.showerror("錯誤", f"刪除失敗: {e}")



    @thread_safe_file
    def action_track_delete_order(self):
        """ 刪除整筆訂單：強化比對邏輯，確保刪除成功 """
        sel = self.tree_track.selection()
        if not sel:
            messagebox.showwarning("提示", "請先選擇要刪除的訂單項目")
            return
        
        # 1. 取得介面上的訂單編號，並清理乾淨
        item = self.tree_track.item(sel[0])
        order_id = str(item['values'][0]).replace("'", "").strip()
        
        if not messagebox.askyesno("刪除確認", f"確定要刪除訂單 [{order_id}] 嗎？\n該訂單內的所有商品都會消失！"):
            return

        try:
            # 2. 讀取目前的追蹤清單
            df = pd.read_excel(FILE_NAME, sheet_name=SHEET_TRACKING)
            
            # 3. 【關鍵修正】：統一 Excel 內的編號格式以便比對
            # 全部轉字串 -> 去掉單引號 -> 去掉 .0
            df['訂單編號_清理'] = df['訂單編號'].astype(str).str.replace(r'^\'', '', regex=True).str.replace(r'\.0$', '', regex=True).str.strip()
            
            # 檢查是否存在該編號 (Debug 用)
            if order_id not in df['訂單編號_清理'].values:
                # 如果找不到，嘗試再次模糊比對
                messagebox.showwarning("刪除失敗", f"在資料庫中找不到編號: {order_id}\n請嘗試手動『重新整理』後再試一次。")
                return

            # 4. 執行過濾：只留下「不等於」該編號的資料
            df_new = df[df['訂單編號_清理'] != order_id].copy()
            
            # 刪除輔助欄位
            df_new.drop(columns=['訂單編號_清理'], inplace=True)

            # 5. 調用萬用存檔引擎 (字典格式)
            save_success = self._universal_save({SHEET_TRACKING: df_new})
            
            if save_success:
                messagebox.showinfo("成功", f"訂單 {order_id} 已從系統中移除。")
                # 6. 強制刷新介面
                self.load_tracking_data()
                
        except Exception as e:
            messagebox.showerror("錯誤", f"刪除操作失敗: {str(e)}")


    @thread_safe_file
    def action_track_return_order(self):
        """ 退貨整筆訂單 (修正存檔格式) """
        from tkinter import simpledialog
        sel = self.tree_track.selection()
        if not sel: return
        item = self.tree_track.item(sel[0]); order_id = str(item['values'][0]).replace("'", "")
        reason = simpledialog.askstring("整筆退貨", "請輸入整筆退貨原因:", parent=self.root)
        if reason is None: return
        
        try:
            df_track = pd.read_excel(FILE_NAME, sheet_name=SHEET_TRACKING)
            df_track['訂單編號'] = df_track['訂單編號'].astype(str).str.replace(r'^\'', '', regex=True).str.replace(r'\.0$', '', regex=True)
            mask = df_track['訂單編號'] == order_id
            rows_to_return = df_track[mask].copy()
            info = self._get_full_order_info(df_track, order_id)
            for col, val in info.items(): rows_to_return[col] = val
            rows_to_return['備註'] = reason
            
            try: df_returns = pd.read_excel(FILE_NAME, sheet_name=SHEET_RETURNS)
            except: df_returns = pd.DataFrame()
            df_returns = pd.concat([df_returns, rows_to_return], ignore_index=True)
            df_track_new = df_track[~mask]
            
            # ---【關鍵修正：使用大括號字典傳參】---
            success = self._universal_save({
                SHEET_TRACKING: df_track_new, 
                SHEET_RETURNS: df_returns
            })
            
            if success:
                messagebox.showinfo("成功", f"訂單 {order_id} 整筆已移至退貨。")
                self.load_tracking_data(); self.load_returns_data()
        except Exception as e: messagebox.showerror("錯誤", str(e))


   
    def setup_returns_tab(self):
        """ 建立退貨紀錄查詢頁面 """
        frame = self.tab_returns
        
        # 頂部控制
        top_frame = ttk.Frame(frame, padding=5)
        top_frame.pack(fill="x")
        ttk.Label(top_frame, text="⚠️ 退貨紀錄為存證性質，不可於此處修改或刪除。", foreground="red").pack(side="left")
        ttk.Button(top_frame, text="🔄 重新整理退貨清單", command=self.load_returns_data).pack(side="right")

        # 列表 Treeview (多了一個「退貨原因」欄位)
        cols = ("訂單編號", "日期", "買家", "商品名稱", "數量", "售價", "退貨原因")
        tree_frame = ttk.Frame(frame)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.tree_returns = ttk.Treeview(tree_frame, columns=cols, show='headings', height=20)
        
        # 設定標題與寬度
        widths = {"訂單編號": 120, "日期": 90, "買家": 100, "商品名稱": 180, "數量": 50, "售價": 60, "退貨原因": 250}
        for c in cols:
            self.tree_returns.heading(c, text=c)
            self.tree_returns.column(c, width=widths[c], anchor="w" if c != "數量" else "center")
        
        sb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree_returns.yview)
        self.tree_returns.configure(yscrollcommand=sb.set)
        self.tree_returns.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self.load_returns_data()

    @thread_safe_file
    def load_returns_data(self):
        """ 讀取『退貨紀錄』分頁的資料 """
        for i in self.tree_returns.get_children():
            self.tree_returns.delete(i)
            
        try:
            if not os.path.exists(FILE_NAME): return
            df = pd.read_excel(FILE_NAME, sheet_name=SHEET_RETURNS)
            
            # 格式化編號
            if '訂單編號' in df.columns:
                df['訂單編號'] = df['訂單編號'].astype(str).str.replace(r'^\'', '', regex=True).str.replace(r'\.0$', '', regex=True)
            
            df = df.fillna("")
            
            # 填入 Treeview
            for _, row in df.iterrows():
                self.tree_returns.insert("", "end", values=(
                    row.get('訂單編號', ''),
                    row.get('日期', ''),
                    row.get('買家名稱', ''),
                    row.get('商品名稱', ''),
                    row.get('數量', 0),
                    row.get('單價(售)', 0),
                    row.get('備註', '') # 對應 Excel Q 列的內容
                ))
        except Exception as e:
            print(f"system: failed to load returns data: {e}")

    #================= 銷售紀錄 =================
    def setup_sales_edit_tab(self):
        main_paned = ttk.PanedWindow(self.tab_sales_edit, orient=tk.VERTICAL)
        main_paned.pack(fill="both", expand=True, padx=10, pady=10)


        # 1. 上方：列表區
        list_frame = ttk.LabelFrame(main_paned, text="銷售歷史紀錄 (點擊項目進行檢視與售後處理)", padding=5)
        main_paned.add(list_frame, weight=3)

        # 建立 Treeview
        cols = ("日期", "買家名稱", "商品", "數量", "售價", "手續費", "淨利", "毛利")
        self.tree_sales_edit = ttk.Treeview(list_frame, columns=cols, show='headings', height=12)
        
        # 設定欄寬
        self.tree_sales_edit.heading("日期", text="日期"); self.tree_sales_edit.column("日期", width=90)
        self.tree_sales_edit.heading("買家名稱", text="買家名稱"); self.tree_sales_edit.column("買家名稱", width=80)
        self.tree_sales_edit.heading("商品", text="商品名稱"); self.tree_sales_edit.column("商品", width=150)
        self.tree_sales_edit.heading("數量", text="數量"); self.tree_sales_edit.column("數量", width=50, anchor="center")
        self.tree_sales_edit.heading("售價", text="售價"); self.tree_sales_edit.column("售價", width=60, anchor="e")
        self.tree_sales_edit.heading("手續費", text="手續費"); self.tree_sales_edit.column("手續費", width=60, anchor="e")
        self.tree_sales_edit.heading("淨利", text="淨利"); self.tree_sales_edit.column("淨利", width=60, anchor="e")
        self.tree_sales_edit.heading("毛利", text="毛利%"); self.tree_sales_edit.column("毛利", width=60, anchor="e")

        scrolly = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree_sales_edit.yview)
        self.tree_sales_edit.configure(yscrollcommand=scrolly.set)
        self.tree_sales_edit.pack(side="left", fill="both", expand=True)
        scrolly.pack(side="right", fill="y")
        
        # 綁定選擇事件
        self.tree_sales_edit.bind("<<TreeviewSelect>>", self.on_sales_edit_select)

        # 重新整理按鈕
        btn_refresh = ttk.Button(list_frame, text="🔄 重新讀取 Excel", command=self.load_sales_records_for_edit)
        btn_refresh.pack(fill="x", side="bottom")

        bottom_container = ttk.PanedWindow(main_paned, orient=tk.HORIZONTAL)
        main_paned.add(bottom_container, weight=2)


        # 2. 下方：改為「訂單詳情檢視 (唯讀)」
        detail_frame = ttk.LabelFrame(bottom_container, text="🔍 訂單完整詳情 (唯讀)", padding=15)
        bottom_container.add(detail_frame, weight=1)

        # 建立一組變數用來顯示
        self.var_view_oid = tk.StringVar()
        self.var_view_date = tk.StringVar()
        self.var_view_buyer = tk.StringVar()
        self.var_view_platform = tk.StringVar()
        self.var_view_ship = tk.StringVar()
        self.var_view_loc = tk.StringVar()
        self.var_view_item = tk.StringVar()
        self.var_view_tax = tk.StringVar()

        # 使用 Grid 排版顯示所有欄位
        opts = {'padx': 10, 'pady': 5, 'sticky': 'w'}
        ttk.Label(detail_frame, text="訂單編號:").grid(row=0, column=0, **opts)
        ttk.Label(detail_frame, textvariable=self.var_view_oid, foreground="blue", font=("Consolas", 10)).grid(row=0, column=1, **opts)

        ttk.Label(detail_frame, text="買家名稱:").grid(row=0, column=2, **opts)
        ttk.Label(detail_frame, textvariable=self.var_view_buyer, font=("", 10, "bold")).grid(row=0, column=3, **opts)

        ttk.Label(detail_frame, text="商品名稱:").grid(row=1, column=0, **opts)
        ttk.Label(detail_frame, textvariable=self.var_view_item, wraplength=400).grid(row=1, column=1, columnspan=3, **opts)

        ttk.Label(detail_frame, text="寄送方式:").grid(row=2, column=0, **opts)
        ttk.Label(detail_frame, textvariable=self.var_view_ship).grid(row=2, column=1, **opts)

        ttk.Label(detail_frame, text="取貨地點:").grid(row=2, column=2, **opts)
        ttk.Label(detail_frame, textvariable=self.var_view_loc).grid(row=2, column=3, **opts)

        ttk.Label(detail_frame, text="該品稅額:").grid(row=3, column=0, **opts)
        ttk.Label(detail_frame, textvariable=self.var_view_tax, foreground="red").grid(row=3, column=1, **opts)


        # --- 售後服務登記區 (UI) ---
        
        after_frame = ttk.LabelFrame(bottom_container, text="🛠️ 售後服務處理", padding=15)
        bottom_container.add(after_frame, weight=1)

        # --- 即時狀態顯示區 ---
        status_frame = ttk.Frame(after_frame, relief="flat")
        status_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        ttk.Label(status_frame, text="🚩 目前售後狀態：", font=("", 10, "bold")).pack(side="left")
        ttk.Label(status_frame, textvariable=self.var_view_after_status, foreground="#d9534f", wraplength=250).pack(side="left")

        ttk.Separator(after_frame, orient="horizontal").grid(row=1, column=0, columnspan=2, sticky="ew", pady=10)

        # --- 輸入區 ---
        a_opts = {'padx': 5, 'pady': 3, 'sticky': 'w'}
        ttk.Label(after_frame, text="處理類型:").grid(row=2, column=0, **a_opts)
        self.combo_after_type = ttk.Combobox(after_frame, textvariable=self.var_after_type, 
                                            values=["補寄商品", "補貼款/退部分金額", "換貨支出", "保固寄新", "其他支出"], state="readonly")
        self.combo_after_type.grid(row=2, column=1, **a_opts)

        ttk.Label(after_frame, text="額外支出($):").grid(row=3, column=0, **a_opts)
        ttk.Entry(after_frame, textvariable=self.var_after_cost, width=15).grid(row=3, column=1, **a_opts)

        ttk.Label(after_frame, text="售後說明:").grid(row=4, column=0, **a_opts)
        ttk.Entry(after_frame, textvariable=self.var_after_remark, width=25).grid(row=4, column=1, **a_opts)

        btn_after = ttk.Button(after_frame, text="🚀 提交售後紀錄", command=self.submit_after_sales)
        btn_after.grid(row=5, column=0, columnspan=2, pady=10)

        self.load_sales_records_for_edit()

        
    
    @thread_safe_file
    def submit_after_sales(self):
        sel = self.tree_sales_edit.selection()
        if not sel:
            messagebox.showwarning("提示", "請先從上方列表選擇要處理的歷史訂單項目")
            return
        
        # 取得選中項目在 Treeview 儲存的原始列索引 (idx)
        item = self.tree_sales_edit.item(sel[0])
        idx = int(item['text'])
        
        after_type = self.var_after_type.get()
        extra_cost = self.var_after_cost.get()
        after_remark = self.var_after_remark.get().strip()
        
        if not after_type:
            messagebox.showwarning("提示", "請選擇處理類型")
            return

        if not messagebox.askyesno("確認登記", f"確認登記售後服務？\n類型：{after_type}\n金額：${extra_cost}\n這將會直接扣除該訂單的淨利紀錄並更新庫存。"):
            return

        try:
            # 1. 讀取相關資料 (一次讀取多個分頁)
            with pd.ExcelFile(FILE_NAME) as xls:
                df_sales = pd.read_excel(xls, sheet_name=SHEET_SALES)
                df_prods = pd.read_excel(xls, sheet_name=SHEET_PRODUCTS)
            
            # 2. 更新銷售紀錄資料 (針對指定行 idx)
            # 扣除淨利
            old_profit = df_sales.at[idx, '總淨利']
            df_sales.at[idx, '總淨利'] = round(old_profit - extra_cost, 2)
            
            # 更新備註 (追加售後資訊)
            current_tags = str(df_sales.at[idx, '扣費項目']) if pd.notna(df_sales.at[idx, '扣費項目']) else ""
            if current_tags == "nan": current_tags = ""
            
            # 建立新的備註標記
            new_tag = f"[{after_type}:-${extra_cost}] {after_remark}"
            full_remark = f"{current_tags} {new_tag}".strip()
            df_sales.at[idx, '扣費項目'] = full_remark
            
            # 重新計算該行的毛利率 (因為淨利減少了)
            total_sales = df_sales.at[idx, '總銷售額']
            if total_sales > 0:
                new_margin = (df_sales.at[idx, '總淨利'] / total_sales) * 100
                df_sales.at[idx, '毛利率'] = round(new_margin, 1)

            # 3. 處理庫存扣除 (若屬於補寄類)
            # 只有在特定的處理類型下才自動扣庫存
            if after_type in ["補寄商品", "保固寄新"]:
                prod_name = df_sales.at[idx, '商品名稱']
                p_idx_list = df_prods[df_prods['商品名稱'] == prod_name].index
                if not p_idx_list.empty:
                    p_idx = p_idx_list[0]
                    old_stock = df_prods.at[p_idx, '目前庫存']
                    df_prods.at[p_idx, '目前庫存'] = old_stock - 1 # 預設補寄 1 個
                    print(f"system: deducted inventory for after-sales: {prod_name} from {old_stock} to {old_stock-1}")

            # 4. 調用萬用引擎一次性儲存 (確保資料一致性)
            save_dict = {
                SHEET_SALES: df_sales,
                SHEET_PRODUCTS: df_prods
            }
            
            if self._universal_save(save_dict):
                messagebox.showinfo("成功", "售後處理已完成！\n1. 淨利已重新計算\n2. 備註已更新\n3. 庫存已同步(若適用)")
                
                # --- [關鍵：即時更新介面顯示] ---
                # A. 更新記憶體內的商品資料
                self.products_df = df_prods 
                
                # B. 刷新銷售紀錄列表 (讓清單上的淨利、毛利數字變動)
                self.load_sales_records_for_edit()
                
                # C. 重設售後輸入框內容
                self.var_after_cost.set(0.0)
                self.var_after_remark.set("")
                
                # D. 重要：更新右側的「目前售後狀態」即時顯示標籤
                # 這裡直接把剛才算好的 full_remark 填進去，使用者就不需要重新點選一次
                self.var_view_after_status.set(full_remark)
                
                # E. 重新計算營收分析 (因為淨利變了)
                self.calculate_analysis_data()

        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("錯誤", f"售後登記作業失敗: {str(e)}")


    @thread_safe_file
    def load_sales_records_for_edit(self):
        """ 讀取銷售紀錄：同步執行資料填充與精準排序 """
        for i in self.tree_sales_edit.get_children():
            self.tree_sales_edit.delete(i)
        
        try:
            if not os.path.exists(FILE_NAME): return
            # 讀取原始資料
            df = pd.read_excel(FILE_NAME, sheet_name=SHEET_SALES)
            df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
            
            if df.empty: return

            # --- [關鍵修正 A：在記憶體中填充空值] ---
            # 確保統計與顯示時，每一行都有完整的買家資訊，解決 nan 問題
            fill_cols = ['訂單編號', '日期', '買家名稱', '交易平台', '寄送方式', '取貨地點']
            # 先將編號轉為字串方便分組
            df['訂單編號'] = df['訂單編號'].astype(str).str.replace(r'^\'', '', regex=True).str.replace(r'\.0$', '', regex=True)
            # 使用 ffill 補齊資訊
            df[fill_cols] = df.groupby('訂單編號', group_keys=False)[fill_cols].ffill().bfill()

            # --- [關鍵修正 B：排序並儲存至暫存變數] ---
            df['tmp_dt'] = pd.to_datetime(df['日期'], errors='coerce')
            df = df.sort_values(by=['tmp_dt', '訂單編號'], ascending=[False, False])
            
            # 將這份「排序好且填滿資訊」的資料存入 self，供點擊時讀取
            self.sales_edit_df = df.copy() 

            for idx, row in df.iterrows():
                # idx 現在是原始 DataFrame 的標籤 (Label)
                self.tree_sales_edit.insert("", "end", text=str(idx), values=(
                    row.get('日期', ''),
                    row.get('買家名稱', ''),
                    row.get('商品名稱', ''),
                    row.get('數量', 0),
                    row.get('單價(售)', 0),
                    row.get('分攤手續費', 0),
                    row.get('總淨利', 0),
                    f"{row.get('毛利率', 0)}%"
                ))
        except Exception as e:
            print(f"讀取歷史列表失敗: {e}")


    @thread_safe_file
    def on_sales_edit_select(self, event):
        """ 修正版：從暫存的 sales_edit_df 中使用 .loc 精準讀取 """
        sel = self.tree_sales_edit.selection()
        if not sel: return
        
        item = self.tree_sales_edit.item(sel[0])
        # idx 是我們在 insert 時存入 text 的原始標籤
        idx = int(item['text']) 

        try:
            # --- [關鍵修正 C：改從記憶體讀取，避免重複讀檔導致索引錯亂] ---
            if hasattr(self, 'sales_edit_df'):
                row = self.sales_edit_df.loc[idx] # 使用 loc 根據標籤抓取
            else:
                # 備援機制：如果暫存不存在，才讀檔（但建議盡量使用暫存）
                df = pd.read_excel(FILE_NAME, sheet_name=SHEET_SALES)
                row = df.loc[idx]
            
            # 更新訂單詳情 (現在 row 已經被 ffill 過了，不會有 nan)
            oid = str(row.get('訂單編號', '')).replace("'", "")
            self.var_view_oid.set(oid)
            self.var_view_buyer.set(str(row.get('買家名稱', '')))
            self.var_view_ship.set(str(row.get('寄送方式', '')))
            self.var_view_item.set(str(row.get('商品名稱', '')))
            
            # 格式化稅額與地點
            self.var_view_loc.set(str(row.get('取貨地點', '未提供')))
            self.var_view_tax.set(f"${float(row.get('稅額', 0)):.1f}")
            
            # 更新售後狀態標籤
            current_after_note = str(row.get('扣費項目', '')).strip()
            if current_after_note == "" or current_after_note == "nan":
                self.var_view_after_status.set("目前無售後紀錄")
            else:
                self.var_view_after_status.set(current_after_note)
            
        except Exception as e:
            print(f"讀取詳情失敗: {e}")


    @thread_safe_file
    def save_sales_edit(self):
        """儲存修改並自動重算 (含 Excel 欄位自動修復)"""
        idx = self.var_edit_idx.get()
        if idx < 0: return

        try:
            # 1. 取得新數值
            qty = self.var_edit_qty.get()
            price_sell = self.var_edit_price.get()
            price_cost = self.var_edit_cost.get()
            fee = self.var_edit_fee.get()
            deduct = self.var_edit_deduct.get()

            # 2. 自動重算
            total_sales = qty * price_sell
            total_cost = qty * price_cost
            net_profit = total_sales - total_cost - fee - deduct
            
            margin_pct = 0.0
            if total_sales > 0:
                margin_pct = (net_profit / total_sales) * 100
            
            # 3. 讀取與修復 Excel
            df = pd.read_excel(FILE_NAME, sheet_name='銷售紀錄')
            
            df = df.loc[:, ~df.columns.str.contains('^Unnamed')]


            cols_to_float = ['單價(售)', '單價(進)', '分攤手續費', '總銷售額', '總成本', '總淨利', '毛利率']
            for col in cols_to_float:
                if col not in df.columns:
                    df[col] = 0.0 # 若欄位遺失則補回
                df[col] = df[col].astype(float)
            # ==========================================

            # 更新資料
            df.at[idx, '日期'] = self.var_edit_date.get()
            df.at[idx, '商品名稱'] = self.var_edit_name.get()
            df.at[idx, '數量'] = qty
            df.at[idx, '單價(售)'] = price_sell
            df.at[idx, '單價(進)'] = price_cost
            df.at[idx, '分攤手續費'] = fee
            
            df.at[idx, '總銷售額'] = total_sales
            df.at[idx, '總成本'] = total_cost
            df.at[idx, '總淨利'] = round(net_profit, 2)
            
            # 存數字 (例如 28.7)
            df.at[idx, '毛利率'] = round(margin_pct, 1)

            with pd.ExcelWriter(FILE_NAME, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                try:
                    df_prods = pd.read_excel(FILE_NAME, sheet_name='商品資料')
                except:
                    df_prods = pd.DataFrame()
                
                df.to_excel(writer, sheet_name='銷售紀錄', index=False)
                df_prods.to_excel(writer, sheet_name='商品資料', index=False)

            messagebox.showinfo("成功", "資料已修正!Excel 欄位格式已自動校正。")
            self.load_sales_records_for_edit()
            self.calculate_analysis_data()
            
        except PermissionError:
            messagebox.showerror("錯誤", "Excel 檔案未關閉，無法寫入！")
        except Exception as e:
            messagebox.showerror("錯誤", f"儲存失敗: {str(e)}")

    @thread_safe_file
    def delete_sales_record(self):
        idx = self.var_edit_idx.get()
        if idx < 0: return
        
        confirm = messagebox.askyesno("確認刪除", "確定要刪除這筆銷售紀錄嗎？\n(注意：這不會自動把庫存加回去，請手動調整庫存)")
        if confirm:
            try:
                df = pd.read_excel(FILE_NAME, sheet_name='銷售紀錄')
                df = df.drop(idx) # 刪除該行
                
                # 讀取商品資料以保留
                df_prods = pd.read_excel(FILE_NAME, sheet_name='商品資料')

                with pd.ExcelWriter(FILE_NAME, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='銷售紀錄', index=False)
                    df_prods.to_excel(writer, sheet_name='商品資料', index=False)
                
                messagebox.showinfo("成功", "紀錄已刪除")
                self.load_sales_records_for_edit()
                self.var_edit_idx.set(-1)
                
            except PermissionError:
                messagebox.showerror("錯誤", "Excel 檔案未關閉！")


    def setup_about_tab(self):
        """ 設定分頁：優化排版並修正安全性設定顯示問題 """
        # 建議：如果內容太多，這裡可以考慮加入 Scrollbar，目前先以優化佈局為主
        main_frame = ttk.Frame(self.tab_about, padding=20)
        main_frame.pack(fill="both", expand=True)

        # --- 第一區：介面顯示設定 ---
        font_frame = ttk.LabelFrame(main_frame, text="🎨 介面顯示設定", padding=15)
        font_frame.pack(fill="x", pady=5)

        ttk.Label(font_frame, text="商家名稱:").pack(side="left", padx=5)
        ttk.Entry(font_frame, textvariable=self.var_shop_name, width=20).pack(side="left", padx=5)
        ttk.Button(font_frame, text="💾 儲存設定", command=self.save_system_settings).pack(side="left", padx=5)

        spin = ttk.Spinbox(font_frame, from_=10, to=20, textvariable=self.var_font_size, width=5, command=self.change_font_size)
        spin.pack(side="right", padx=5)
        ttk.Label(font_frame, text="字型大小:").pack(side="right", padx=5)

        # --- 第二區：自訂費率管理 ---
        fee_mgmt_frame = ttk.LabelFrame(main_frame, text="💰 銷售費率管理", padding=15)
        fee_mgmt_frame.pack(fill="x", pady=8)

        # 1. 上排：輸入與操作按鈕 (全部橫向排列)
        input_f = ttk.Frame(fee_mgmt_frame)
        input_f.pack(fill="x", pady=(0, 10))

        ttk.Label(input_f, text="名稱:").pack(side="left")
        self.ent_fee_name = ttk.Entry(input_f, width=12)
        self.ent_fee_name.pack(side="left", padx=5)

        ttk.Label(input_f, text="費率(%):").pack(side="left", padx=5)
        self.ent_fee_val = ttk.Entry(input_f, width=8)
        self.ent_fee_val.pack(side="left", padx=5)

        ttk.Label(input_f, text="固定金額($):").pack(side="left", padx=5)
        self.ent_fee_fixed = ttk.Entry(input_f, width=8)
        self.ent_fee_fixed.insert(0, "0")
        self.ent_fee_fixed.pack(side="left", padx=5)

        # --- 按鈕群組 ---
        ttk.Button(input_f, text="➕ 新增/更新", 
                   command=self.action_add_custom_fee).pack(side="left", padx=10)
        
        # [關鍵修正]：刪除按鈕現在緊跟在新增按鈕右側
        ttk.Button(input_f, text="🗑️刪除", 
                   command=self.action_delete_custom_fee).pack(side="left", padx=2)

        # 2. 下排：列表區域 (Treeview)
        list_f = ttk.Frame(fee_mgmt_frame)
        list_f.pack(fill="x")
        
        self.fee_tree = ttk.Treeview(list_f, columns=("名稱", "百分比", "固定"), show='headings', height=4)
        self.fee_tree.heading("名稱", text="費率名稱")
        self.fee_tree.heading("百分比", text="百分比 (%)")
        self.fee_tree.heading("固定", text="固定金額 ($)")
        self.fee_tree.column("名稱", width=150)
        self.fee_tree.column("百分比", width=80, anchor="center")
        self.fee_tree.column("固定", width=80, anchor="center")
        
        sc_fee = ttk.Scrollbar(list_f, orient="vertical", command=self.fee_tree.yview)
        self.fee_tree.configure(yscrollcommand=sc_fee.set)
        
        self.fee_tree.pack(side="left", fill="x", expand=True)
        sc_fee.pack(side="left", fill="y")




        # --- 第三區：安全性與帳密 (整合在一起節省高度) ---
        security_main_f = ttk.LabelFrame(main_frame, text="🛡️ 安全性與存取控制", padding=15)
        security_main_f.pack(fill="x", pady=5)

        # 帳密變更列
        auth_f = ttk.Frame(security_main_f)
        auth_f.pack(fill="x", pady=5)
        ttk.Label(auth_f, text="帳號:").pack(side="left")
        ttk.Entry(auth_f, textvariable=self.var_new_user, width=12).pack(side="left", padx=5)
        ttk.Label(auth_f, text="密碼:").pack(side="left", padx=5)
        ttk.Entry(auth_f, textvariable=self.var_new_pass, show="*", width=12).pack(side="left", padx=5)
        ttk.Button(auth_f, text="更新憑證", command=self.update_system_auth).pack(side="left", padx=10)

        # 自動登入列 (修正顯示問題)
        bypass_f = ttk.Frame(security_main_f)
        bypass_f.pack(fill="x", pady=(10, 0))
        
        # 確保變數存在
        if not hasattr(self, 'var_auto_login'):
            self.var_auto_login = tk.BooleanVar()
        
        # 讀取目前狀態
        try:
            with open(AUTH_FILE, "r") as f:
                curr_auth = json.load(f)
                self.var_auto_login.set(curr_auth.get("remember", False))
        except:
            self.var_auto_login.set(False)

        self.chk_auto_login = ttk.Checkbutton(bypass_f, text="啟動程式時自動登入 (Bypass Login)", 
                                             variable=self.var_auto_login, command=self.toggle_auto_login)
        self.chk_auto_login.pack(side="left")
        ttk.Label(bypass_f, text="* 勾選後下次啟動將跳過登入視窗", foreground="gray", font=("", 9)).pack(side="left", padx=20)

        #--- 第四區補充：廠商績效評估參數設定 ---
        vendor_kpi_f = ttk.LabelFrame(main_frame, text="📊 廠商績效評估設定", padding=15)
        vendor_kpi_f.pack(fill="x", pady=5)

        # 功能總開關
        chk_master = ttk.Checkbutton(vendor_kpi_f, text="啟用廠商自動績效評估功能", 
                                     variable=self.var_enable_vendor_kpi,
                                     command=self.toggle_vendor_kpi_ui)
        chk_master.pack(anchor="w", pady=(0, 10))

        # 容器
        kpi_grid = ttk.Frame(vendor_kpi_f)
        kpi_grid.pack(fill="x")
        self.kpi_entries = [] # 重新初始化清單
        self.kpi_labels = []  # 建議也存標籤，讓它們一起變灰色更專業

        # 設定欄位權重
        for i in range(4): kpi_grid.columnconfigure(i, weight=1)

        # 輔助函式：快速建立標籤與輸入框並加入追蹤
        def add_kpi_field(row, col, text, var):
            lbl = ttk.Label(kpi_grid, text=text)
            lbl.grid(row=row, column=col*2, sticky="e", padx=5, pady=5)
            ent = ttk.Entry(kpi_grid, textvariable=var, width=8)
            ent.grid(row=row, column=col*2+1, sticky="w")
            self.kpi_labels.append(lbl)
            self.kpi_entries.append(ent)

        # 第一排
        add_kpi_field(0, 0, "品質權重:", self.var_w_quality)
        add_kpi_field(0, 1, "備貨權重:", self.var_w_prep)

        # 第二排
        add_kpi_field(1, 0, "滿足率權重:", self.var_w_fulfill)
        add_kpi_field(1, 1, "運輸權重:", self.var_w_transit)

        # 第三排
        add_kpi_field(2, 0, "備貨標準(天):", self.var_std_prep)
        add_kpi_field(2, 1, "運輸標準(天):", self.var_std_transit)

        # 第四排
        # 標籤
        lbl_ratio = ttk.Label(kpi_grid, text="系統數據佔比:")
        lbl_ratio.grid(row=3, column=0, sticky="e", padx=5, pady=5)
        self.kpi_labels.append(lbl_ratio)
        
        # 輸入框
        ent_ratio = ttk.Entry(kpi_grid, textvariable=self.var_w_system_ratio, width=8)
        ent_ratio.grid(row=3, column=1, sticky="w")
        self.kpi_entries.append(ent_ratio) # 補上這一行，解決它沒被禁用的問題

        # 儲存按鈕
        self.btn_save_kpi_ctrl = ttk.Button(kpi_grid, text="💾 儲存評分參數", command=self.save_vendor_kpi_settings, width=15)
        self.btn_save_kpi_ctrl.grid(row=3, column=2, columnspan=2, sticky="w", padx=20)

        ttk.Label(vendor_kpi_f, text="* 權重建議：四項權重加總應為 1.0。系統數據佔比 0.8 代表人為星等佔 0.2。", 
                  foreground="gray", font=("", 9)).pack(anchor="w", pady=(5,0))
        
        # 立即更新一次介面狀態
        self.toggle_vendor_kpi_ui()

        # --- 第五區：商品欄位顯示 ---
        field_cfg_frame = ttk.LabelFrame(main_frame, text="👁️ 商品資料欄位顯示", padding=15)
        field_cfg_frame.pack(fill="x", pady=5)

        row_f = ttk.Frame(field_cfg_frame)
        row_f.pack(fill="x")

        for label, var in self.show_fields.items():
            ttk.Checkbutton(row_f, text=label, variable=var, 
                            command=self.refresh_product_ui_layout).pack(side="left", padx=10)

        self.refresh_fee_tree()


    @thread_safe_file
    def save_vendor_kpi_settings(self):
        """ 儲存 KPI 參數 (修復型別衝突版) """
        try:
            w_q = self.var_w_quality.get()
            w_p = self.var_w_prep.get()
            w_f = self.var_w_fulfill.get()
            w_t = self.var_w_transit.get()
            s_ratio = self.var_w_system_ratio.get()
            
            # 防呆檢查 (加總是否為 1.0)
            if abs((w_q + w_p + w_f + w_t) - 1.0) > 0.001:
                messagebox.showerror("權重錯誤", "四大權重加總必須等於 1.0")
                return

            # 1. 讀取現有設定
            df_sys = pd.read_excel(FILE_NAME, sheet_name=SHEET_SYS_SETTINGS)
            
            # --- [關鍵修正]：強制轉為字串/物件型態，避免 float64/int64 衝突 ---
            df_sys['參數值'] = df_sys['參數值'].astype(object)

            # 2. 準備更新資料
            new_data = {
                "VENDOR_W_QUALITY": str(w_q),
                "VENDOR_W_PREP": str(w_p),
                "VENDOR_W_FULFILL": str(w_f),
                "VENDOR_W_TRANSIT": str(w_t),
                "VENDOR_STD_PREP": str(self.var_std_prep.get()),
                "VENDOR_STD_TRANSIT": str(self.var_std_transit.get()),
                "VENDOR_W_SYSTEM_RATIO": str(s_ratio),
                "VENDOR_ENABLE_KPI": str(self.var_enable_vendor_kpi.get())
            }

            # 3. 更新或新增
            for key, val in new_data.items():
                if key in df_sys['設定名稱'].values:
                    df_sys.loc[df_sys['設定名稱'] == key, '參數值'] = val
                else:
                    new_row = pd.DataFrame([{"設定名稱": key, "參數值": val}])
                    df_sys = pd.concat([df_sys, new_row], ignore_index=True)

            # 4. 存檔
            if self._universal_save({SHEET_SYS_SETTINGS: df_sys}):
                messagebox.showinfo("成功", "評分參數已成功儲存！")
                # 儲存後立即重新整理廠商頁面的 UI 顯示
                self.refresh_vendor_management_ui()
                
        except Exception as e:
            messagebox.showerror("錯誤", f"儲存失敗: {e}")

    def toggle_vendor_kpi_ui(self):
        """ 根據開關狀態，切換設定頁面的輸入框權限，並通知廠商頁面隱藏/顯示 """
        is_enabled = self.var_enable_vendor_kpi.get()
        state = "normal" if is_enabled else "disabled"
        
        # 1. 處理設定頁面的輸入框與按鈕
        for entry in self.kpi_entries:
            entry.config(state=state)
        self.btn_save_kpi_ctrl.config(state=state)

        # 2. 通知廠商管理頁面進行更新 (稍後在第三階段實作)
        if hasattr(self, 'refresh_vendor_management_ui'):
            self.refresh_vendor_management_ui()
            
        # 同步儲存至 Excel 防止忘記按儲存
        self.save_vendor_kpi_master_switch()

    @thread_safe_file
    def save_vendor_kpi_master_switch(self):
        """ 獨立儲存開關狀態 """
        try:
            df_sys = pd.read_excel(FILE_NAME, sheet_name=SHEET_SYS_SETTINGS)
            val = str(self.var_enable_vendor_kpi.get())
            if "VENDOR_ENABLE_KPI" in df_sys['設定名稱'].values:
                df_sys.loc[df_sys['設定名稱'] == "VENDOR_ENABLE_KPI", '參數值'] = val
                self._universal_save({SHEET_SYS_SETTINGS: df_sys})
        except: pass



    def update_system_auth(self):
        new_u = self.var_new_user.get().strip()
        new_p = self.var_new_pass.get().strip()

        if len(new_u) < 4 or len(new_p) < 6:
            messagebox.showwarning("警告", "帳號至少4位，密碼至少6位")
            return

        if messagebox.askyesno("確認", f"確定要將系統管理員變更為「{new_u}」嗎？\n請務必記住新密碼！"):
            try:

                auth_data = {
                    "user": new_u,
                    "pass": secure_hash(new_p) # 儲存雜湊值，而非明文
                }
                with open(AUTH_FILE, "w") as f:
                    json.dump(auth_data, f)
                
                messagebox.showinfo("成功", "系統存取憑證已更新！\n下次啟動程式時請使用新帳密。")
                self.var_new_user.set("")
                self.var_new_pass.set("")
            except Exception as e:
                messagebox.showerror("失敗", f"更新失敗: {e}")


    def toggle_auto_login(self):
        """ 更新自動登入設定到檔案 """
        try:
            with open(AUTH_FILE, "r") as f:
                data = json.load(f)
            data["remember"] = self.var_auto_login.get()
            with open(AUTH_FILE, "w") as f:
                json.dump(data, f)
            print(f"system: auto-login set to: {data['remember']}")
        except Exception as e:
            messagebox.showerror("錯誤", f"system: failed to save security settings: {e}")



    @thread_safe_file
    def refresh_fee_tree(self):
        """ 從『手續費設定』分頁載入，不再受系統參數干擾 """
        if hasattr(self, 'fee_tree'):
            for i in self.fee_tree.get_children(): self.fee_tree.delete(i)
        
        self.fee_lookup = {}
        fee_options = ["自訂手動輸入"]

        try:
            if not os.path.exists(FILE_NAME): return
            df = pd.read_excel(FILE_NAME, sheet_name=SHEET_FEES)
            df = df.dropna(subset=['設定名稱'])

            for _, row in df.iterrows():
                name = str(row['設定名稱']).strip()
                perc = float(row['費率百分比'])
                fixed = float(row.get('固定金額', 0))
                
                display_str = f"{name} ({perc}% + ${fixed})" if fixed > 0 else f"{name} ({perc}%)"
                self.fee_lookup[display_str] = (perc, fixed)
                fee_options.append(display_str)
                
                if hasattr(self, 'fee_tree'):
                    self.fee_tree.insert("", "end", values=(name, perc, fixed))
            
            if hasattr(self, 'combo_fee_rate'):
                self.combo_fee_rate['values'] = fee_options
        except Exception as e:
            print(f"system: failed to load fee rates: {e}")


    @thread_safe_file
    def action_add_custom_fee(self):
        #""" 新增或更新自訂費率 (修正版：解決 df 變數未定義問題) """
        name = self.ent_fee_name.get().strip()
        raw_val = self.ent_fee_val.get().strip()
        raw_fixed = self.ent_fee_fixed.get().strip() # 取得固定金額

        if not name or not raw_val:
            messagebox.showwarning("警告", "請輸入名稱與費率")
            return

        try:
            # 1. 數值預處理 (過濾 % 號並轉為數字)
            clean_val = raw_val.replace("%", "")
            val = float(clean_val)
            fixed_val = float(raw_fixed) if raw_fixed else 0.0
            
            target_cols = ["設定名稱", "費率百分比", "固定金額"]
            df = None # 【核心修正】：先將 df 初始化為 None

            # 2. 嘗試讀取現有的 Excel 設定
            if os.path.exists(FILE_NAME):
                try:
                    df = pd.read_excel(FILE_NAME, sheet_name=SHEET_FEES)
                    
                    # 檢查並補齊缺失欄位 (防止舊版 Excel 報錯)
                    for col in target_cols:
                        if col not in df.columns:
                            df[col] = 0.0
                except Exception:
                    # 如果分頁不存在或讀取失敗，建立全新的 DataFrame
                    df = pd.DataFrame(columns=target_cols)
            else:
                # 檔案根本不存在
                df = pd.DataFrame(columns=target_cols)

            # 如果到這裡 df 還是 None (極端情況)，補上初始化
            if df is None:
                df = pd.DataFrame(columns=target_cols)

            # 3. 執行新增或更新邏輯
            # 確保內容是乾淨的字串進行比對
            df['設定名稱'] = df['設定名稱'].astype(str).str.strip()
            
            if not df.empty and name in df['設定名稱'].values:
                # 更新現有費率
                df.loc[df['設定名稱'] == name, '費率百分比'] = val
                df.loc[df['設定名稱'] == name, '固定金額'] = fixed_val
            else:
                # 新增一筆
                new_row = pd.DataFrame([[name, val, fixed_val]], columns=target_cols)
                df = pd.concat([df, new_row], ignore_index=True)

            # 4. 調用全能存檔引擎 (我們剛剛統一過的函式)
            # 注意：這裡呼叫的是 _universal_save，它會保護其他所有分頁
            save_success = self._universal_save({SHEET_FEES: df})
            
            if save_success:
                # 5. 刷新介面
                self.refresh_fee_tree()
                
                # 清空輸入框
                self.ent_fee_name.delete(0, tk.END)
                self.ent_fee_val.delete(0, tk.END)
                self.ent_fee_fixed.delete(0, tk.END)
                self.ent_fee_fixed.insert(0, "0") # 重設為 0
                messagebox.showinfo("成功", f"費率「{name}」設定已儲存至 Excel。")

        except ValueError:
            messagebox.showerror("錯誤", "費率與固定金額必須是有效的數字！")
        except Exception as e:
            messagebox.showerror("儲存失敗", f"發生非預期錯誤: {str(e)}")

    def action_delete_custom_fee(self):
        """ 刪除選取費率：增加二次確認視窗 """
        # 1. 檢查是否有選中項目
        sel = self.fee_tree.selection()
        if not sel:
            messagebox.showwarning("提示", "請先點選清單中要刪除的費率項目。")
            return
        
        # 2. 取得選中項目的名稱
        item_data = self.fee_tree.item(sel[0])
        fee_name = item_data['values'][0] # 取得第一欄「費率名稱」

        # 3. 彈出二次確認視窗 (關鍵新增)
        confirm = messagebox.askyesno(
            "⚠️ 確認刪除", 
            f"您確定要刪除費率項目「{fee_name}」嗎？\n\n注意：刪除後，「銷售輸入」頁面的下拉選單將不再出現此選項。"
        )

        # 4. 如果使用者點選「是」(True)，則執行刪除
        if confirm:
            try:
                # 讀取現有設定
                df = pd.read_excel(FILE_NAME, sheet_name=SHEET_FEES)
                
                # 執行過濾：只留下名稱不等於要刪除項目的資料
                df = df[df['設定名稱'].astype(str).str.strip() != str(fee_name).strip()]
                
                # 使用萬用存檔引擎儲存
                if self._universal_save({SHEET_FEES: df}):
                    messagebox.showinfo("成功", f"費率項目「{fee_name}」已成功移除。")
                    # 重新整理介面清單
                    self.refresh_fee_tree()
            except Exception as e:
                messagebox.showerror("錯誤", f"刪除費率失敗: {e}")

    def _save_config_to_excel(self, df_config):
        """ 專門儲存『手續費設定』分頁的輔助函式 (對接萬用引擎) """
        # SHEET_FEES 對應您之前定義的 '手續費設定'
        
        try:
            # 呼叫萬用引擎，傳入字典格式
            success = self._universal_save({SHEET_FEES: df_config})
            
            if success:
                # 更新成功後，重新整理 UI 下拉選單與表格
                self.refresh_fee_tree()
                print("system: fee settings saved successfully.")
            else:
                # 失敗時 (例如 Excel 沒關)，_universal_save 內部已經會跳出 messagebox 提示
                print("system: failed to save fee settings.")
        except Exception as e:
            messagebox.showerror("錯誤", f"system: error occurred while saving settings: {str(e)}")


    def setup_about_us_tab(self):
        """ 建立『關於我/軟體資訊』頁面 """
        # 清空舊頁面，防止重複渲染
        for widget in self.tab_about_us.winfo_children():
            widget.destroy()

        main_frame = ttk.Frame(self.tab_about_us, padding=30)
        main_frame.pack(fill="both", expand=True)

        # --- 頂部：標題與版本 ---
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill="x", pady=(0, 20))
        
        lbl_title = ttk.Label(header_frame, text="蝦皮/網拍智慧進銷存管理系統", font=("微軟正黑體", 20, "bold"))
        lbl_title.pack(anchor="center")
        
        lbl_version = ttk.Label(header_frame, text="Version 5.4 (採購決策優化版)", font=("Consolas", 11), foreground="gray")
        lbl_version.pack(anchor="center")

        # --- 中間：功能簡介與開發者資訊 ---
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill="both", expand=True)

        # 左側：核心功能
        left_box = ttk.LabelFrame(content_frame, text="🚀 系統核心價值", padding=15)
        left_box.pack(side="left", fill="both", expand=True, padx=10)
        
        features = [
            "● 高精度財務運算引擎：導入 Decimal 模組與 ROUND_HALF_UP 會計算法，杜絕浮點數誤差。",
            "● 供應商績效評鑑體系：實作動態 KPI 權重引擎，針對品質、備貨、滿足率及運輸時效進行量化分析。",
            "● 智慧採購需求分析：結合商品銷售速率 (Velocity) 與前置時間 (Lead Time) 建立 ROP 補貨模型。",
            "● 資料防禦機制：實作萬用存檔引擎 (Universal Save Engine) 與標頭繼承算法，確保數據原子性。",
            "● 混合雲資安防護架構：整合 Google Drive API 雲端冗餘備份，並採用 SHA-256 加密認證與金鑰隔離。"
        ]
        for f in features:
            ttk.Label(left_box, text=f, font=("微軟正黑體", 11)).pack(anchor="w", pady=4)

        # 右側：聯絡開發者
        right_box = ttk.LabelFrame(content_frame, text="👨‍💻 開發者資訊", padding=15)
        right_box.pack(side="left", fill="both", expand=True, padx=10)

        ttk.Label(right_box, text="開發者:redmaple", font=("微軟正黑體", 12, "bold")).pack(anchor="w")
        ttk.Label(right_box, text="電子信箱:az062596216@gmail.com", font=("微軟正黑體", 10)).pack(anchor="w", pady=5)
        
        ttk.Separator(right_box, orient="horizontal").pack(fill="x", pady=15)
        
        ttk.Label(right_box, text="📊 檔案存放位置：", font=("微軟正黑體", 11, "bold")).pack(anchor="w")
        db_path = os.path.abspath(FILE_NAME)
        ttk.Label(right_box, text=db_path, foreground="blue", wraplength=300, justify="left").pack(anchor="w", pady=5)
        
        btn_open_folder = ttk.Button(right_box, text="📂 打開所在資料夾", command=lambda: os.startfile(os.path.dirname(db_path)))
        btn_open_folder.pack(anchor="w", pady=10)

        ttk.Label(right_box, text=f"當前設備識別碼: {get_machine_id()}", font=("Consolas", 8), foreground="gray").pack(anchor="w")


        # --- 底部：更新日誌 ---
        log_frame = ttk.LabelFrame(main_frame, text="📝 更新日誌", padding=10)
        log_frame.pack(fill="x", pady=20)
        
        log_text = tk.Text(log_frame, height=8, font=("微軟正黑體", 10), bg="#F8F9FA", relief="flat")
        log_text.pack(fill="x")
        
        logs = (
            "[2026-03-04] V5.1: 實作動態 KPI 權重引擎、導入 Decimal 高精度財務運算、強化數據清洗管線與防呆機制。\n"
            "[2026-03-02] V5.0: 資料儲存架構解耦重構（手續費與系統設定分離）、新增廠商資料批次匯入精靈 (Vendor Wizard)。\n"
            "[2026-02-28] V4.8: 新增 SHA-256 系統啟動認證介面、實作進貨模組『批次複選』與『異步狀態同步』功能。\n"
            "[2026-02-24] V4.6: 建立廠商績效評鑑體系 (質/備/運)、優化跨境物流多節點追蹤與自動化狀態更動紀錄。\n"
            "[2026-02-08] V4.3: 引入採購需求分析模組 (ROP 模型)、優化商品銷售速率計算與回溯算法。\n"
            "[2026-02-05] V4.2: 進貨與銷售端同步支援『內含營業稅 (5%)』回推運算邏輯。\n"
            "[2026-02-02] V4.1: 進貨管理全面單據化，支援批量入庫驗收與加權平均成本 (WAC) 公式。\n"
        )
        log_text.insert("1.0", logs)
        log_text.config(state="disabled") 

        # 版權宣告
        lbl_copyright = ttk.Label(main_frame, text="© 2026 redmaple. All Rights Reserved.", foreground="#CED4DA")
        lbl_copyright.pack(side="bottom", pady=5)
    # ---------------- 邏輯功能區 ----------------


    @thread_safe_file
    def action_cancel_purchase(self):
        """ 標記遺失或取消：支援批次選取刪除 """
        selected_items = self.tree_pur_track.selection()
        if not selected_items:
            messagebox.showwarning("提示", "請先選擇要取消的進貨項目")
            return
        
        count = len(selected_items)
        if not messagebox.askyesno("取消確認", f"確定要【完全刪除】選中的 {count} 筆進貨項目嗎？\n(這將同時移除進貨紀錄與追蹤清單，且無法復原)"):
            return

        try:
            # 1. 讀取資料
            with pd.ExcelFile(FILE_NAME) as xls:
                df_track = pd.read_excel(xls, sheet_name=SHEET_PUR_TRACKING)
                df_hist = pd.read_excel(xls, sheet_name=SHEET_PURCHASES)

            # 2. 蒐集要刪除的資訊
            # 追蹤表使用 Row Index 刪除；歷史表使用單號+品名刪除
            indices_to_drop = []
            for item_id in selected_items:
                item_data = self.tree_pur_track.item(item_id)
                vals = item_data['values']
                
                # 紀錄追蹤表的索引
                indices_to_drop.append(int(item_data['text']))
                
                # 從歷史總表中過濾掉該項目
                pur_id = str(vals[0]).strip()
                p_name = str(vals[2]).strip()
                
                # 這裡執行反向篩選 (只留下不符合條件的)
                df_hist = df_hist[~((df_hist['進貨單號'].astype(str).str.contains(pur_id)) & 
                                    (df_hist['商品名稱'].astype(str).str.strip() == p_name))]

            # 3. 從追蹤表刪除指定索引列
            df_track = df_track.drop(indices_to_drop)

            # 4. 存檔 (保護其他分頁)
            if self._universal_save({
                SHEET_PUR_TRACKING: df_track, 
                SHEET_PURCHASES: df_hist
            }):
                messagebox.showinfo("成功", f"已成功移除 {count} 筆進貨紀錄。")
                self.load_purchase_tracking()
                
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("錯誤", f"取消失敗: {str(e)}")

    
    @thread_safe_file
    def action_confirm_inbound(self):
        """ [強力轉型版] 確認收貨：徹底解決 float64 型別衝突 """
        sel = self.tree_pur_track.selection()
        if not sel: 
            messagebox.showwarning("提示", "請先選擇要入庫的單據項目")
            return
        
        item = self.tree_pur_track.item(sel[0])
        target_pur_id = str(item['values'][0]).replace("'", "").strip()

        if not messagebox.askyesno("整筆入庫確認", f"確認將進貨單號：[{target_pur_id}] \n內的所有商品全部執行「入庫」嗎？"):
            return

        try:
            today_str = datetime.now().strftime("%Y-%m-%d")
            now_full = datetime.now().strftime("%Y-%m-%d %H:%M")

            with pd.ExcelFile(FILE_NAME) as xls:
                df_prods = pd.read_excel(xls, sheet_name=SHEET_PRODUCTS)
                df_tracking = pd.read_excel(xls, sheet_name=SHEET_PUR_TRACKING)
                df_history = pd.read_excel(xls, sheet_name=SHEET_PURCHASES)

            # --- [核心暴力修正：立即切斷 float64 的聯繫] ---
            # 針對可能出錯的文字欄位，讀取後「立刻」強行轉換為純字串，並把 nan 字串清空
            for col in ['入庫日期', '備註', '物流追蹤', '進貨單號']:
                if col in df_history.columns:
                    # 先轉成字串，再把 'nan' 取代掉，這能徹底解決 float64 問題
                    df_history[col] = df_history[col].astype(str).replace('nan', '')
                else:
                    # 萬一欄位不見了，補回空字串列
                    df_history[col] = ""

            # 對於數值欄位，確保它們是數字型態，防止出現奇怪的字串
            for col in ['分攤運費', '海關稅金', '數量', '進貨單價']:
                if col in df_history.columns:
                    df_history[col] = pd.to_numeric(df_history[col], errors='coerce').fillna(0.0)

            # --- 建立臨時匹配 ID ---
            df_tracking['tmp_id'] = df_tracking['進貨單號'].astype(str).str.replace("'", "").str.strip()
            df_history['tmp_id'] = df_history['進貨單號'].astype(str).str.replace("'", "").str.strip()
            
            batch_items = df_tracking[df_tracking['tmp_id'] == target_pur_id].copy()
            
            if batch_items.empty:
                messagebox.showerror("錯誤", f"找不到單號 {target_pur_id}")
                return

            # 處理每一個品項
            for _, row in batch_items.iterrows():
                p_name = str(row['商品名稱']).strip()
                new_qty = float(row.get('數量', 0))
                new_price = float(row.get('進貨單價', 0))
                
                # 清理運費稅金
                ship_fee = pd.to_numeric(row.get('分攤運費', 0), errors='coerce')
                tax_fee = pd.to_numeric(row.get('海關稅金', 0), errors='coerce')
                ship_fee = 0.0 if pd.isna(ship_fee) else ship_fee
                tax_fee = 0.0 if pd.isna(tax_fee) else tax_fee

                # A. 更新商品庫存與成本 (WAC)
                p_mask = df_prods['商品名稱'].astype(str).str.strip() == p_name
                if not df_prods[p_mask].empty:
                    p_idx = df_prods[p_mask].index[0]
                    old_stock = pd.to_numeric(df_prods.at[p_idx, '目前庫存'], errors='coerce')
                    old_cost = pd.to_numeric(df_prods.at[p_idx, '預設成本'], errors='coerce')
                    old_stock = 0.0 if pd.isna(old_stock) else old_stock
                    old_cost = 0.0 if pd.isna(old_cost) else old_cost
                    
                    total_val = (old_stock * old_cost) + (new_qty * new_price) + ship_fee + tax_fee
                    total_qty = old_stock + new_qty
                    
                    if total_qty > 0:
                        new_wac = total_val / total_qty
                        df_prods.at[p_idx, '預設成本'] = round(new_wac, 2)
                        df_prods.at[p_idx, '目前庫存'] = int(total_qty)
                        df_prods.at[p_idx, '最後進貨時間'] = today_str
                        df_prods.at[p_idx, '最後更新時間'] = now_full

                # B. 更新進貨歷史總表
                h_mask = (df_history['tmp_id'] == target_pur_id) & (df_history['商品名稱'] == p_name)
                if not df_history[h_mask].empty:
                    # 這裡現在絕對是字串欄位了，賦值不會再跳 float64 錯誤
                    df_history.loc[h_mask, '入庫日期'] = today_str
                    df_history.loc[h_mask, '備註'] = "已完成入庫"
                    df_history.loc[h_mask, '分攤運費'] = ship_fee
                    df_history.loc[h_mask, '海關稅金'] = tax_fee
                    df_history.loc[h_mask, '物流狀態'] = "已完成入庫"


            # 5. 移除追蹤與清理輔助欄位
            df_tracking_new = df_tracking[df_tracking['tmp_id'] != target_pur_id].copy()
            df_tracking_new.drop(columns=['tmp_id'], inplace=True, errors='ignore')
            df_history.drop(columns=['tmp_id'], inplace=True, errors='ignore')

            # 6. 存檔
            if self._universal_save({
                SHEET_PRODUCTS: df_prods,
                SHEET_PUR_TRACKING: df_tracking_new,
                SHEET_PURCHASES: df_history
            }):
                # --- [關鍵修改：在此處獲取廠商名稱並觸發效能更新] ---
                try:
                    # 從本次入庫的清單中抓取廠商名稱 (假設整筆單號都是同一家廠商)
                    vendor_name = str(batch_items.iloc[0]['供應商']).strip()
                    # 呼叫效能更新引擎
                    self.update_vendor_performance(vendor_name)
                except Exception as ve:
                    print(f"system: failed to update vendor performance: {ve}")
                # ------------------------------------------------

                messagebox.showinfo("成功", f"單號 [{target_pur_id}] 及其商品已全數入庫。")
                self.load_purchase_tracking()
                self.products_df = self.load_products()
                self.update_sales_prod_list()
                self.calculate_analysis_data()

        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("入庫發生嚴重錯誤", str(e))


    @thread_safe_file
    def update_vendor_performance(self, vendor_name):
        """ 核心績效引擎：自動計算前置天數、毀損率、滿足率並更新廠商評等 """
        if not vendor_name or vendor_name == "nan" or vendor_name == "未填":
            return

        try:
            with pd.ExcelFile(FILE_NAME) as xls:
                df_h = pd.read_excel(xls, sheet_name=SHEET_PURCHASES)
                df_v = pd.read_excel(xls, sheet_name=SHEET_VENDORS)

            float_cols = ['平均前置天數', '綜合評等分數']
            for col in float_cols:
                if col in df_v.columns:
                    # 先轉成 numeric 處理 NaN，再強制轉成 float
                    df_v[col] = pd.to_numeric(df_v[col], errors='coerce').astype(float)
            
            # 星等通常是整數，但也建議轉成 float 以免意外
            if '星等' in df_v.columns:
                df_v['星等'] = pd.to_numeric(df_v['星等'], errors='coerce').fillna(5).astype(float)

            # 1. 篩選該廠商的所有已入庫紀錄
            v_mask = (df_h['供應商'].astype(str).str.strip() == vendor_name)
            finished_purchases = df_h[v_mask & (df_h['入庫日期'].notna()) & (df_h['入庫日期'] != "")].copy()

            if finished_purchases.empty:
                return

            # --- A. 計算平均前置天數 (Lead Time) ---
            p_date = pd.to_datetime(finished_purchases['採購日期'], errors='coerce')
            i_date = pd.to_datetime(finished_purchases['入庫日期'], errors='coerce')
            # 過濾無效日期
            valid_dates = (p_date.notna()) & (i_date.notna())
            days_diffs = (i_date[valid_dates] - p_date[valid_dates]).dt.days
            avg_lead_time = round(days_diffs.mean(), 1) if not days_diffs.empty else 0

            # --- B. 計算品質合格率 (Quality Rate) ---
            # 合格率 = (1 - 總瑕疵數 / 總到貨數)
            total_qty = pd.to_numeric(finished_purchases['數量'], errors='coerce').sum()
            total_defects = pd.to_numeric(finished_purchases.get('瑕疵數量', 0), errors='coerce').sum()
            quality_rate = round((1 - (total_defects / total_qty)) * 100, 1) if total_qty > 0 else 100

            # --- C. 計算到貨滿足率 (Fulfillment Rate) ---
            # 滿足率 = (實際到貨數 / 原始預計數)
            original_expected = pd.to_numeric(finished_purchases.get('原始預計數量', total_qty), errors='coerce').sum()
            fulfillment_rate = round((total_qty / original_expected) * 100, 1) if original_expected > 0 else 100




            
            # --- D. 綜合評分演算法 (進化版：數據 80% + 人為 20%) ---
            # 1. 取得人為印象分數 (1-5 轉為 0-100)
            try:
                manual_val = int(self.var_v_manual_adj.get()) * 20 
            except:
                manual_val = 100 # 沒填預設滿分印象

            # 2. 計算系統數據得分 (品質 40%, 時效 30%, 滿足率 10%)
            time_score = max(100 - (avg_lead_time * 5), 0)
            system_data_score = (quality_rate * 0.4) + (time_score * 0.3) + (fulfillment_rate * 0.1)

            # 3. 最終混合評分
            final_score = (system_data_score * 0.8) + (manual_val * 0.2)

            # 轉換為星等 (5星制)
            star = 1
            if final_score >= 90: star = 5
            elif final_score >= 80: star = 4
            elif final_score >= 70: star = 3
            elif final_score >= 60: star = 2

            # 2. 更新回廠商分頁
            if vendor_name in df_v['廠商名稱'].astype(str).str.strip().values:
                idx = df_v[df_v['廠商名稱'].astype(str).str.strip() == vendor_name].index[0]
                
                # 強制轉為數值型態後再賦值
                df_v.at[idx, '平均前置天數'] = float(avg_lead_time)
                df_v.at[idx, '綜合評等分數'] = float(round(final_score, 1))
                df_v.at[idx, '星等'] = int(star)
                
                # 強制轉為數值型態後再賦值
                df_v.at[idx, '總到貨率'] = f"{fulfillment_rate}%"
                df_v.at[idx, '總合格率'] = f"{quality_rate}%"
                df_v.at[idx, '最後更新'] = datetime.now().strftime("%Y-%m-%d %H:%M")
                # 存檔 (保護其他分頁)

                for col in ['平均前置天數', '綜合評等分數', '星等']:
                    df_v[col] = pd.to_numeric(df_v[col], errors='coerce').fillna(0)

                self._universal_save({SHEET_VENDORS: df_v})

                if self._universal_save({SHEET_VENDORS: df_v}):
                    print(f"system: vendor performance updated (Score: {final_score})")

        except Exception as e:
            print(f"system: failed to update vendor analysis: {e}")
            import traceback
            traceback.print_exc()


    def update_pur_prod_list(self):
        """ 初始化/重新載入進貨商品清單 """
        if not hasattr(self, 'list_pur_prod'): return
        self.list_pur_prod.delete(0, tk.END)
        
        if not self.products_df.empty:
            for _, row in self.products_df.iterrows():
                p_name = str(row['商品名稱'])
                raw_tag = row.get('分類Tag', '')
                display_tag = str(raw_tag).strip() if pd.notna(raw_tag) else ""
                
                full_display_name = f"[{display_tag}] {p_name}" if display_tag else p_name
                self.list_pur_prod.insert(tk.END, full_display_name)

    def on_pur_prod_select(self, event):
        """ 當進貨選中商品時，自動帶入目前的成本作為參考 """
        selected_name = self.var_pur_sel_name.get()
        
        # 根據選中的名稱去找原始資料
        record = self.products_df[self.products_df['商品名稱'] == selected_name]
        if not record.empty:
            current_cost = record.iloc[0]['預設成本']
            self.var_pur_sel_cost.set(current_cost)
            
            # 可選：選中後自動刷新 values 回全部清單，方便下次搜尋
            self.combo_pur_prod['values'] = self.products_df['商品名稱'].tolist()


    def add_to_pur_cart(self):
        """ 加入商品到進貨購物車 (修正為總額直乘稅率邏輯) """
        name = self.var_pur_sel_name.get()
        qty = self.var_pur_sel_qty.get()
        cost = self.var_pur_sel_cost.get() 
        
        if not name or qty <= 0: 
            messagebox.showwarning("提示", "請先選擇商品並輸入正確數量")
            return

        # 含稅總額 (小計)
        total_inclusive = qty * cost
        
        if self.var_pur_tax_enabled.get():
            tax = round(total_inclusive * 0.05, 2)
        else:
            tax = 0.0

        self.pur_cart_data.append({
            "name": name, "qty": qty, "cost": cost, "tax": tax, "total": total_inclusive
        })
        
        # 這裡的 values 順序必須跟上面的 pur_cols 一致
        self.tree_pur_cart.insert("", "end", values=(name, qty, cost, tax, total_inclusive))
        
        # 加入後自動清空輸入框以便下一筆
        self.var_pur_sel_name.set("")
        self.var_pur_sel_qty.set(1)
        self.var_pur_sel_cost.set(0.0)
        self.update_pur_cart_total() 
        self.ent_pur_search.delete(0, tk.END) # 清空搜尋框
        self.update_pur_prod_list() # 恢復完整列表


    def remove_from_pur_cart(self):
        """ 移除進貨購物車中的選定單項商品 (修正報錯) """
        sel = self.tree_pur_cart.selection()
        if not sel:
            messagebox.showwarning("提示", "請先點選要移除的商品項目")
            return
    
        for item in sel:
            idx = self.tree_pur_cart.index(item)
            if 0 <= idx < len(self.pur_cart_data):
                del self.pur_cart_data[idx]
            self.tree_pur_cart.delete(item)
    
    # [修正點]：呼叫統一更新函式
        self.update_pur_cart_total()

# [新增這個輔助函式]：統一計算並更新介面
    def update_pur_cart_total(self):
        """ 使用 Decimal 重新計算進貨購物車總計 """
        total_sum = Decimal("0.00")
        for item in self.pur_cart_data:
            total_sum += Decimal(str(item['total']))
        
        if hasattr(self, 'lbl_pur_total'):
            # 格式化輸出：加上千分位與固定兩位小數
            self.lbl_pur_total.config(text=f"本次進貨總額: ${float(total_sum):,.2f}")


    @thread_safe_file
    def submit_purchase(self):
        """ 提交進貨：更新庫存、更新成本、記錄進貨單 (V4.7 萬用引擎版) """
        name = self.var_pur_name.get().strip()
        qty = self.var_pur_qty.get()
        cost = self.var_pur_cost.get()
        supplier = self.var_pur_supplier.get().strip()
        logistics = self.var_pur_logistics.get().strip()
        date_str = self.var_pur_date.get()

        if not name or qty <= 0:
            messagebox.showwarning("警告", "請填寫正確商品與數量")
            return

        # 生成編號: I + YYYYMMDDHHMMSS
        pur_id = "I" + datetime.now().strftime("%Y%m%d%H%M%S")

        try:
            # 1. 讀取需要更動的分頁
            with pd.ExcelFile(FILE_NAME) as xls:
                df_prods = pd.read_excel(xls, sheet_name=SHEET_PRODUCTS)
                df_pur = pd.read_excel(xls, sheet_name=SHEET_PURCHASES)

            # 2. 更新商品庫存與成本 (WAC 邏輯)
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
            if name in df_prods['商品名稱'].values:
                idx = df_prods[df_prods['商品名稱'] == name].index[0]
                
                # 更新數據
                df_prods.at[idx, '目前庫存'] += qty
                df_prods.at[idx, '預設成本'] = cost 
                df_prods.at[idx, '最後更新時間'] = now_str
                df_prods.at[idx, '最後進貨時間'] = now_str
            else:
                messagebox.showerror("錯誤", f"找不到商品「{name}」，請先到商品管理新增。")
                return

            # 3. 建立進貨紀錄 DataFrame
            new_pur_row = pd.DataFrame([{
                "進貨單號": f"'{pur_id}", # 強制字串
                "進貨日期": date_str,
                "供應商": supplier,
                "物流追蹤編號": logistics,
                "商品名稱": name,
                "數量": qty,
                "進貨單價": cost,
                "進貨總額": qty * cost,
                "備註": "直接入庫"
            }])
            df_pur = pd.concat([df_pur, new_pur_row], ignore_index=True)

            # 4. 【核心改變】：呼叫萬用引擎一次性寫回所有變動
            # 這裡我們傳入一個字典，包含這次要更新的兩個 DataFrame
            # 引擎會自動保護 SHEET_FEES, SHEET_SYS_SETTINGS 等其他所有分頁
            save_dict = {
                SHEET_PRODUCTS: df_prods,
                SHEET_PURCHASES: df_pur
            }

            if self._universal_save(save_dict):
                messagebox.showinfo("成功", f"進貨單 {pur_id} 已入庫！\n庫存已自動增加 {qty}。")
                
                # 清除輸入並刷新介面
                self.var_pur_qty.set(1)
                self.var_pur_cost.set(0.0)
                self.var_pur_logistics.set("")
                self.load_purchase_data()
                self.products_df = df_prods 
                self.update_sales_prod_list() 
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("錯誤", f"進貨作業失敗: {str(e)}")

    @thread_safe_file
    def load_purchase_data(self):
        """ 載入最近進貨清單 """
        for i in self.tree_purchase.get_children(): self.tree_purchase.delete(i)
        try:
            df = pd.read_excel(FILE_NAME, sheet_name=SHEET_PURCHASES)
            # 只顯示最近 20 筆
            for _, row in df.tail(20).iloc[::-1].iterrows():
                self.tree_purchase.insert("", "end", values=(
                    str(row['進貨單號']).replace("'", ""),
                    row['進貨日期'],
                    row['供應商'],
                    row['商品名稱'],
                    row['數量'],
                    row['物流追蹤編號']
                ))
        except: pass

    @thread_safe_file
    def action_update_pur_logistics(self):
        """ 
        物流維護 V6.9 (跨單偵測強化版)
        1. 針對多筆選取，在視窗最上方顯示紅字警告。
        2. [核心新增]：偵測是否選到不同訂單號，若是則顯示橘色警告。
        """
        selected_items = self.tree_pur_track.selection()
        if not selected_items:
            messagebox.showwarning("提示", "請先選擇商品項目")
            return

        # 讀取完整資料表
        try:
            df_full = pd.read_excel(FILE_NAME, sheet_name=SHEET_PUR_TRACKING)
        except Exception as e:
            messagebox.showerror("錯誤", f"讀取追蹤表失敗: {e}")
            return

        batch_list = []
        unique_order_ids = set() # 用於儲存不重複的單號

        for item_id in selected_items:
            item_data = self.tree_pur_track.item(item_id)
            vals = item_data['values']
            df_idx = int(item_data['text'])
            
            p_id = str(vals[0]).replace("'", "").strip() # 取得單號
            unique_order_ids.add(p_id) # 加入集合

            # 抓取並清理欄位
            raw_logi = str(df_full.at[df_idx, '物流追蹤']).strip() if '物流追蹤' in df_full.columns else ""
            raw_remark = str(df_full.at[df_idx, '備註']).strip() if '備註' in df_full.columns else ""

            def clean_str(s):
                if s.lower() in ["nan", "none", "nat", ""]: return ""
                return s.lstrip("'")

            batch_list.append({
                "df_idx": df_idx, 
                "p_name": str(vals[2]).strip(),
                "pur_id": p_id,
                "qty": int(vals[3]),
                "status": str(vals[7]),
                "logi_id": clean_str(raw_logi),
                "remark": clean_str(raw_remark)
            })

        is_batch = len(batch_list) > 1
        is_mixed_orders = len(unique_order_ids) > 1 # 判斷是否包含複數訂單
        
        win = tk.Toplevel(self.root)
        win.title("📦 物流批次維護" if is_batch else "物流維護")
        win.geometry("500x720") # 稍微增加高度
        win.grab_set()

        # --- [頂部提示區塊：動態警告邏輯] ---
        header_warning_f = ttk.Frame(win, padding=10)
        header_warning_f.pack(fill="x")

        if is_batch:
            # 1. 基礎複選警告
            lbl_warn = ttk.Label(header_warning_f, 
                                 text=f"⚠️ 注意：目前已選取 {len(batch_list)} 筆商品", 
                                 foreground="red", 
                                 font=("微軟正黑體", 12, "bold"))
            lbl_warn.pack(anchor="center")
            
            # 2. [新增] 跨單號警告 (Orange Alert)
            if is_mixed_orders:
                mixed_warn_f = ttk.Frame(header_warning_f)
                mixed_warn_f.pack(pady=5)
                
                ttk.Label(mixed_warn_f, 
                          text="❗ 偵測到選中項目包含不同的進貨單號！", 
                          foreground="#FF8C00", # 橘色
                          font=("微軟正黑體", 10, "bold")).pack()
                
                # 顯示受影響的單號簡述，方便確認
                ids_str = ", ".join(list(unique_order_ids)[:3]) # 只列前三筆
                if len(unique_order_ids) > 3: ids_str += "..."
                ttk.Label(mixed_warn_f, 
                          text=f"(涉及單號: {ids_str})", 
                          foreground="#666", 
                          font=("微軟正黑體", 9)).pack()

            lbl_sub_warn = ttk.Label(header_warning_f, 
                                     text="批量儲存後，所選項目的狀態與備註將被『完全覆蓋』。", 
                                     foreground="#888", 
                                     font=("微軟正黑體", 9))
            lbl_sub_warn.pack(anchor="center", pady=(5,0))
        else:
            ttk.Label(header_warning_f, text=f"商品：{batch_list[0]['p_name']}", 
                      font=("微軟正黑體", 11, "bold")).pack(anchor="w")

        ttk.Separator(win, orient="horizontal").pack(fill="x", padx=10)

        # --- [其餘介面部分 (數量、狀態、備註) 保持不變] ---
        body = ttk.Frame(win, padding=20)
        body.pack(fill="both", expand=True)

        # 1. 數量與瑕疵 (僅單選)
        var_actual_qty = tk.IntVar(value=batch_list[0]['qty'])
        var_defects = tk.IntVar(value=0)
        if not is_batch:
            qty_f = ttk.LabelFrame(body, text="📦 數量與品質驗收", padding=10)
            qty_f.pack(fill="x", pady=(0, 10))
            ttk.Label(qty_f, text="實際收到數量:").pack(anchor="w")
            ttk.Entry(qty_f, textvariable=var_actual_qty).pack(fill="x", pady=2)
            ttk.Label(qty_f, text="瑕疵/損壞數量:").pack(anchor="w")
            ttk.Entry(qty_f, textvariable=var_defects).pack(fill="x", pady=2)

        # 2. 狀態與單號
        logi_f = ttk.LabelFrame(body, text="🚚 物流狀態更新", padding=10)
        logi_f.pack(fill="x", pady=10)
        ttk.Label(logi_f, text="變更階段為:").pack(anchor="w")
        var_status = tk.StringVar(value=batch_list[0]['status'])
        status_cb = ttk.Combobox(logi_f, textvariable=var_status, state="readonly")
        status_cb['values'] = ("待出貨", "廠商已發貨", "貨到集運倉", "集運倉已發貨", "抵達台灣海關", "國內配送中")
        status_cb.pack(fill="x", pady=5)

        ttk.Label(logi_f, text="物流單號:").pack(anchor="w", pady=(10,0))
        var_logi_id = tk.StringVar(value=batch_list[0]['logi_id'])
        ttk.Entry(logi_f, textvariable=var_logi_id).pack(fill="x", pady=5)

        # 3. 備註 (P欄)
        rem_f = ttk.LabelFrame(body, text="📝 備註事項 (P欄)", padding=10)
        rem_f.pack(fill="x", pady=10)
        var_custom_remark = tk.StringVar(value=batch_list[0]['remark'])
        ttk.Entry(rem_f, textvariable=var_custom_remark).pack(fill="x", pady=5)

        # --- [儲存邏輯 保持不變] ---
        def perform_save():
            try:
                today_str = datetime.now().strftime("%Y-%m-%d")
                with pd.ExcelFile(FILE_NAME) as xls:
                    df_track = pd.read_excel(xls, sheet_name=SHEET_PUR_TRACKING)
                    df_hist = pd.read_excel(xls, sheet_name=SHEET_PURCHASES)

                # 清理文字欄位
                text_cols = ['物流狀態', '物流追蹤', '備註', '進貨單號', '商品名稱', '時間_廠商出貨', 
                             '時間_抵達集運倉', '時間_集運倉出貨', '時間_抵達台灣海關', '時間_國內配送中']
                for df in [df_track, df_hist]:
                    for col in text_cols:
                        if col not in df.columns: df[col] = ""
                    df[text_cols] = df[text_cols].fillna("").astype(str).replace(['nan', 'NaN', 'None'], '')

                for item in batch_list:
                    t_idx = item['df_idx']
                    df_track.at[t_idx, '物流狀態'] = var_status.get()
                    if var_logi_id.get().strip():
                        df_track.at[t_idx, '物流追蹤'] = f"'{var_logi_id.get().strip()}"
                    df_track.at[t_idx, '備註'] = var_custom_remark.get().strip()

                    if not is_batch:
                        df_track.at[t_idx, '數量'] = var_actual_qty.get()
                        df_track.at[t_idx, '瑕疵數量'] = var_defects.get()
                        u_price = pd.to_numeric(df_track.at[t_idx, '進貨單價'], errors='coerce')
                        df_track.at[t_idx, '進貨總額'] = var_actual_qty.get() * u_price

                    time_map = {"廠商已發貨": "時間_廠商出貨", "貨到集運倉": "時間_抵達集運倉", 
                                "集運倉已發貨": "時間_集運倉出貨", "抵達台灣海關": "時間_抵達台灣海關", "國內配送中": "時間_國內配送中"}
                    col_name = time_map.get(var_status.get())
                    if col_name: df_track.at[t_idx, col_name] = today_str

                    h_mask = (df_hist['進貨單號'].str.replace("'", "").str.strip() == item['pur_id']) & \
                             (df_hist['商品名稱'].str.strip() == item['p_name'])
                    if not df_hist[h_mask].empty:
                        df_hist.loc[h_mask, '物流狀態'] = var_status.get()
                        df_hist.loc[h_mask, '物流追蹤'] = df_track.at[t_idx, '物流追蹤']
                        df_hist.loc[h_mask, '備註'] = var_custom_remark.get().strip()
                        if col_name: df_hist.loc[h_mask, col_name] = today_str

                if self._universal_save({SHEET_PUR_TRACKING: df_track, SHEET_PURCHASES: df_hist}):
                    messagebox.showinfo("成功", "物流資訊更新成功")
                    self.load_purchase_tracking()
                    win.destroy()
            except Exception as e:
                messagebox.showerror("錯誤", f"存檔失敗: {e}")

        ttk.Button(body, text="🚀 執行更新並存檔", command=perform_save, style="Accent.TButton").pack(pady=20, fill="x")

    @thread_safe_file
    def _get_full_order_info(self, df, order_id):
        """ 強化版：從同一編號中找出『任何一列』含有資料的內容，確保不會因刪除首行而遺失資訊 """
        clean_id = str(order_id).replace("'", "").strip()
        # 找出該訂單的所有列
        subset = df[df['訂單編號'].astype(str).str.contains(clean_id)]
        
        if subset.empty: return {}

        # 定義需要找尋的標頭欄位
        header_cols = ['日期', '買家名稱', '交易平台', '寄送方式', '取貨地點', '扣費項目']
        result = {}

        for col in header_cols:
            if col in subset.columns:
                # 找尋該欄位中第一個不是空的、不是 NaN 的值
                valid_rows = subset[subset[col].notna() & (subset[col].astype(str).str.strip() != "")]
                if not valid_rows.empty:
                    result[col] = valid_rows.iloc[0][col]
                else:
                    result[col] = "" # 若真的都沒資料則留空
        return result
    
    @thread_safe_file
    def action_track_return_item(self):
        """ 退貨單一商品：若刪除的是標頭行，自動將資訊傳承給下一筆商品 """
        from tkinter import simpledialog
        sel = self.tree_track.selection()
        if not sel: return
        
        item = self.tree_track.item(sel[0])
        idx = int(item['text']) # Excel 原始行號
        order_id = str(item['values'][0]).replace("'", "").strip()
        prod_name = str(item['values'][4])

        reason = simpledialog.askstring("退貨", f"商品: {prod_name}\n請輸入退貨原因:", parent=self.root)
        if reason is None: return

        try:
            df_track = pd.read_excel(FILE_NAME, sheet_name=SHEET_TRACKING)
            df_track['訂單編號'] = df_track['訂單編號'].astype(str).str.replace(r'^\'', '', regex=True).str.replace(r'\.0$', '', regex=True).str.strip()

            # A. 備份要移走的這一行
            row_to_move = df_track.loc[[idx]].copy()
            
            # B. 檢查這行是否帶有標頭資訊
            has_info = pd.notna(df_track.at[idx, '日期']) and str(df_track.at[idx, '日期']) != ""
            
            if has_info:
                # 找出同訂單的其他商品
                others = df_track[(df_track['訂單編號'] == order_id) & (df_track.index != idx)].index.tolist()
                if others:
                    # 將標頭資訊傳承給下一個商品 (補位)
                    new_header_idx = others[0]
                    header_cols = ['日期', '買家名稱', '交易平台', '寄送方式', '取貨地點', '扣費項目']
                    for col in header_cols:
                        df_track.at[new_header_idx, col] = df_track.at[idx, col]

            # C. 執行移動
            df_track.drop(idx, inplace=True)
            try: df_returns = pd.read_excel(FILE_NAME, sheet_name=SHEET_RETURNS)
            except: df_returns = pd.DataFrame()
            
            # 存入退貨區前，確保退貨區的那一行資訊是完整的 (方便查帳)
            full_info = self._get_full_order_info(df_track, order_id) # 這裡要稍微注意邏輯順序
            for col, val in full_info.items(): row_to_move[col] = val
            row_to_move['備註'] = reason

            df_returns = pd.concat([df_returns, row_to_move], ignore_index=True)

            if self._universal_save({SHEET_TRACKING: df_track, SHEET_RETURNS: df_returns}):
                messagebox.showinfo("成功", f"商品「{prod_name}」已移至退貨，資料已自動補位。")
                self.load_tracking_data(); self.load_returns_data()
        except Exception as e: 
            messagebox.showerror("錯誤", str(e))


    @thread_safe_file
    def action_track_complete_order(self):
        """ 
        完成訂單 V5.1 (防斷鏈修正版):
        解決留白資料排序後掉到最底部的問題
        """
        sel = self.tree_track.selection()
        if not sel: return
        item = self.tree_track.item(sel[0])
        order_id = str(item['values'][0]).replace("'", "").strip()

        if not messagebox.askyesno("結案確認", f"確定訂單 [{order_id}] 已完成？"): return

        try:
            # 1. 讀取追蹤與銷售紀錄
            df_track = pd.read_excel(FILE_NAME, sheet_name=SHEET_TRACKING)
            df_track['訂單編號'] = df_track['訂單編號'].astype(str).str.replace(r'^\'', '', regex=True).str.replace(r'\.0$', '', regex=True).str.strip()
            
            try: 
                df_sales = pd.read_excel(FILE_NAME, sheet_name=SHEET_SALES)
            except: 
                df_sales = pd.DataFrame()

            # --- [核心修正步驟 A：處理舊有銷售紀錄，防止排序斷鏈] ---
            if not df_sales.empty:
                # 先將所有空字串轉為真正的 NaN，這樣 ffill 才會生效
                df_sales = df_sales.replace(r'^\s*$', pd.NA, regex=True)
                
                # 為了防止排序時空日期掉到最後，我們先按『訂單編號』分組，把遺失的日期與買家補回來
                header_cols = ['日期', '買家名稱', '交易平台', '寄送方式', '取貨地點', '扣費項目']
                # 暫時清理編號以便匹配
                df_sales['tmp_id'] = df_sales['訂單編號'].astype(str).str.replace("'", "").str.strip()
                # 分組向下填充：這確保每一行商品都暫時擁有正確的日期
                df_sales[header_cols] = df_sales.groupby('tmp_id', group_keys=False)[header_cols].apply(lambda x: x.ffill().bfill())

            # 2. 獲取本次要結案的訂單資訊
            full_info = self._get_full_order_info(df_track, order_id)
            mask = df_track['訂單編號'] == order_id
            rows_to_finish = df_track[mask].copy()

            # 確保這次結案的資料也是完整的
            for col, val in full_info.items():
                rows_to_finish[col] = val

            # 3. 合併新舊資料
            df_sales_combined = pd.concat([df_sales, rows_to_finish], ignore_index=True)

            # --- [核心修正步驟 B：執行精準排序] ---
            # 統一轉換日期格式，這時因為 A 步驟補過位，日期不會是空的
            df_sales_combined['tmp_date'] = pd.to_datetime(df_sales_combined['日期'], errors='coerce')
            # 將訂單編號轉為清潔字串進行排序
            df_sales_combined['tmp_clean_id'] = df_sales_combined['訂單編號'].astype(str).str.replace("'", "").str.strip()
            
            # 排序：日期(新到舊) -> 編號(大到小)
            df_sales_combined = df_sales_combined.sort_values(by=['tmp_date', 'tmp_clean_id'], ascending=[False, False])
            df_sales_combined = df_sales_combined.reset_index(drop=True)

            # --- [核心修正步驟 C：重新執行視覺去重 (美化)] ---
            prev_id = None
            for i in range(len(df_sales_combined)):
                curr_id = df_sales_combined.at[i, 'tmp_clean_id']
                if curr_id == prev_id:
                    # 同一單的後續列，清空重複資訊
                    for col in header_cols:
                        df_sales_combined.at[i, col] = ""
                else:
                    # 新訂單的第一列，保留資訊
                    prev_id = curr_id
            
            # 移除所有臨時輔助欄位
            drop_cols = ['tmp_date', 'tmp_clean_id', 'tmp_id']
            df_sales_combined = df_sales_combined.drop(columns=[c for c in drop_cols if c in df_sales_combined.columns])

            # 4. 執行萬用存檔
            success = self._universal_save({
                SHEET_TRACKING: df_track[~mask], 
                SHEET_SALES: df_sales_combined
            })
        
            if success:
                messagebox.showinfo("成功", f"訂單 {order_id} 結案成功！全表已重新自動校準。")
                print(f"system: order {order_id} marked as completed and sales data re-aligned.")
                self.load_tracking_data()
                self.calculate_analysis_data()
                
        except Exception as e:   
            import traceback
            traceback.print_exc()
            messagebox.showerror("錯誤", f"結案失敗: {str(e)}")



    @thread_safe_file
    def _universal_save(self, updates_dict):
        """ 
        終極防禦版萬用存檔引擎：
        1. 執行緒鎖 (Thread Lock)：防止併發衝突。
        2. 原子性寫入 (Atomic Write)：使用臨時檔置換，防止寫入中斷導致檔案毀損。
        3. 資料校準 (Data Scrubbing)：消滅 nan,保護 ID 格式。
        """
    
        # 拆解路徑，確保 temp_ 只加在「檔名」前面，而不是整個路徑前面
        directory = os.path.dirname(FILE_NAME)
        base_name = os.path.basename(FILE_NAME)
        
        temp_file = os.path.join(directory, "temp_" + base_name)
        bak_file = FILE_NAME + ".bak"
            
        try:
            all_data = {}
            # 1. 讀取現有分頁
            if os.path.exists(FILE_NAME):
                with pd.ExcelFile(FILE_NAME) as xls:
                    for sn in xls.sheet_names:
                        all_data[sn] = pd.read_excel(xls, sheet_name=sn)
            
            # 2. 更新資料並進行保護
            for sheet_name, df in updates_dict.items():
                # 數據完整性保護：防止意外存入空表
                if sheet_name in all_data and not all_data[sheet_name].empty:
                    if df is None or df.empty:
                        print(f"⚠️ 攔截到分頁 [{sheet_name}] 的洗除嘗試。")
                        continue 
                all_data[sheet_name] = df

            # 3. 核心數據清洗 (您原本的高階邏輯)
            text_protection_cols = ['訂單編號', '進貨單號', '物流追蹤', '商品編號', '廠商名稱', '商店名']
            for sn, df in all_data.items():
                if df is None or df.empty: continue
                df = df.fillna("")
                for col in df.columns:
                    if col in text_protection_cols:
                        def clean_logic(x):
                            s = str(x).strip()
                            if s.lower() in ['nan', 'none', '', 'nat']: return ""
                            if s.endswith('.0'): s = s[:-2]
                            s = s.lstrip("'")
                            if col in ['訂單編號', '進貨單號', '物流追蹤']: return f"'{s}"
                            return s
                        df[col] = df[col].apply(clean_logic)
                
                if sn == SHEET_VENDORS and '廠商名稱' in df.columns:
                    df = df[df['廠商名稱'].astype(str).str.lower() != "nan"]
                    df = df[df['廠商名稱'].astype(str).str.strip() != ""]
                all_data[sn] = df

            # --- 4. 關鍵修正：先寫入「臨時檔案」 ---
            standard_order = [SHEET_PRODUCTS, SHEET_SALES, SHEET_TRACKING, SHEET_PURCHASES, SHEET_PUR_TRACKING, SHEET_RETURNS, SHEET_FEES, SHEET_SYS_SETTINGS, SHEET_VENDORS]
            
            with pd.ExcelWriter(temp_file, engine='openpyxl') as writer:
                for sn in standard_order:
                    if sn in all_data:
                        all_data[sn].to_excel(writer, sheet_name=sn, index=False)
                for sn, df in all_data.items():
                    if sn not in standard_order:
                        df.to_excel(writer, sheet_name=sn, index=False)

            # --- 5. 檔案原子置換 (The Atomic Swap) ---
            # 走到這裡，代表臨時檔寫入成功了，現在才動原始檔案
            if os.path.exists(FILE_NAME):
                # 建立 .bak 備份（以防萬一）
                if os.path.exists(bak_file):
                    os.remove(bak_file)
                os.rename(FILE_NAME, bak_file)
            
            # 將臨時檔改名為正式檔 (此動作在 OS 層級是極快的，幾乎不會中斷)
            os.rename(temp_file, FILE_NAME)
            
            return True

        except PermissionError:
            messagebox.showerror("存檔失敗", "Excel 檔案正被其他程式開啟中，請先關閉 Excel!")
            if os.path.exists(temp_file): os.remove(temp_file)
            return False
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("嚴重錯誤", f"存檔引擎故障: {str(e)}")
            if os.path.exists(temp_file): os.remove(temp_file)
            return False
    

    def load_existing_tags(self, event=None):
        """ 從目前的商品資料中抓取不重複的分類 """
        if not self.products_df.empty:
            tags = sorted([str(t) for t in self.products_df["分類Tag"].dropna().unique() if str(t).strip() != ""])
            # 同步更新兩個下拉選單
            if hasattr(self, 'combo_add_tag'):
                self.combo_add_tag['values'] = tags
            if hasattr(self, 'combo_upd_tag'):
                self.combo_upd_tag['values'] = tags


    def toggle_cust_info(self):
        state = "normal" if self.var_enable_cust.get() else "disabled"
        self.entry_cust_name.config(state=state)
        self.combo_platform.config(state="readonly" if state == "normal" else "disabled")
        self.combo_ship.config(state="readonly" if state == "normal" else "disabled")
        self.combo_loc.config(state=state)

    def filter_cities(self, event):
        typed = self.var_cust_loc.get()
        if typed == '': self.combo_loc['values'] = TAIWAN_CITIES
        else: self.combo_loc['values'] = [i for i in TAIWAN_CITIES if typed in i]

    def on_ship_method_change(self, event):
        method = self.var_ship_method.get()
        if "面交" in method: 
            self.var_cust_loc.set("面交")
        elif self.var_cust_loc.get() == "面交": 
            self.var_cust_loc.set("")

    def update_sales_prod_list(self, event=None):
        """ 
        優化搜尋：
        1. 自動過濾前後空白
        2. 支援多關鍵字搜尋 (以空格隔開)
        """
        # 取得輸入並轉小寫，使用 split() 自動切分關鍵字並過濾掉多餘空格
        search_raw = self.var_search.get().lower()
        search_keywords = search_raw.split() # 這會將 "Arctic  " 轉為 ["arctic"]

        self.listbox_sales.delete(0, tk.END)
        
        if not self.products_df.empty:
            for index, row in self.products_df.iterrows():
                p_name = str(row['商品名稱']).lower()
                p_tag = str(row['分類Tag']).lower() if pd.notna(row['分類Tag']) else ""
                sku = str(row.get('商品編號', '')).lower()
                
                # 如果沒有輸入任何內容，則顯示全部
                if not search_keywords:
                    display_str = f"[{row['分類Tag']}] {row['商品名稱']} (庫存: {row['目前庫存']})"
                    self.listbox_sales.insert(tk.END, display_str)
                    continue

                # --- 核心優化：檢查是否符合「所有」關鍵字 ---
                # 這樣搜尋 "P12 白" 就能過濾出 "Arctic P12 (白框)"
                match_all = True
                for kw in search_keywords:
                    if not (kw in p_name or kw in p_tag or kw in sku):
                        match_all = False
                        break
                
                if match_all:
                    display_str = f"[{row['分類Tag']}] {row['商品名稱']} (庫存: {row['目前庫存']})"
                    self.listbox_sales.insert(tk.END, display_str)

    def on_sales_prod_select(self, event):
        selection = self.listbox_sales.curselection()
        if selection:
            display_str = self.listbox_sales.get(selection[0])
            # 解析名稱：拿最後一個 "]" 之後的文字，並切掉後面的 "(庫存:..."
            try:
                temp = display_str.rsplit(" (庫存:", 1)[0]
                selected_name = temp.split("]")[-1].strip() if "]" in temp else temp
            except:
                selected_name = display_str 

            self.var_sel_name.set(selected_name)
            self.var_sel_qty.set(1)
            
            # 從資料庫抓取該商品的詳細資料
            record = self.products_df[self.products_df['商品名稱'] == selected_name]
            if not record.empty:
                # --- 讀取編號並處理空值 ---
                raw_sku = record.iloc[0].get('商品編號', '')
                sku = str(raw_sku) if pd.notna(raw_sku) else ""
                if sku.lower() == "nan": sku = "" # 移除 pandas 的 nan 噪音
                
                # 這裡就是剛才報錯的地方，現在 self.var_sel_sku 已經在 __init__ 定義好了
                self.var_sel_sku.set(sku) 
                
                self.var_sel_cost.set(record.iloc[0]['預設成本'])
                try: 
                    stock = int(record.iloc[0]['目前庫存'])
                except: 
                    stock = 0
                self.var_sel_stock_info.set(str(stock)) 
                self.var_sel_price.set(0) # 清空上次售價
    

    def add_to_cart(self):
        name = self.var_sel_name.get()
        sku = self.var_sel_sku.get() # 這裡讀取剛才存進去的編號
        if not name: return
        
        # 容錯：如果沒編號顯示 --
        display_sku = sku if sku.strip() != "" else "--"

        try:
            qty = self.var_sel_qty.get()
            cost = self.var_sel_cost.get()
            price = self.var_sel_price.get()
            if qty <= 0: return

            total_sales = price * qty
            total_cost = cost * qty
            
            self.cart_data.append({
                "sku": sku, # 存入記憶體
                "name": name, "qty": qty, "unit_cost": cost, "unit_price": price,
                "total_sales": total_sales, "total_cost": total_cost
            })
            
            # 寫入 Treeview (確保第一欄是編號/位置)
            self.tree.insert("", "end", values=(display_sku, name, qty, price, total_sales))
            
            self.update_totals()
            
            # 清空選取狀態
            self.var_sel_name.set("")
            self.var_sel_sku.set("") # 記得也要清空編號
            self.var_sel_price.set(0)
            self.var_sel_qty.set(1)
            self.var_sel_stock_info.set("--")
            
        except ValueError: 
            messagebox.showerror("錯誤", "數字格式錯誤")

    def remove_from_cart(self):
        sel = self.tree.selection()
        if not sel: return
        for item in sel:
            idx = self.tree.index(item)
            del self.cart_data[idx]
            self.tree.delete(item)
        self.update_totals()

    def on_fee_option_selected(self, event):
        """ 當選擇費率選項時的處理邏輯 """
        selected_text = self.var_fee_rate_str.get()
        
        if "自訂" in selected_text:
            # 切換為可輸入模式
            self.combo_fee_rate.config(state="normal")
            self.var_fee_rate_str.set("")  # 清空文字讓使用者輸入
            self.combo_fee_rate.focus()    # 自動聚焦
        else:
            # 切換回唯讀模式 (針對從 Excel 讀取的固定費率)
            self.combo_fee_rate.config(state="readonly")
        
        self.update_totals()


    def update_totals_event(self, event): self.update_totals()
    
    
    def update_totals(self):
        """ 銷售輸入：使用 Decimal 進行高精度財務運算 """
        try:
            # 1. 基礎商品總額與成本 (從購物車累加)
            t_sales = Decimal("0.00")
            t_cost = Decimal("0.00")
            for i in self.cart_data:
                t_sales += Decimal(str(i['total_sales']))
                t_cost += Decimal(str(i['total_cost']))
            
            # 2. 獲取費率與固定金額 (轉為 Decimal)
            selection = self.var_fee_rate_str.get().strip()
            d_rate = Decimal("0.00")
            d_fixed = Decimal("0.00")

            if selection in self.fee_lookup:
                rate_val, fixed_val = self.fee_lookup[selection]
                d_rate = Decimal(str(rate_val))
                d_fixed = Decimal(str(fixed_val))
            else:
                try:
                    clean_input = selection.replace("%", "")
                    if clean_input: d_rate = Decimal(clean_input)
                except: d_rate = Decimal("0.00")

            # 3. 獲取運費與額外折扣
            try: d_ship = Decimal(str(self.var_ship_fee.get()))
            except: d_ship = Decimal("0.00")
            try: d_extra = Decimal(str(self.var_extra_fee.get()))
            except: d_extra = Decimal("0.00")

            payer = self.var_ship_payer.get()
            
            # --- [核心財務邏輯計算] ---
            # A. 平台手續費 (總銷售額 * 費率 + 固定費)
            platform_fee = self.dec_round(t_sales * (d_rate / Decimal("100")) + d_fixed)
            
            # B. 淨利計算
            # 基本淨利 = 售價 - 成本 - 手續費 - 折扣
            profit = t_sales - t_cost - platform_fee - d_extra
            if payer == "賣家付":
                profit -= d_ship
            
            # C. 撥款總額 (預估入帳)
            if payer == "買家付":
                income = t_sales + d_ship - platform_fee - d_extra
            else:
                income = t_sales - platform_fee - d_extra

            # --- 更新 UI 顯示 ---
            self.lbl_gross.config(text=f"商品小計: ${float(t_sales):,.2f}")
            self.lbl_fee.config(text=f"手續費({d_rate}%): -${float(platform_fee):,.2f} | 運費: ${float(d_ship):,.2f} | 折扣: -${float(d_extra):,.2f}")
            self.lbl_income.config(text=f"實收/撥款總額: ${float(income):,.2f}")
            self.lbl_profit.config(text=f"本單純利: ${float(profit):,.2f}", 
                                   foreground="green" if profit > 0 else "red")

            # 回傳 Decimal 供 submit_order 分攤使用
            return t_sales, platform_fee, Decimal("0.00")
        except Exception as e:
            print(f"system: failed to calculate sales: {e}")
            return Decimal("0"), Decimal("0"), Decimal("0")
        
    
    @thread_safe_file
    def submit_order(self):
        """ 修正版：使用 Decimal 高精度分攤手續費並存入追蹤區 """
        if not self.cart_data: return
        
        if self.var_enable_cust.get():
            cust_name = self.var_cust_name.get().strip()
            if not cust_name:
                messagebox.showerror("欄位缺失", "請務必輸入『買家名稱』！")
                self.entry_cust_name.focus(); return
            cust_loc = self.var_cust_loc.get().strip()
            ship_method = self.var_ship_method.get()
            platform_name = self.var_platform.get()
        else:
            cust_name = cust_loc = ship_method = "未提供"; platform_name = "零售/現場"
            
        date_str = self.var_date.get().strip()
        order_id = datetime.now().strftime("%Y%m%d%H%M%S") 

        # --- [獲取 Decimal 精度總額] ---
        # 確保您的 update_totals() 現在回傳的是 (Decimal, Decimal, Decimal)
        t_sales, t_fee, t_tax = self.update_totals() 
        
        fee_tag = self.var_fee_tag.get()
        try: d_extra = Decimal(str(self.var_extra_fee.get()))
        except: d_extra = Decimal("0")

        final_fee_tag = fee_tag if d_extra > Decimal("0") else ""

        try:
            rows = []
            df_prods_current = pd.read_excel(FILE_NAME, sheet_name=SHEET_PRODUCTS)

            for i, item in enumerate(self.cart_data):
                # 表頭資訊僅填於第一列
                is_first = (i == 0)
                
                # --- [Decimal 分攤計算] ---
                d_item_sales = Decimal(str(item['total_sales']))
                d_item_cost = Decimal(str(item['total_cost']))
                
                # 按銷售佔比分攤手續費與稅額
                ratio = d_item_sales / t_sales if t_sales > 0 else Decimal("0")
                alloc_fee = self.dec_round(t_fee * ratio)
                alloc_tax = self.dec_round(t_tax * ratio)
                
                # 淨利計算 (每一筆都必須精確)
                # 注意：折扣(d_extra)通常在總帳扣除，這裡分攤到每一項
                alloc_extra = self.dec_round(d_extra * ratio)
                net = d_item_sales - d_item_cost - alloc_fee - alloc_tax - alloc_extra
                
                margin_pct = (net / d_item_sales * 100) if d_item_sales > 0 else Decimal("0")

                rows.append({
                    "訂單編號": order_id,
                    "商品編號": item.get('sku', ''),
                    "日期": date_str if is_first else "",
                    "買家名稱": cust_name if is_first else "",
                    "交易平台": platform_name if is_first else "",
                    "寄送方式": ship_method if is_first else "",
                    "取貨地點": cust_loc if is_first else "",
                    "商品名稱": item['name'],
                    "數量": int(item['qty']),
                    "單價(售)": float(item['unit_price']),
                    "單價(進)": float(item['unit_cost']),
                    "總銷售額": float(d_item_sales),
                    "總成本": float(d_item_cost),
                    "分攤手續費": float(alloc_fee),
                    "扣費項目": final_fee_tag if is_first else "", # 使用清洗後的標籤
                    "總淨利": float(self.dec_round(net)),
                    "毛利率": float(self.dec_round(margin_pct, 1)),
                    "稅額": float(alloc_tax)
                })

                # 扣庫存邏輯
                p_name = item['name']
                idxs = df_prods_current[df_prods_current['商品名稱'] == p_name].index
                if not idxs.empty:
                    df_prods_current.at[idxs[0], '目前庫存'] -= int(item['qty'])


            # 讀取並合併追蹤表
            try: df_track_existing = pd.read_excel(FILE_NAME, sheet_name=SHEET_TRACKING)
            except: df_track_existing = pd.DataFrame()

            df_new_batch = pd.DataFrame(rows)
            df_new_batch['訂單編號'] = df_new_batch['訂單編號'].apply(lambda x: f"'{x}")
            df_track_combined = pd.concat([df_track_existing, df_new_batch], ignore_index=True)

            # 存檔
            if self._universal_save({
                SHEET_PRODUCTS: df_prods_current, 
                SHEET_TRACKING: df_track_combined
            }):
                self.products_df = df_prods_current
                self.update_sales_prod_list()
                self.load_tracking_data() 
                messagebox.showinfo("成功", f"訂單 {order_id} 已送出。")

                # 重置 UI
                self.cart_data = []
                for i in self.tree.get_children(): self.tree.delete(i)
                
                # 重置顧客與費用欄位
                self.var_cust_name.set("")
                self.var_ship_fee.set(0.0)
                self.var_extra_fee.set(0.0)
                self.var_fee_tag.set("")
                self.var_sel_stock_info.set("--")
                
                # 重置平台費率下拉選單狀態 (回到唯讀)
                self.combo_fee_rate.config(state="readonly")
                
                # 重新觸發計算 (會變回全 0)
                self.update_totals()

        except Exception as e: 
            messagebox.showerror("錯誤", f"存檔失敗: {str(e)}")


    def update_mgmt_prod_list(self):
        """ 及時更新商品管理清單 (過濾關鍵字) """
        search_term = self.var_mgmt_search.get().lower()
        self.listbox_mgmt.delete(0, tk.END)
        
        if not self.products_df.empty:
            for index, row in self.products_df.iterrows():
                p_name = str(row['商品名稱'])
                p_tag = str(row['分類Tag']) if pd.notna(row['分類Tag']) else "無"
                
                try: p_stock = int(row['目前庫存'])
                except: p_stock = 0
                
                display_str = f"[{p_tag}] {p_name} (庫存: {p_stock})"
                
                # 如果關鍵字出現在名稱或分類中，就顯示出來
                if search_term in p_name.lower() or search_term in p_tag.lower():
                    self.listbox_mgmt.insert(tk.END, display_str)

   
    def on_mgmt_prod_select(self, event):
        """ 當點選商品清單時，即時將資料填入右側編輯區 """
        selection = self.listbox_mgmt.curselection()
        if not selection: 
            return

        # 1. 優先定義清理工具函式 (確保在呼叫前已存在)
        def clean_val(val, default="", is_num=False):
            if pd.isna(val) or str(val).lower() == "nan":
                return 0.0 if is_num else default
            return val

        try:
            # 2. 解析 Listbox 選中的字串
            display_str = self.listbox_mgmt.get(selection[0])
            # 格式: [分類] 商品名稱 (庫存: 數量)
            temp = display_str.rsplit(" (庫存:", 1)[0]
            selected_name = temp.split("]", 1)[1].strip() if "]" in temp else temp

            # 3. 搜尋對應資料 (加上 strip 防止空格干擾)
            record = self.products_df[self.products_df['商品名稱'].astype(str).str.strip() == selected_name]
            
            if not record.empty:
                row = record.iloc[0]

                # 4. 填入字串類欄位
                self.var_upd_sku.set(clean_val(row.get('商品編號', '')))
                self.var_upd_name.set(clean_val(row.get('商品名稱', '')))
                self.var_upd_tag.set(clean_val(row.get('分類Tag', '')))
                self.var_upd_url.set(clean_val(row.get('商品連結', '')))
                self.var_upd_remarks.set(clean_val(row.get('商品備註', '')))
                
                # 5. 填入數值類欄位 (加入 is_num=True 確保 NaN 會變回 0)
                self.var_upd_safety.set(int(clean_val(row.get('安全庫存', 0), is_num=True)))
                self.var_upd_stock.set(int(clean_val(row.get('目前庫存', 0), is_num=True)))
                self.var_upd_cost.set(float(clean_val(row.get('預設成本', 0.0), is_num=True)))
                
                # 6. 填入單位權重 (特別處理)
                if hasattr(self, 'var_upd_weight'):
                    weight_val = row.get('單位權重', 1.0)
                    # 如果重量是空的，給預設值 1.0
                    if pd.isna(weight_val) or str(weight_val).lower() == "nan":
                        weight_val = 1.0
                    self.var_upd_weight.set(float(weight_val))

                self.var_upd_time.set(clean_val(row.get('最後更新時間', '無資料')))

        except Exception as e:
            # 在後台印出錯誤細節，方便除錯
            import traceback
            traceback.print_exc()
            print(f"system: failed to display selected product: {e}")


    @thread_safe_file
    def submit_new_product(self):
        """ 建立新商品：加入重複名稱檢查與資料保護邏輯 """
        name = self.var_add_name.get().strip()
        
        # 1. 基本檢查：不能為空
        if not name:
            messagebox.showwarning("警告", "『商品名稱』為必填項目！")
            return
        
        # 2. 【核心新增】重複檢查：防止同名商品再次建檔
        # 我們將所有現有商品名稱轉為小寫並去除空白後進行比對 (最嚴格檢查)
        existing_names = self.products_df['商品名稱'].astype(str).str.strip().str.lower().values
        if name.lower() in existing_names:
            messagebox.showerror("建檔失敗", 
                                 f"商品名稱「{name}」已存在於資料庫中！\n\n"
                                 "提示：\n"
                                 "1. 若要調整庫存或成本，請使用右側的「搜尋」功能選中該商品後進行修改。\n"
                                 "2. 若這是不同規格，請在名稱加上區分 (例如：黑色、白色)。")
            return

        try:
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
            url = self.var_add_url.get().strip()
            remarks = self.var_add_remarks.get().strip()

            # 準備新資料列
            new_row = {
                "商品編號": self.var_add_sku.get().strip().upper(),
                "分類Tag": self.var_add_tag.get().strip() if self.var_add_tag.get() else "未分類",
                "商品名稱": name,
                "預設成本": 0.0,
                "目前庫存": 0,
                "最後更新時間": now_str,
                "初始上架時間": now_str,
                "最後進貨時間": "",
                "安全庫存": self.var_add_safety.get(),
                "商品連結": url if url else "無",
                "商品備註": remarks if remarks else "無",
                "單位權重": self.var_add_weight.get()
            }
            
            # 合併資料並排序
            df_new = pd.concat([self.products_df, pd.DataFrame([new_row])], ignore_index=True)
            
            # 使用萬用引擎存檔
            if self._universal_save({SHEET_PRODUCTS: df_new}):
                # 重新讀取並應用我們之前的「自然排序」
                self.products_df = self.load_products() 
                self.update_mgmt_prod_list()
                self.update_pur_prod_list()
                self.update_sales_prod_list()
                
                messagebox.showinfo("成功", f"商品「{name}」已成功建檔！")
                
                # 清空左側輸入框，以便輸入下一個新商品
                self.var_add_name.set("")
                self.var_add_sku.set("")
                self.var_add_url.set("")
                self.var_add_remarks.set("")
                self.var_add_safety.set(0)
                self.var_add_weight.set(1.0)
                
        except Exception as e:
            messagebox.showerror("錯誤", f"建檔失敗: {e}")

    @thread_safe_file
    def submit_update_product(self):
        name = self.var_upd_name.get()
        if not name: return
        
        try:
            # --- [安全數值抓取] ---
            # 使用 try-except 確保即使介面上有 NaN 字樣，程式也不會崩潰
            try: new_cost = float(self.var_upd_cost.get())
            except: new_cost = 0.0
            
            try: new_stock = int(self.var_upd_stock.get())
            except: new_stock = 0

            try: new_safety = int(self.var_upd_safety.get())
            except: new_safety = 0

            try: new_weight = float(self.var_upd_weight.get())
            except: new_weight = 1.0

            now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
            
            # 1. 讀取商品資料分頁
            df_prods = pd.read_excel(FILE_NAME, sheet_name=SHEET_PRODUCTS)
            
            # 2. 定位商品
            idx = df_prods[df_prods['商品名稱'] == name].index
            if not idx.empty:
                # 取得舊庫存 (處理可能的 NaN)
                old_stock = df_prods.loc[idx, '目前庫存'].values[0]
                if pd.isna(old_stock): old_stock = 0
                
                # --- [補齊舊資料欄位/補貨邏輯] ---
                if "初始上架時間" not in df_prods.columns: 
                    df_prods["初始上架時間"] = df_prods["最後更新時間"]
                if "最後進貨時間" not in df_prods.columns: 
                    df_prods["最後進貨時間"] = df_prods["最後更新時間"]

                if new_stock > old_stock:
                    df_prods.loc[idx, '最後進貨時間'] = now_str
                    print(f"system: detected restock for product {name}, updated restock time.")
                
                # --- [更新資料列] ---
                df_prods.loc[idx, '商品編號'] = self.var_upd_sku.get()
                df_prods.loc[idx, '分類Tag'] = self.var_upd_tag.get()
                df_prods.loc[idx, '商品名稱'] = self.var_upd_name.get()
                df_prods.loc[idx, '預設成本'] = new_cost
                df_prods.loc[idx, '目前庫存'] = new_stock
                df_prods.loc[idx, '安全庫存'] = new_safety
                df_prods.loc[idx, '商品連結'] = self.var_upd_url.get()
                df_prods.loc[idx, '商品備註'] = self.var_upd_remarks.get()
                df_prods.loc[idx, '單位權重'] = new_weight
                df_prods.loc[idx, '最後更新時間'] = now_str
                
                # --- [呼叫萬用存檔引擎] ---
                # 這是最強的保護措施，它會自動讀取 SHEET_SALES, SHEET_TRACKING 等所有分頁
                # 並一次性寫回，防止任何資料丟失。
                if self._universal_save({SHEET_PRODUCTS: df_prods}):
                    # 更新成功後的後續動作
                    self.products_df = self.load_products() 
                    self.update_mgmt_prod_list()
                    self.update_sales_prod_list() # 讓銷售頁面也同步看到新庫存
                    self.var_upd_time.set(now_str) 
                    messagebox.showinfo("成功", f"商品「{name}」資訊已更新！")
                
        except PermissionError: 
            messagebox.showerror("錯誤", "Excel 檔案未關閉，無法寫入！")
        except Exception as e:
            import traceback
            traceback.print_exc() # 在後台印出詳細錯誤以便除錯
            messagebox.showerror("錯誤", f"更新失敗: {e}")
    
    @thread_safe_file
    def delete_product(self):
        name = self.var_upd_name.get()
        if not name: return
        confirm = messagebox.askyesno("確認刪除", f"確定要刪除「{name}」嗎？\n\n此動作無法復原！")
        if not confirm: return
        try:
            df_old = pd.read_excel(FILE_NAME, sheet_name='商品資料')
            df_new = df_old[df_old['商品名稱'] != name]
            with pd.ExcelWriter(FILE_NAME, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df_new.to_excel(writer, sheet_name='商品資料', index=False)
            self.products_df = df_new
            self.update_sales_prod_list()
            self.update_mgmt_prod_list()
            
            self.var_upd_name.set("")
            self.var_upd_tag.set("")
            self.var_upd_cost.set(0)
            self.var_upd_stock.set(0)
            self.var_upd_time.set("尚無資料")
            messagebox.showinfo("成功", f"已刪除商品：{name}")
        except PermissionError: messagebox.showerror("錯誤", "Excel 未關閉！")



def start_main_app():
    """ 這是原本啟動主程式的邏輯，包裝成一個 function """
    root = tk.Tk()
    style = ttk.Style()

     # --- [核心修正：動態判斷作業系統] ---
    current_os = platform.system()
    try:
        if current_os == "Windows":
            # Windows 專用主題
            style.theme_use('vista') 
        elif current_os == "Darwin": # Darwin 代表 macOS
            # Mac 內建主題是 'aqua'，通常不需要特別呼叫
            # 或者使用相容性最好的 'clam'
            style.theme_use('clam')
        else:
            # Linux 等其他系統
            style.theme_use('clam')
    except Exception as e:
        print(f"Theme Error: {e}")
    # 如果出錯，就維持系統預設，不強行設定

    SalesApp(root)

    root.mainloop()


if __name__ == "__main__":
    # 1. 先顯示登入視窗
    # 2. 傳入 start_main_app 作為成功後的執行動作
    login = LoginWindow(start_main_app)
    login.run()



# if __name__ == "__main__":
#     root = tk.Tk()
#     style = ttk.Style()
#     try:
#         style.theme_use('vista') 
#     except:
#         pass 
#     app = SalesApp(root)
#     root.mainloop()


