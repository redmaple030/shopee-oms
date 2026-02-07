#shopee-oms 3.9 完整版

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


# 1. 匯入敏感資料
try:
    from secrets_config import SECRET_SALT
except ImportError:
    SECRET_SALT = "DEMO_SALT_FOR_OPENSOURCE"


# 2. 加入這段函式：用來處理打包後的資源路徑
def resource_path(relative_path):
    """ 獲取資源的絕對路徑，兼容 Dev 和 PyInstaller """
    try:
        # PyInstaller 創建臨時文件夾，路徑存儲在 _MEIPASS 中
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)
    


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
FILE_NAME = 'sales_data.xlsx'
CREDENTIALS_FILE = resource_path('credentials.json')  
TOKEN_FILE = 'token.json'             
SCOPES = ['https://www.googleapis.com/auth/drive.file'] 

SHEET_SALES = '銷售紀錄'      # 歷史已完成訂單
SHEET_TRACKING = '訂單追蹤'   # 未完成/出貨中 (緩衝區)
SHEET_RETURNS = '退貨紀錄'    # 退貨區
SHEET_PRODUCTS = '商品資料'
SHEET_CONFIG = '系統設定'


# 設定雲端硬碟上的備份資料夾名稱
BACKUP_FOLDER_NAME = "蝦皮進銷存系統_備份"

TAIWAN_CITIES = [
    "基隆市", "臺北市", "新北市", "桃園市", "新竹市", "新竹縣", "苗栗縣",
    "臺中市", "彰化縣", "南投縣", "雲林縣", "嘉義市", "嘉義縣", "臺南市",
    "高雄市", "屏東縣", "宜蘭縣", "花蓮縣", "臺東縣", "澎湖縣", "金門縣", "連江縣",
    "海外", "面交"
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

SHOPEE_FEE_OPTIONS = [
    "自訂手動輸入",
    "一般賣家-平日 (14.0%)",         
    "一般賣家-促銷檔期 (16.0%)",     
    "一般賣家-較長備貨-平日 (17.0%)", 
    "一般賣家-較長備貨-促銷 (19.0%)", 
    "商城-平日 (17.0%)",             
    "商城-促銷檔期 (20.9%)",         
    "商城-較長備貨-平日 (20.0%)",
    "商城-較長備貨-促銷 (23.9%)"
]


class GoogleDriveSync:
    """處理 Google Drive 認證、資料夾管理、上傳與下載邏輯"""
    def __init__(self):
        self.creds = None
        self.service = None
        self.is_authenticated = False
        self.folder_id = None 

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
            print(f"資料夾建立失敗: {e}")
            return None

    def upload_file(self, filepath):
        """上傳檔案到指定資料夾，並維持最多 10 筆備份"""
        if not self.is_authenticated: return False, "尚未登入 Google 帳號"
        if not self.folder_id: self.folder_id = self.get_or_create_folder()

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
            
            if len(items) > 10:
                # 取得第 11 筆之後的所有檔案 (即最舊的檔案們)
                files_to_delete = items[10:] 
                for old_file in files_to_delete:
                    file_id = old_file.get('id')
                    try:
                        self.service.files().delete(fileId=file_id).execute()
                        print(f"自動清理舊備份: {old_file.get('name')}")
                    except Exception as delete_error:
                        print(f"刪除舊檔失敗: {delete_error}")

            return True, f"備份成功！\n雲端檔名: {file_name}\n(系統已自動保留最新 10 筆紀錄)"
        except Exception as e:
            return False, f"上傳失敗: {str(e)}"

    def list_backups(self):
        """列出備份資料夾內的檔案"""
        if not self.is_authenticated: return []
        if not self.folder_id: self.folder_id = self.get_or_create_folder()
        
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
        if not self.is_authenticated: return False, "尚未登入"
        
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

class SalesApp:
    
    def __init__(self, root):
        self.root = root
        self.root.title("蝦皮/網拍進銷存系統 (V3.8 完整版)")
        self.root.geometry("1280x850") 

        # --- 字型設定 ---
        self.default_font_size = 11
        self.style = ttk.Style()
        self.setup_fonts(self.default_font_size)

        self.drive_manager = GoogleDriveSync()

        # --- 變數初始化 ---
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
        self.create_tabs()
         # 啟動時自動檢查授權
        self.check_license_on_startup()

    
   

    def setup_fonts(self, size):
        default_font = font.nametofont("TkDefaultFont")
        default_font.configure(family="微軟正黑體", size=size)
        
        text_font = font.nametofont("TkTextFont")
        text_font.configure(family="微軟正黑體", size=size)

        self.style.configure(".", font=("微軟正黑體", size))
        self.style.configure("Treeview", rowheight=size*3) 
        self.style.configure("Treeview.Heading", font=("微軟正黑體", size, "bold"))
        self.style.configure("TLabelframe.Label", font=("微軟正黑體", size, "bold"))

    def change_font_size(self, event=None):
        try:
            new_size = int(self.var_font_size.get())
            self.setup_fonts(new_size)
        except:
            pass



    def check_excel_file(self):
            cols_sales = ["訂單編號", "日期", "買家名稱", "交易平台", "寄送方式", "取貨地點", 
                      "商品名稱", "數量", "單價(售)", "單價(進)", "總銷售額", "總成本", "分攤手續費", "扣費項目", "總淨利", "毛利率"]
            cols_prods = ["分類Tag", "商品名稱", "預設成本", "目前庫存", "最後更新時間", "初始上架時間", "最後進貨時間"]

            cols_config = ["設定名稱", "費率百分比"]
            default_fees = [["一般賣家-平日", 14.5], ["一般賣家-大促", 16.5], ["免運賣家", 19.5], ["自訂費率", 10.0]]
    
            if not os.path.exists(FILE_NAME):
                try:
                    with pd.ExcelWriter(FILE_NAME, engine='openpyxl') as writer:
                        pd.DataFrame(columns=cols_sales).to_excel(writer, sheet_name=SHEET_SALES, index=False)
                        pd.DataFrame(columns=cols_sales).to_excel(writer, sheet_name=SHEET_TRACKING, index=False)
                        pd.DataFrame(columns=cols_sales).to_excel(writer, sheet_name=SHEET_RETURNS, index=False)
                        # 建立商品範例
                        now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

                        df_prods = pd.DataFrame([["範例分類", "範例商品A", 100, 10, datetime.now().strftime("%Y-%m-%d %H:%M")]], 
                                                columns=["分類Tag", "商品名稱", "預設成本", "目前庫存", "最後更新時間"])
                        df_prods.to_excel(writer, sheet_name=SHEET_PRODUCTS, index=False)
                        # 建立預設費率
                        pd.DataFrame(default_fees, columns=cols_config).to_excel(writer, sheet_name=SHEET_CONFIG, index=False)
                except Exception as e:
                    messagebox.showerror("錯誤", f"無法建立 Excel: {e}")
                else:
                    # 檢查是否缺少設定分頁
                    try:
                        with pd.ExcelWriter(FILE_NAME, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                            if SHEET_CONFIG not in writer.book.sheetnames:
                                pd.DataFrame(default_fees, columns=cols_config).to_excel(writer, sheet_name=SHEET_CONFIG, index=False)
                    except: pass

    def load_products(self):
        try:
            df = pd.read_excel(FILE_NAME, sheet_name='商品資料')

            # --- [核心相容邏輯] ---
            # 如果是舊檔案，缺這兩欄，就用現有的「最後更新時間」填補
            if "初始上架時間" not in df.columns:
                df["初始上架時間"] = df["最後更新時間"]
            if "最後進貨時間" not in df.columns:
                df["最後進貨時間"] = df["最後更新時間"]
            # ---------------------

            if "分類Tag" not in df.columns: df["分類Tag"] = ""
            if "目前庫存" not in df.columns: 
                df["目前庫存"] = 0 
            else:
                df["目前庫存"] = df["目前庫存"].fillna(0).astype(int)
            df = df.sort_values(by=['分類Tag', '商品名稱'], na_position='last')
            return df
        except:
            return pd.DataFrame(columns=["分類Tag", "商品名稱", "預設成本", "目前庫存", "最後更新時間"])

    def create_tabs(self):
        tab_control = ttk.Notebook(self.root)
        
        self.tab_sales = ttk.Frame(tab_control)
        self.tab_tracking = ttk.Frame(tab_control) 
        self.tab_returns = ttk.Frame(tab_control) # [新增] 退貨紀錄頁面
        self.tab_sales_edit = ttk.Frame(tab_control) 
        self.tab_products = ttk.Frame(tab_control)
        self.tab_analysis = ttk.Frame(tab_control)
        self.tab_backup = ttk.Frame(tab_control) 
        self.tab_about = ttk.Frame(tab_control)
        
        tab_control.add(self.tab_sales, text='銷售輸入')
        tab_control.add(self.tab_tracking, text='訂單追蹤查詢')
        tab_control.add(self.tab_returns, text='退貨紀錄查詢')
        tab_control.add(self.tab_sales_edit, text='銷售紀錄(已結案)') 
        tab_control.add(self.tab_products, text='商品資料管理')
        tab_control.add(self.tab_analysis, text='營收分析')
        tab_control.add(self.tab_backup, text='雲端備份/資料復原') 
        tab_control.add(self.tab_about, text='設定及關於')
        
        tab_control.pack(expand=1, fill="both")
        
        self.setup_about_tab()   
        self.setup_sales_tab()
        self.setup_tracking_tab()
        self.setup_returns_tab()
        self.setup_sales_edit_tab()
        self.setup_product_tab()
        self.setup_analysis_tab()
        self.setup_backup_tab() 



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

    def calculate_analysis_data(self):
        """ 核心分析邏輯修正版：使用『初始上架時間』計算長期銷售速度 """
        if not hasattr(self, 'tree_time_stats') or not hasattr(self, 'tree_prod_stats'): return
        
        for i in self.tree_time_stats.get_children(): self.tree_time_stats.delete(i)
        for i in self.tree_prod_stats.get_children(): self.tree_prod_stats.delete(i)
        
        if not os.path.exists(FILE_NAME): return

        try:
            with pd.ExcelFile(FILE_NAME) as xls:
                df_sales = pd.read_excel(xls, sheet_name=SHEET_SALES)
                df_prods = pd.read_excel(xls, sheet_name=SHEET_PRODUCTS)

            if df_sales.empty: return

            # --- 補齊 Excel 視覺空白 (ffill) ---
            fill_cols = ['訂單編號', '日期', '買家名稱', '交易平台']
            for col in fill_cols:
                if col in df_sales.columns: df_sales[col] = df_sales[col].ffill()
            df_sales = df_sales.dropna(subset=['商品名稱'])

            # --- 資料清洗 ---
            for col in ['總銷售額', '總淨利', '數量']:
                df_sales[col] = pd.to_numeric(df_sales[col], errors='coerce').fillna(0)
            df_sales['日期'] = pd.to_datetime(df_sales['日期'], errors='coerce')
            df_sales = df_sales.dropna(subset=['日期'])
            df_sales['毛利率_數值'] = pd.to_numeric(df_sales['毛利率'].astype(str).str.replace('%', ''), errors='coerce').fillna(0)

            # --- 左側：月份統計 (邏輯不變) ---
            # ... (此處維持您原本的月份統計顯示，略)

            # --- 右側：商品分析 (修正速度計算) ---
            # 1. 取得『初始上架時間』作為計算基準 (分母)
            # 如果舊資料沒有這個欄位，會自動用最後更新時間補齊
            if "初始上架時間" not in df_prods.columns:
                df_prods["初始上架時間"] = df_prods["最後更新時間"]
            
            df_prods['初始上架時間'] = pd.to_datetime(df_prods['初始上架時間'], errors='coerce').fillna(pd.Timestamp.now())
            
            # 建立上架時間對照表
            first_upload_map = df_prods.set_index('商品名稱')['初始上架時間']

            # 2. 聚合銷售數據 (分子)
            prod_group = df_sales.groupby('商品名稱').agg({
                '毛利率_數值': 'mean',
                '總淨利': 'sum',
                '數量': 'sum'
            }).reset_index()

            # 3. 【核心修正點】計算長期銷售速度
            now = pd.Timestamp.now()
            
            # 取得該商品自從上架以來的總天數
            prod_group['start_date'] = prod_group['商品名稱'].map(first_upload_map).fillna(now)
            
            # 計算總時長 (天數)，最少為 1 天避免除以 0
            prod_group['total_days'] = (now - prod_group['start_date']).dt.days.clip(lower=1)
            
            # 銷售速率 = 總銷量 / 總天數
            prod_group['velocity'] = (prod_group['數量'] / prod_group['total_days']).round(2)

            # 4. 排序與顯示
            sort_mode = self.var_prod_sort_by.get()
            sort_map = {
                "平均毛利率": '毛利率_數值',
                "總銷量排行": '數量',
                "總獲利排行": '總淨利',
                "銷售速度排行": 'velocity'
            }
            prod_group = prod_group.sort_values(sort_map.get(sort_mode, '毛利率_數值'), ascending=False)

            for _, row in prod_group.iterrows():
                self.tree_prod_stats.insert("", "end", values=(
                    row['商品名稱'],
                    f"{row['毛利率_數值']:.1f}%",
                    f"${row['總淨利']:,.0f}",
                    int(row['數量']),
                    f"{row['velocity']} 件/日"
                ))

        except Exception as e:
            print(f"分析失敗: {e}")

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

        # 讀取全域變數的 SALT
        # raw_string = user_id + SECRET_SALT  <-- 記得這裡要用全域變數，不要重複定義
        try:
            # 確保有讀到 SECRET_SALT，如果沒有定義，就用預設值 (避免報錯)
            salt = globals().get('SECRET_SALT', "DEMO_SALT_FOR_OPENSOURCE")
            raw_string = user_id + salt
        except:
             raw_string = user_id + "DEMO_SALT_FOR_OPENSOURCE"

        expected_code = hashlib.md5(raw_string.encode()).hexdigest()[:8].upper()
        
        if input_code == expected_code:
            self.is_vip = True
            
            # === 【新增這段】儲存授權檔與路徑 ===
            try:
                current_path = os.path.abspath(sys.executable)
                save_data = {
                    "user_id": user_id,
                    "license_key": input_code,
                    "install_path": current_path  # 綁定目前路徑
                }
                with open("license.json", "w", encoding="utf-8") as f:
                    json.dump(save_data, f)
            except Exception as e:
                messagebox.showerror("錯誤", f"授權存檔失敗: {e}")
            # ===================================

            messagebox.showinfo("成功", "VIP 功能已解鎖！\n程式已綁定此資料夾。\n若移動程式位置，需重新輸入啟用碼。")
            
            # 解鎖按鈕
            self.btn_login.config(state="normal")
            self.lbl_auth_status.config(text="狀態: 尚未連結 (請點擊登入)", foreground="red")
            
            if self.drive_manager.is_authenticated:
                 self.btn_upload.config(state="normal")
                 self.btn_refresh.config(state="normal")
        else:
            messagebox.showerror("錯誤", "啟用碼錯誤！")


        

    def check_license_on_startup(self):
        """
        程式啟動時，檢查是否有有效的授權檔
        驗證:1. 金鑰正確性 2. 執行路徑是否改變
        """
        if not os.path.exists("license.json"):
            return # 沒有授權檔，保持鎖定
            
        try:
            with open("license.json", "r", encoding="utf-8") as f:
                data = json.load(f)
            
            saved_user = data.get("user_id", "")
            saved_key = data.get("license_key", "")
            bound_path = data.get("install_path", "")
            
            # === 1. 檢查路徑是否改變 (防複製/移動) ===
            # sys.executable 會抓到目前 .exe 的絕對路徑
            current_path = os.path.abspath(sys.executable)
            
            # 如果是在開發環境 (py檔)，sys.executable 會是 python.exe 的路徑，
            # 為了方便測試，我們可以放寬開發環境的檢查，只針對打包後的 EXE 檢查
            if getattr(sys, 'frozen', False): 
                # 這是打包後的 EXE 環境
                if current_path != bound_path:
                    # 路徑不符，視為非法移動
                    messagebox.showwarning("授權失效", "偵測到程式已被移動或複製！\n為了安全起見，請重新輸入啟用碼進行綁定。")
                    try:
                        os.remove("license.json") # 刪除舊授權
                    except:
                        pass
                    return 

            # === 2. 重新驗證金鑰 (防修改存檔) ===
            try:
                salt = globals().get('SECRET_SALT', "DEMO_SALT_FOR_OPENSOURCE")
                raw_string = saved_user + salt
            except:
                raw_string = saved_user + "DEMO_SALT_FOR_OPENSOURCE"
                
            expected_code = hashlib.md5(raw_string.encode()).hexdigest()[:8].upper()
            
            if saved_key == expected_code:
                # 通過驗證！自動解鎖
                self.is_vip = True
                self.var_vip_user.set(saved_user)
                self.var_vip_code.set(saved_key)
                
                # 解鎖 UI
                self.btn_login.config(state="normal")
                self.lbl_auth_status.config(text="狀態: 🔒 VIP 授權有效 (自動登入)", foreground="green")
                
                # 如果有 token，連備份按鈕也一起開
                if self.drive_manager.is_authenticated:
                    self.btn_upload.config(state="normal")
                    self.btn_refresh.config(state="normal")
                    self.lbl_auth_status.config(text="狀態: ✅ 系統就緒 (已連結 Google)", foreground="green")
        except Exception as e:
            print(f"授權讀取失敗: {e}")

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

        cols = ("商品名稱", "數量", "單價", "總計")
        self.tree = ttk.Treeview(right_frame, columns=cols, show='headings', height=8)
        self.tree.heading("商品名稱", text="商品名稱")
        self.tree.column("商品名稱", width=120)
        self.tree.heading("單價", text="售價")
        self.tree.column("單價", width=80, anchor="e")
        self.tree.heading("數量", text="數量")
        self.tree.column("數量", width=60, anchor="center")
        self.tree.heading("總計", text="小計")
        self.tree.column("總計", width=70, anchor="e")
        self.tree.pack(fill="both", expand=True)

        ttk.Button(right_frame, text="(x) 移除", command=self.remove_from_cart).pack(anchor="e", pady=2)

        fee_frame = ttk.LabelFrame(right_frame, text="費用與折扣", padding=10)
        fee_frame.pack(fill="x", pady=5)
        
        f1 = ttk.Frame(fee_frame)
        f1.pack(fill="x")
        ttk.Label(f1, text="費率:").pack(side="left")
        
        self.combo_fee_rate = ttk.Combobox(f1, textvariable=self.var_fee_rate_str, values=SHOPEE_FEE_OPTIONS, width=28)
        self.combo_fee_rate.pack(side="left", padx=5)
        self.combo_fee_rate.set("一般賣家-平日 (14.5%)") 
        self.combo_fee_rate.bind('<<ComboboxSelected>>', self.on_fee_option_selected)
        self.combo_fee_rate.bind('<KeyRelease>', self.update_totals_event)

        self.var_tax_enabled = tk.BooleanVar(value=False)


        f2 = ttk.Frame(fee_frame)
        f2.pack(fill="x", pady=5)

        self.var_tax_enabled = tk.BooleanVar(value=False)
        ttk.Checkbutton(f2, text="開發票(5%稅)", variable=self.var_tax_enabled, command=self.update_totals).pack(side="left", padx=5)

        
        tag_opts = ["", "活動費", "運費補貼", "補償金額", "私人預定", "補寄補貼", "固定成本"]
        self.combo_tag = ttk.Combobox(f2, textvariable=self.var_fee_tag, values=tag_opts, state="readonly", width=12)
        self.combo_tag.pack(side="left")
        self.combo_tag.set("扣費原因")

        ttk.Label(f2, text=" 金額$").pack(side="left", padx=2)
        e_extra = ttk.Entry(f2, textvariable=self.var_extra_fee, width=8)
        e_extra.pack(side="left")
        e_extra.bind('<KeyRelease>', self.update_totals_event)
        
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

        ttk.Button(sum_frame, text="✔ 送出訂單", command=self.submit_order).pack(fill="x", pady=5)

        self.refresh_fee_tree()



    def setup_product_tab(self):
        paned = ttk.PanedWindow(self.tab_products, orient=tk.HORIZONTAL)
        paned.pack(fill="both", expand=True, padx=10, pady=10)

        frame_add = ttk.LabelFrame(paned, text="新增商品", padding=15)
        paned.add(frame_add, weight=1)

        ttk.Label(frame_add, text="1. 分類Tag:").pack(anchor="w", pady=(0,5))
        self.combo_add_tag = ttk.Combobox(frame_add, textvariable=self.var_add_tag)
        self.combo_add_tag.pack(fill="x", pady=5)
        self.combo_add_tag.bind('<Button-1>', self.load_existing_tags)

        ttk.Label(frame_add, text="2. 商品名稱:").pack(anchor="w", pady=(10,5))
        ttk.Entry(frame_add, textvariable=self.var_add_name).pack(fill="x", pady=5)

        ttk.Label(frame_add, text="3. 進貨成本:").pack(anchor="w", pady=(10,5))
        ttk.Entry(frame_add, textvariable=self.var_add_cost).pack(fill="x", pady=5)
        
        ttk.Label(frame_add, text="4. 初始庫存:").pack(anchor="w", pady=(10,5))
        ttk.Entry(frame_add, textvariable=self.var_add_stock).pack(fill="x", pady=5)

        ttk.Button(frame_add, text="+ 新增", command=self.submit_new_product).pack(fill="x", pady=20)

        frame_upd = ttk.LabelFrame(paned, text="更新商品", padding=15)
        paned.add(frame_upd, weight=1)

        ttk.Label(frame_upd, text="搜尋關鍵字:").pack(anchor="w")
        e_search = ttk.Entry(frame_upd, textvariable=self.var_mgmt_search)
        e_search.pack(fill="x", pady=5)
        e_search.bind('<KeyRelease>', self.update_mgmt_prod_list)

        list_frame = ttk.Frame(frame_upd)
        list_frame.pack(fill="both", expand=True, pady=5)
        self.listbox_mgmt = tk.Listbox(list_frame, height=10)
        sb = ttk.Scrollbar(list_frame, orient="vertical", command=self.listbox_mgmt.yview)
        self.listbox_mgmt.configure(yscrollcommand=sb.set)
        self.listbox_mgmt.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.listbox_mgmt.bind('<<ListboxSelect>>', self.on_mgmt_prod_select)

        edit_frame = ttk.LabelFrame(frame_upd, text="編輯選中商品", padding=10)
        edit_frame.pack(fill="x", pady=10)

        ttk.Label(edit_frame, text="名稱 (不可改):").grid(row=0, column=0, sticky="w")
        ttk.Entry(edit_frame, textvariable=self.var_upd_name, state="readonly").grid(row=0, column=1, sticky="ew", padx=5)

        ttk.Label(edit_frame, text="分類Tag:").grid(row=1, column=0, sticky="w", pady=5)
        self.combo_upd_tag = ttk.Combobox(edit_frame, textvariable=self.var_upd_tag, width=18)
        self.combo_upd_tag.grid(row=1, column=1, sticky="ew", padx=5, pady=5)
        self.combo_upd_tag.bind('<Button-1>', self.load_existing_tags)

        ttk.Label(edit_frame, text="成本:").grid(row=2, column=0, sticky="w", pady=5)
        ttk.Entry(edit_frame, textvariable=self.var_upd_cost).grid(row=2, column=1, sticky="ew", padx=5, pady=5)
        
        ttk.Label(edit_frame, text="庫存(補貨):").grid(row=3, column=0, sticky="w", pady=5)
        ttk.Entry(edit_frame, textvariable=self.var_upd_stock).grid(row=3, column=1, sticky="ew", padx=5, pady=5)

        ttk.Label(edit_frame, text="更新時間:").grid(row=4, column=0, sticky="w")
        ttk.Label(edit_frame, textvariable=self.var_upd_time, foreground="gray").grid(row=4, column=1, sticky="w", padx=5)

        btn_frame = ttk.Frame(edit_frame)
        btn_frame.grid(row=5, column=0, columnspan=2, pady=10, sticky="ew")
        
        ttk.Button(btn_frame, text="💾 儲存", command=self.submit_update_product).pack(side="left", fill="x", expand=True, padx=(0, 5))
        ttk.Button(btn_frame, text="🗑️ 刪除", command=self.delete_product).pack(side="left", fill="x", expand=True, padx=(5, 0))

        self.update_mgmt_prod_list()


    def setup_tracking_tab(self):
        """ 建立訂單追蹤區 (緩衝區) """
        frame = self.tab_tracking
        # 1. 頂部操作
        top_frame = ttk.Frame(frame, padding=5)
        top_frame.pack(fill="x")
        ttk.Button(top_frame, text="🔄 重新整理列表", command=self.load_tracking_data).pack(side="right")
        ttk.Label(top_frame, text="此處為緩衝區。結案後進入「銷售紀錄」，退貨後進入「退貨紀錄」。", foreground="gray").pack(side="left")

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


    def load_tracking_data(self):
        """ 讀取『訂單追蹤』分頁的資料 (新增此函式) """
        for i in self.tree_track.get_children():
            self.tree_track.delete(i)
        try:
            if not os.path.exists(FILE_NAME): return
            df = pd.read_excel(FILE_NAME, sheet_name=SHEET_TRACKING)
            if '訂單編號' in df.columns:
                df['訂單編號'] = df['訂單編號'].astype(str).str.replace(r'\.0$', '', regex=True)
            df = df.fillna("")
            last_id, last_date, last_platform, last_buyer = "", "", "", ""
            for idx, row in df.iterrows():
                order_id = str(row.get('訂單編號', ''))
                date = str(row.get('日期', ''))
                platform = str(row.get('交易平台', ''))
                buyer = str(row.get('買家名稱', ''))
                if order_id == "nan" or order_id == "": order_id = last_id
                else: last_id = order_id
                if date == "": date = last_date
                else: last_date = date
                if platform == "": platform = last_platform
                else: last_platform = platform
                if buyer == "": buyer = last_buyer
                else: last_buyer = buyer
                self.tree_track.insert("", "end", text=str(idx), values=(
                    order_id, date, platform, buyer,
                    row.get('商品名稱', ''),
                    int(row.get('數量', 0) if row.get('數量') != "" else 0),
                    float(row.get('單價(售)', 0) if row.get('單價(售)') != "" else 0)
                ))
        except Exception as e:
            print(f"讀取追蹤清單失敗: {e}")

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
                self._save_all_sheets(df, SHEET_TRACKING)
                messagebox.showinfo("成功", "資料已更新"); self.load_tracking_data(); win.destroy()
            except Exception as e: messagebox.showerror("錯誤", f"存檔失敗: {e}")
        tk.Button(win, text="確認修改", command=save_mod).pack(pady=15)

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
            self._save_all_sheets(df, SHEET_TRACKING)
            messagebox.showinfo("成功", "商品已刪除"); self.load_tracking_data()
        except Exception as e: messagebox.showerror("錯誤", f"刪除失敗: {e}")

    def action_track_delete_order(self):
        """ 刪除整筆訂單 """
        sel = self.tree_track.selection()
        if not sel: return
        item = self.tree_track.item(sel[0]); order_id = str(item['values'][0])
        if not messagebox.askyesno("刪除整筆", f"確定要刪除訂單 [{order_id}] 嗎？"): return
        try:
            df = pd.read_excel(FILE_NAME, sheet_name=SHEET_TRACKING)
            df['訂單編號'] = df['訂單編號'].astype(str).str.replace(r'\.0$', '', regex=True)
            df_new = df[df['訂單編號'] != order_id]
            self._save_all_sheets(df_new, SHEET_TRACKING)
            messagebox.showinfo("成功", "整筆訂單已刪除"); self.load_tracking_data()
        except Exception as e: messagebox.showerror("錯誤", f"刪除失敗: {e}")

    def action_track_return_order(self):
        #""" 退貨整筆訂單 """
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
            for col, val in info.items(): rows_to_return[col] = val # 補齊資料
            rows_to_return['備註'] = reason
            
            try: df_returns = pd.read_excel(FILE_NAME, sheet_name=SHEET_RETURNS)
            except: df_returns = pd.DataFrame()
            df_returns = pd.concat([df_returns, rows_to_return], ignore_index=True)
            df_track_new = df_track[~mask]
            
            self._save_all_sheets_with_protect(df_track_new, SHEET_TRACKING, df_returns, SHEET_RETURNS)
            messagebox.showinfo("成功", f"訂單 {order_id} 整筆已移至退貨。")
            self.load_tracking_data(); self.load_returns_data()
        except Exception as e: messagebox.showerror("錯誤", str(e))

    def _save_all_sheets(self, df_target, target_sheet_name):
        """ 輔助函式：保留其他分頁並儲存 (新增此函式) """
        with pd.ExcelWriter(FILE_NAME, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_target.to_excel(writer, sheet_name=target_sheet_name, index=False)
            for sheet in [SHEET_SALES, SHEET_PRODUCTS, SHEET_RETURNS]:
                if sheet != target_sheet_name:
                    try:
                        df = pd.read_excel(FILE_NAME, sheet_name=sheet)
                        df.to_excel(writer, sheet_name=sheet, index=False)
                    except:
                        pd.DataFrame().to_excel(writer, sheet_name=sheet, index=False)


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
            print(f"讀取退貨紀錄失敗: {e}")
    
    #================= 銷售紀錄 =================
    def setup_sales_edit_tab(self):
        paned = ttk.PanedWindow(self.tab_sales_edit, orient=tk.VERTICAL)
        paned.pack(fill="both", expand=True, padx=10, pady=10)

        # 1. 上方：列表區
        list_frame = ttk.LabelFrame(paned, text="銷售歷史紀錄 (點擊項目進行修改)", padding=5)
        paned.add(list_frame, weight=3)

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

        # 2. 下方：編輯區
        edit_frame = ttk.LabelFrame(paned, text="✏️ 修改選中資料 (數值修改後，系統會自動重算毛利)", padding=15)
        paned.add(edit_frame, weight=1)

        # 變數宣告
        self.var_edit_idx = tk.IntVar(value=-1) # 紀錄 Excel 中的原始索引
        self.var_edit_date = tk.StringVar()
        self.var_edit_name = tk.StringVar()
        self.var_edit_qty = tk.IntVar(value=0)
        self.var_edit_price = tk.DoubleVar(value=0)
        self.var_edit_cost = tk.DoubleVar(value=0)
        self.var_edit_fee = tk.DoubleVar(value=0)
        self.var_edit_deduct = tk.DoubleVar(value=0) # 其他扣費

        # 排版 (Grid)
        grid_opts = {'padx': 5, 'pady': 5, 'sticky': 'w'}
        
        ttk.Label(edit_frame, text="訂單日期:").grid(row=0, column=0, **grid_opts)
        ttk.Entry(edit_frame, textvariable=self.var_edit_date, width=15).grid(row=0, column=1, **grid_opts)

        ttk.Label(edit_frame, text="商品名稱:").grid(row=0, column=2, **grid_opts)
        ttk.Entry(edit_frame, textvariable=self.var_edit_name, width=25).grid(row=0, column=3, **grid_opts)

        ttk.Label(edit_frame, text="數量:").grid(row=1, column=0, **grid_opts)
        ttk.Entry(edit_frame, textvariable=self.var_edit_qty, width=10).grid(row=1, column=1, **grid_opts)

        ttk.Label(edit_frame, text="單價(售):").grid(row=1, column=2, **grid_opts)
        ttk.Entry(edit_frame, textvariable=self.var_edit_price, width=10).grid(row=1, column=3, **grid_opts)

        ttk.Label(edit_frame, text="單價(進):").grid(row=2, column=0, **grid_opts)
        ttk.Entry(edit_frame, textvariable=self.var_edit_cost, width=10).grid(row=2, column=1, **grid_opts)

        ttk.Label(edit_frame, text="手續費:").grid(row=2, column=2, **grid_opts)
        ttk.Entry(edit_frame, textvariable=self.var_edit_fee, width=10).grid(row=2, column=3, **grid_opts)
        
        ttk.Label(edit_frame, text="其他扣費:").grid(row=2, column=4, **grid_opts)
        ttk.Entry(edit_frame, textvariable=self.var_edit_deduct, width=8).grid(row=2, column=5, **grid_opts)

        # 按鈕區
        btn_area = ttk.Frame(edit_frame)
        btn_area.grid(row=3, column=0, columnspan=6, pady=15, sticky="ew")
        
        ttk.Button(btn_area, text="💾 確認修改並重算", command=self.save_sales_edit).pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(btn_area, text="🗑️ 刪除此筆紀錄", command=self.delete_sales_record).pack(side="left", fill="x", expand=True, padx=5)

        # 初始載入
        self.load_sales_records_for_edit()
        self.calculate_analysis_data()

    def load_sales_records_for_edit(self):
        """ 讀取銷售紀錄到列表 (確保顯示也是最新日期在最前) """
        for i in self.tree_sales_edit.get_children():
            self.tree_sales_edit.delete(i)
        
        try:
            if not os.path.exists(FILE_NAME): return
            df = pd.read_excel(FILE_NAME, sheet_name=SHEET_SALES)
            df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
            
            if df.empty: return

            # --- [排序優化] 顯示時也強制最新在最前 ---
            df['tmp_dt'] = pd.to_datetime(df['日期'], errors='coerce')
            df = df.sort_values(by=['tmp_dt', '訂單編號'], ascending=[False, False])
            # ----------------------------------------

            last_date, last_buyer = "", ""

            for idx, row in df.iterrows():
                raw_date = str(row.get('日期', '')) if pd.notna(row.get('日期')) else ""
                raw_buyer = str(row.get('買家名稱', '')) if pd.notna(row.get('買家名稱')) else ""
                item_name = str(row.get('商品名稱', ''))

                if raw_date == "" and raw_buyer == "" and item_name != "":
                    display_date, display_buyer = last_date, last_buyer
                else:
                    display_date, display_buyer = raw_date, raw_buyer
                    if raw_date != "": last_date = raw_date
                    if raw_buyer != "": last_buyer = raw_buyer

                # 取出數值
                qty = row.get('數量', 0)
                price = row.get('單價(售)', 0)
                fee = row.get('分攤手續費', 0)
                profit = row.get('總淨利', 0)
                margin = str(row.get('毛利率', '0.0')) + "%"

                self.tree_sales_edit.insert("", "end", text=str(idx), values=(
                    display_date, display_buyer, item_name, qty, price, fee, profit, margin
                ))
        except Exception as e:
            print(f"讀取歷史列表失敗: {e}")

    def on_sales_edit_select(self, event):
        """點擊列表時，將資料填入編輯框"""
        sel = self.tree_sales_edit.selection()
        if not sel: return
        
        item = self.tree_sales_edit.item(sel[0])
        idx = int(item['text']) # 取出原始 Excel Index
        self.var_edit_idx.set(idx)

        # 從 Excel 讀取完整資料 (因為 Treeview 只顯示部分欄位)
        try:
            df = pd.read_excel(FILE_NAME, sheet_name='銷售紀錄')
            row = df.iloc[idx]
            
            self.var_edit_date.set(str(row['日期']))
            self.var_edit_name.set(str(row['商品名稱']))
            self.var_edit_qty.set(int(row['數量']))
            self.var_edit_price.set(float(row['單價(售)']))
            self.var_edit_cost.set(float(row['單價(進)']))
            self.var_edit_fee.set(float(row['分攤手續費']))
            
            # 其他扣費不是每個訂單都有，需計算: 總銷售 - 總成本 - 淨利 - 手續費
            # 但 Excel 其實沒有直接存 "其他扣費金額"，而是 "扣費項目" 字串
            # 這裡我們為了簡化，不做反推，我們假設使用者修改的是「手續費」或「商品本身數據」
            # 若要精確，可以預設為 0，除非使用者自己有紀錄

            self.var_edit_deduct.set(0) 

        except Exception as e:
            messagebox.showerror("讀取錯誤", str(e))

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
        """ 設定分頁：包含字體設定與費率清單管理 """
        # 使用 Canvas 加上 Scrollbar 以防內容過多
        main_frame = ttk.Frame(self.tab_about, padding=20)
        main_frame.pack(fill="both", expand=True)

        # --- 第一區：顯示設定 ---
        font_frame = ttk.LabelFrame(main_frame, text="🎨 介面顯示設定", padding=15)
        font_frame.pack(fill="x", pady=10)
        ttk.Label(font_frame, text="字型大小 (10-20):").pack(side="left", padx=5)
        spin = ttk.Spinbox(font_frame, from_=10, to=20, textvariable=self.var_font_size, width=5, command=self.change_font_size)
        spin.pack(side="left", padx=5)
        ttk.Label(font_frame, text="(調整後需重啟或切換分頁生效)", foreground="gray").pack(side="left", padx=10)

        # --- 第二區：自訂費率管理 (核心功能) ---
        fee_mgmt_frame = ttk.LabelFrame(main_frame, text="💰 銷售費率清單管理 (儲存於 Excel)", padding=15)
        fee_mgmt_frame.pack(fill="both", expand=True, pady=10)

        # 左側清單
        list_frame = ttk.Frame(fee_mgmt_frame)
        list_frame.pack(side="left", fill="both", expand=True)
        
        self.fee_tree = ttk.Treeview(list_frame, columns=("名稱", "百分比"), show='headings', height=8)
        self.fee_tree.heading("名稱", text="費率名稱")
        self.fee_tree.heading("百分比", text="費率 (%)")
        self.fee_tree.column("百分比", width=80, anchor="center")
        self.fee_tree.pack(fill="both", expand=True)

        # 右側控制按鈕
        ctrl_frame = ttk.Frame(fee_mgmt_frame, padding=10)
        ctrl_frame.pack(side="right", fill="y")

        ttk.Label(ctrl_frame, text="名稱:").pack(anchor="w")
        self.ent_fee_name = ttk.Entry(ctrl_frame, width=15)
        self.ent_fee_name.pack(pady=5)

        ttk.Label(ctrl_frame, text="費率 (%):").pack(anchor="w")
        self.ent_fee_val = ttk.Entry(ctrl_frame, width=15)
        self.ent_fee_val.pack(pady=5)

        ttk.Button(ctrl_frame, text="➕ 新增/更新", command=self.action_add_custom_fee).pack(fill="x", pady=5)
        ttk.Button(ctrl_frame, text="🗑️ 刪除選取", command=self.action_delete_custom_fee).pack(fill="x", pady=5)
        ttk.Label(ctrl_frame, text="*修改後銷售頁面\n選單會同步更新", foreground="gray", font=("", 9)).pack(pady=10)

        # 載入初始費率資料
        self.refresh_fee_tree()

    def refresh_fee_tree(self):
        """ 刷新設定頁面的 Treeview 並同步更新銷售頁面的 Combobox (修正版：加入安全檢查) """
        
        # 【修正點 1】：檢查 fee_tree 是否已經被 setup_about_tab 建立
        if hasattr(self, 'fee_tree'):
            for i in self.fee_tree.get_children(): 
                self.fee_tree.delete(i)

        try:
            # 讀取 Excel 內的費率設定
            df = pd.read_excel(FILE_NAME, sheet_name=SHEET_CONFIG)
            fee_options = ["自訂手動輸入"]
            
            for _, row in df.iterrows():
                name, val = row['設定名稱'], row['費率百分比']
                
                # 【修正點 2】：只有當介面物件存在時才插入資料到列表
                if hasattr(self, 'fee_tree'):
                    self.fee_tree.insert("", "end", values=(name, val))
                
                # 組合出顯示在下拉選單的文字：例如 "一般賣家 (14.5%)"
                fee_options.append(f"{name} ({val}%)")
            
            # 同步更新銷售輸入頁面的 Combobox (如果它存在的話)
            if hasattr(self, 'combo_fee_rate'):
                self.combo_fee_rate['values'] = fee_options
        except Exception as e:
            print(f"讀取費率失敗: {e}")

    def action_add_custom_fee(self):
        name = self.ent_fee_name.get().strip()
        raw_val = self.ent_fee_val.get().strip()
        
        if not name or not raw_val:
            messagebox.showwarning("警告", "請輸入名稱與費率")
            return

        try:
            clean_val = raw_val.replace("%", "")
            val = float(clean_val)
        except ValueError:
            messagebox.showerror("錯誤", f"費率「{raw_val}」不是有效數字")
            return

        try:
            # --- [修正開始] 強大讀取邏輯 ---
            target_cols = ["設定名稱", "費率百分比"]
            try:
                # 嘗試讀取現有的設定
                df = pd.read_excel(FILE_NAME, sheet_name=SHEET_CONFIG)
                
                # 如果讀進來的欄位不對，強制重設
                if '設定名稱' not in df.columns:
                    df = pd.DataFrame(columns=target_cols)
            except Exception:
                # 如果分頁不存在或讀取失敗，建立新的
                df = pd.DataFrame(columns=target_cols)
            # --- [修正結束] ---

            # 如果名稱重複則更新，不重複則新增
            if not df.empty and name in df['設定名稱'].values:
                df.loc[df['設定名稱'] == name, '費率百分比'] = val
            else:
                new_row = pd.DataFrame([[name, val]], columns=target_cols)
                df = pd.concat([df, new_row], ignore_index=True)
            
            # 存回 Excel
            self._save_config_to_excel(df)
            self.refresh_fee_tree()
            
            # 清空輸入框
            self.ent_fee_name.delete(0, tk.END)
            self.ent_fee_val.delete(0, tk.END)
            messagebox.showinfo("成功", f"費率「{name}」已儲存。")
            
        except Exception as e:
            messagebox.showerror("儲存失敗", f"發生錯誤: {str(e)}")

    def action_delete_custom_fee(self):
        sel = self.fee_tree.selection()
        if not sel: return
        name = self.fee_tree.item(sel[0])['values'][0]
        
        try:
            df = pd.read_excel(FILE_NAME, sheet_name=SHEET_CONFIG)
            df = df[df['設定名稱'] != name]
            self._save_config_to_excel(df)
            self.refresh_fee_tree()
        except Exception as e: messagebox.showerror("錯誤", str(e))

    def _save_config_to_excel(self, df_config):
        """ 專門儲存設定分頁的輔助函式 (強化安全版) """
        try:
            # 1. 先讀取目前 Excel 裡所有的分頁，確保等等寫入時不會弄丟
            with pd.ExcelFile(FILE_NAME) as xls:
                sheet_names = xls.sheet_names
                all_data = {sn: pd.read_excel(xls, sheet_name=sn) for sn in sheet_names}
            
            # 2. 將我們要更新的「系統設定」放進資料字典中
            all_data[SHEET_CONFIG] = df_config

            # 3. 一次性全部寫回 Excel
            with pd.ExcelWriter(FILE_NAME, engine='openpyxl') as writer:
                for sn, df in all_data.items():
                    df.to_excel(writer, sheet_name=sn, index=False)
                    
        except PermissionError:
            messagebox.showerror("錯誤", "Excel 檔案被開啟中，請先關閉 Excel 再按儲存！")
        except Exception as e:
            messagebox.showerror("錯誤", f"存檔過程出錯: {str(e)}")

    # ---------------- 邏輯功能區 ----------------

    def action_track_delete_item(self):
        """ 刪除單一商品 (含表頭自動遞補邏輯) """
        sel = self.tree_track.selection()
        if not sel: return
        
        item = self.tree_track.item(sel[0])
        idx = int(item['text']) # 取得 Excel 中的列索引 (Row Index)
        order_id = str(item['values'][0]) # 取得訂單編號
        prod_name = str(item['values'][4])

        if not messagebox.askyesno("刪除商品", f"確定要從訂單 [{order_id}] 中\n刪除商品「{prod_name}」嗎？"):
            return

        try:
            # 讀取完整資料
            df = pd.read_excel(FILE_NAME, sheet_name=SHEET_TRACKING)
            
            # 確保訂單編號格式一致
            df['訂單編號'] = df['訂單編號'].astype(str).str.replace(r'\.0$', '', regex=True)
            
            # --- [核心邏輯] 表頭遞補檢查 ---
            # 1. 檢查要刪除的這行，是否包含重要資訊 (日期/買家)？
            #    (即檢查它是否為該訂單的第一行/表頭)
            is_header = pd.notna(df.at[idx, '日期']) or pd.notna(df.at[idx, '買家名稱'])
            
            if is_header:
                # 2. 找出同一張訂單的其他商品 (排除掉自己)
                # mask: 訂單編號相同 且 Index 不同
                mask_others = (df['訂單編號'] == order_id) & (df.index != idx)
                others_indices = df[mask_others].index.tolist()
                
                # 3. 如果還有其他商品，把表頭資訊移交給順位第一的商品
                if others_indices:
                    new_header_idx = others_indices[0] # 找到接班人
                    
                    # 需要移交的欄位
                    cols_to_inherit = ['日期', '交易平台', '買家名稱', '寄送方式', '取貨地點', '扣費項目']
                    
                    for col in cols_to_inherit:
                        # 把即將被刪除的資料 (idx) 複製給接班人 (new_header_idx)
                        df.at[new_header_idx, col] = df.at[idx, col]
                    
                    print(f"表頭已從 row {idx} 轉移至 row {new_header_idx}")

            # --- 刪除資料 ---
            df.drop(idx, inplace=True)
            
            # 寫回 Excel (保留其他分頁)
            self._save_all_sheets(df, SHEET_TRACKING)
            
            messagebox.showinfo("成功", "商品已刪除，訂單資料已自動修正。")
            self.load_tracking_data()

        except Exception as e:
            messagebox.showerror("錯誤", f"刪除失敗: {e}")

    def _get_full_order_info(self, df, order_id):
        """ 輔助函式：從同一編號中找出有資料的列，回傳表頭資訊字典 """
        # 確保 order_id 是乾淨的字串
        clean_id = str(order_id).replace("'", "")
        subset = df[df['訂單編號'].astype(str).str.contains(clean_id)]
        
        # 找尋第一個日期不為空的列
        headers = subset[subset['日期'].notna() & (subset['日期'] != "")]
        if not headers.empty:
            h = headers.iloc[0]
            return {
                '日期': h['日期'], '買家名稱': h['買家名稱'], 
                '交易平台': h['交易平台'], '寄送方式': h['寄送方式'], 
                '取貨地點': h['取貨地點']
            }
        return {}
    def action_track_return_item(self):
        """ 退貨單一商品 (含自動補足詳情與補位) """
        from tkinter import simpledialog
        sel = self.tree_track.selection()
        if not sel: return
        
        item = self.tree_track.item(sel[0])
        idx = int(item['text'])
        order_id = str(item['values'][0]).replace("'", "")
        prod_name = str(item['values'][4])

        reason = simpledialog.askstring("退貨", f"商品: {prod_name}\n請輸入退貨原因:", parent=self.root)
        if reason is None: return

        try:
            df_track = pd.read_excel(FILE_NAME, sheet_name=SHEET_TRACKING)
            df_track['訂單編號'] = df_track['訂單編號'].astype(str).str.replace(r'^\'', '', regex=True).str.replace(r'\.0$', '', regex=True)

            # 1. 取得這張訂單的完整資訊 (避免移走的是沒名字的那行)
            info = self._get_full_order_info(df_track, order_id)
            
            # 2. 準備要移走的這行資料，並補滿詳情
            row_to_move = df_track.loc[[idx]].copy()
            for col, val in info.items():
                row_to_move[col] = val
            row_to_move['備註'] = reason

            # 3. 處理追蹤表的補位邏輯
            is_header = pd.notna(df_track.at[idx, '日期']) and str(df_track.at[idx, '日期']) != ""
            if is_header:
                others = df_track[(df_track['訂單編號'] == order_id) & (df_track.index != idx)].index.tolist()
                if others:
                    new_h = others[0]
                    for col in info.keys(): df_track.at[new_h, col] = df_track.at[idx, col]

            # 4. 執行移動
            df_track.drop(idx, inplace=True)
            try: df_returns = pd.read_excel(FILE_NAME, sheet_name=SHEET_RETURNS)
            except: df_returns = pd.DataFrame()
            df_returns = pd.concat([df_returns, row_to_move], ignore_index=True)

            # 5. 存檔
            self._save_all_sheets_with_protect(df_track, SHEET_TRACKING, df_returns, SHEET_RETURNS)
            messagebox.showinfo("成功", f"商品「{prod_name}」已單獨移至退貨紀錄。")
            self.load_tracking_data(); self.load_returns_data()
        except Exception as e: messagebox.showerror("錯誤", str(e))

    def action_track_complete_order(self):
        """ 完成訂單 (整筆結案：移至銷售紀錄並自動排序) """
        sel = self.tree_track.selection()
        if not sel: return
        item = self.tree_track.item(sel[0])
        order_id = str(item['values'][0]).replace("'", "")

        if not messagebox.askyesno("結案確認", f"確定訂單 [{order_id}] 已完成？\n這將會把整筆訂單移至銷售紀錄並自動按日期排序。"):
            return

        try:
            # 1. 讀取追蹤表與歷史表
            df_track = pd.read_excel(FILE_NAME, sheet_name=SHEET_TRACKING)
            df_sales = pd.read_excel(FILE_NAME, sheet_name=SHEET_SALES)
            
            # 統一格式化編號
            df_track['訂單編號'] = df_track['訂單編號'].astype(str).str.replace(r'^\'', '', regex=True).str.replace(r'\.0$', '', regex=True)
            df_sales['訂單編號'] = df_sales['訂單編號'].astype(str).str.replace(r'^\'', '', regex=True).str.replace(r'\.0$', '', regex=True)

            # 2. 提取並補齊新結案的資料
            mask = df_track['訂單編號'] == order_id
            rows_to_finish = df_track[mask].copy()
            info = self._get_full_order_info(df_track, order_id)
            for col, val in info.items():
                rows_to_finish[col] = val

            # 3. 合併舊資料與新資料
            df_sales_combined = pd.concat([df_sales, rows_to_finish], ignore_index=True)

            # --- [核心排序邏輯] ---
            # 將日期轉為 datetime 格式以便精準排序
            df_sales_combined['tmp_date'] = pd.to_datetime(df_sales_combined['日期'], errors='coerce')
            
            # 排序：日期由新到舊 (descending)，訂單編號也由新到舊
            # 這樣可以確保「最新結案」或「日期最新」的永遠在 Excel 最上方
            df_sales_combined = df_sales_combined.sort_values(
                by=['tmp_date', '訂單編號'], 
                ascending=[False, False]
            ).drop(columns=['tmp_date']) # 刪除暫存的排序欄位
            # ----------------------

            # 4. 從追蹤表移除
            df_track_new = df_track[~mask]

            # 5. 存檔 (呼叫我們之前寫的保護編號函式)
            self._save_all_sheets_with_protect(df_track_new, SHEET_TRACKING, df_sales_combined, SHEET_SALES)
            
            messagebox.showinfo("成功", f"訂單 {order_id} 已結案並完成日期歸檔。")
            self.load_tracking_data()
            self.load_sales_records_for_edit() # 更新歷史列表
            self.calculate_analysis_data()    # 更新營收分析
            
        except Exception as e:
            messagebox.showerror("錯誤", f"結案失敗: {str(e)}")

    def _save_all_sheets_with_protect(self, df1, name1, df2, name2):
        """ 萬用存檔輔助：增加全自動排序與編號保護 """
        
        def process_df(df, name):
            # 保護編號 (加上單引號)
            if '訂單編號' in df.columns:
                df['訂單編號'] = df['訂單編號'].apply(lambda x: f"'{str(x).replace('\'','')}")
            
            # 如果是銷售紀錄或退貨紀錄，存檔前強制再排一次序
            if name in [SHEET_SALES, SHEET_RETURNS] and '日期' in df.columns:
                df['tmp_sort_dt'] = pd.to_datetime(df['日期'], errors='coerce')
                df = df.sort_values(by=['tmp_sort_dt', '訂單編號'], ascending=[False, False])
                df = df.drop(columns=['tmp_sort_dt'])
            return df

        df1 = process_df(df1, name1)
        df2 = process_df(df2, name2)

        with pd.ExcelWriter(FILE_NAME, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df1.to_excel(writer, sheet_name=name1, index=False)
            df2.to_excel(writer, sheet_name=name2, index=False)
            # 寫回其他沒變動的分頁... (其餘邏輯不變)
            for s in [SHEET_SALES, SHEET_TRACKING, SHEET_RETURNS, SHEET_PRODUCTS, SHEET_CONFIG]:
                if s != name1 and s != name2:
                    try:
                        temp_df = pd.read_excel(FILE_NAME, sheet_name=s)
                        temp_df.to_excel(writer, sheet_name=s, index=False)
                    except: pass
    

    def load_existing_tags(self, event=None):
        if not self.products_df.empty and "分類Tag" in self.products_df.columns:
            tags = self.products_df["分類Tag"].dropna().unique().tolist()
            self.combo_add_tag['values'] = tags
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
        search_term = self.var_search.get().lower()
        self.listbox_sales.delete(0, tk.END)
        if not self.products_df.empty:
            for index, row in self.products_df.iterrows():
                p_name = str(row['商品名稱'])
                p_tag = str(row['分類Tag']) if pd.notna(row['分類Tag']) else "無"
                try: p_stock = int(row['目前庫存'])
                except: p_stock = 0
                display_str = f"[{p_tag}] {p_name} (庫存: {p_stock})"
                
                if search_term in p_name.lower() or search_term in p_tag.lower():
                    self.listbox_sales.insert(tk.END, display_str)

    def on_sales_prod_select(self, event):
        selection = self.listbox_sales.curselection()
        if selection:
            display_str = self.listbox_sales.get(selection[0])
            try:
                temp = display_str.rsplit(" (庫存:", 1)[0]
                selected_name = temp.split("]", 1)[1].strip() if "]" in temp else temp
            except:
                selected_name = display_str 

            self.var_sel_name.set(selected_name)
            self.var_sel_qty.set(1)
            
            record = self.products_df[self.products_df['商品名稱'] == selected_name]
            if not record.empty:
                self.var_sel_cost.set(record.iloc[0]['預設成本'])
                try: stock = int(record.iloc[0]['目前庫存'])
                except: stock = 0
                self.var_sel_stock_info.set(str(stock)) 
                self.var_sel_price.set(0)

    def add_to_cart(self):
        name = self.var_sel_name.get()
        if not name: return
        try:
            qty = self.var_sel_qty.get()
            cost = self.var_sel_cost.get()
            price = self.var_sel_price.get()
            
            if qty <= 0: return

            current_stock = 0
            record = self.products_df[self.products_df['商品名稱'] == name]
            if not record.empty:
                try: current_stock = int(record.iloc[0]['目前庫存'])
                except: current_stock = 0

            if qty > current_stock:
                confirm = messagebox.askyesno("庫存不足警告", f"商品「{name}」目前庫存僅剩 {current_stock}，但您想賣出 {qty}。\n\n是否仍要加入清單 (超賣/預購)？")
                if not confirm:
                    return

            total_sales = price * qty
            total_cost = cost * qty
            self.cart_data.append({
                "name": name, "qty": qty, "unit_cost": cost, "unit_price": price,
                "total_sales": total_sales, "total_cost": total_cost
            })
            self.tree.insert("", "end", values=(name, qty, price, total_sales))
            self.update_totals()
            
            self.var_sel_name.set("")
            self.var_search.set("")
            self.var_sel_price.set(0)
            self.var_sel_qty.set(1)
            self.var_sel_stock_info.set("--")
            self.update_sales_prod_list()
            
        except ValueError: messagebox.showerror("錯誤", "數字格式錯誤")

    def remove_from_cart(self):
        sel = self.tree.selection()
        if not sel: return
        for item in sel:
            idx = self.tree.index(item)
            del self.cart_data[idx]
            self.tree.delete(item)
        self.update_totals()

    def on_fee_option_selected(self, event):
        selected_text = self.combo_fee_rate.get()
        match = re.search(r"\((\d+\.?\d*)%\)", selected_text)
        if match: self.update_totals()
        elif "自訂" in selected_text: self.combo_fee_rate.set("") 
        self.update_totals()

    def update_totals_event(self, event): self.update_totals()
    
    def update_totals(self):
        try:
            # 1. 總銷售額 (Revenue) 與 商品總進貨成本 (COGS)
            t_sales = sum(i['total_sales'] for i in self.cart_data)
            t_cost = sum(i['total_cost'] for i in self.cart_data)
            
            # 2. 解析平台手續費率 (例如 14.5%)
            raw_rate = self.var_fee_rate_str.get()
            rate = 0.0
            try: 
                rate = float(raw_rate)
            except ValueError:
                match = re.search(r"\((\d+\.?\d*)%\)", raw_rate)
                rate = float(match.group(1)) if match else 0.0

            # 3. 取得其他額外扣費 (廣告、補貼等)
            try: 
                extra = float(self.var_extra_fee.get())
            except: 
                extra = 0.0
            
            # 4. 計算平台收走的手續費
            platform_fee = (t_sales * (rate/100)) + extra
            
            # 5. 【關鍵修正：營業稅】
            # 直接以「銷售總額」乘以 5% 計算應繳稅金
            tax_amount = 0
            if hasattr(self, 'var_tax_enabled') and self.var_tax_enabled.get():
                tax_amount = t_sales * 0.05  # 正確：總額的 5%

            # 6. 計算預估入帳 (平台撥給您的金額 = 總額 - 平台費)
            income = t_sales - platform_fee

            # 7. 【關鍵修正：實收淨利】
            # 公式：總營收 - 平台費 - 營業稅 - 商品成本
            profit = t_sales - platform_fee - tax_amount - t_cost
            
            # 8. 更新介面顯示
            self.lbl_gross.config(text=f"總金額: ${t_sales:,.0f}")
            self.lbl_fee.config(text=f"平台扣費: -${platform_fee:,.1f}")
            self.lbl_income.config(text=f"預估入帳(平台撥款): ${income:,.1f}")

            if tax_amount > 0:
                # 這裡清楚標示營業稅是基於銷售額產生的
                self.lbl_profit.config(text=f"實收淨利: ${profit:,.1f} (營業稅: -${tax_amount:,.0f})")
            else:
                self.lbl_profit.config(text=f"實收淨利: ${profit:,.1f}")

            return t_sales, platform_fee
        except: 
            return 0, 0
        
    def submit_order(self):
        if not self.cart_data: return
        
        # --- 1. 資料清洗 ---
        def clean_text(text):
            if not text: return ""
            return text.replace("\n", "").replace("\r", "").strip()

        # --- 2. 讀取介面資料 ---
        if self.var_enable_cust.get():
            cust_name = clean_text(self.var_cust_name.get())
            cust_loc = clean_text(self.var_cust_loc.get())
            ship_method = self.var_ship_method.get()
            platform_name = self.var_platform.get()
        else:
            cust_name = ""
            cust_loc = ""
            ship_method = ""
            platform_name = ""
            
        date_str = self.var_date.get().strip()

        # --- 3. 生成訂單編號 ---
        now = datetime.now()
        order_id = now.strftime("%Y%m%d%H%M%S") 

        # --- 4. 計算金額 ---
        t_sales, t_fee = self.update_totals()
        fee_tag = self.var_fee_tag.get()
        try: extra_val = float(self.var_extra_fee.get())
        except: extra_val = 0
        if extra_val > 0 and not fee_tag: fee_tag = "其他"
        elif extra_val == 0: fee_tag = ""

        try:
            rows = []
            out_of_stock_warnings = [] 
            
            # 讀取商品資料以更新庫存
            df_prods_current = pd.read_excel(FILE_NAME, sheet_name='商品資料')

            for i, item in enumerate(self.cart_data):
                # 第一筆商品才顯示表頭，其餘留白
                if i == 0:
                    row_date = date_str
                    row_platform = platform_name 
                    row_buyer = cust_name
                    row_ship = ship_method
                    row_loc = cust_loc
                else:
                    row_date = ""
                    row_platform = "" 
                    row_buyer = ""
                    row_ship = ""
                    row_loc = ""

                ratio = item['total_sales'] / t_sales if t_sales > 0 else 0
                alloc_fee = t_fee * ratio
                net = item['total_sales'] - item['total_cost'] - alloc_fee
                margin_pct = (net / item['total_sales']) * 100 if item['total_sales'] > 0 else 0.0
                
                rows.append({
                    "訂單編號": order_id,
                    "日期": row_date, 
                    "買家名稱": row_buyer,     # 確保這裡變數是對的
                    "交易平台": row_platform,  # 確保這裡變數是對的
                    "寄送方式": row_ship, 
                    "取貨地點": row_loc,
                    "商品名稱": item['name'], 
                    "數量": item['qty'], 
                    "單價(售)": item['unit_price'], 
                    "單價(進)": item['unit_cost'],
                    "總銷售額": item['total_sales'], 
                    "總成本": item['total_cost'], 
                    "分攤手續費": round(alloc_fee, 2),
                    "扣費項目": fee_tag, 
                    "總淨利": round(net, 2),
                    "毛利率": round(margin_pct, 1)
                })

                # 庫存扣除
                prod_name = item['name']
                sold_qty = item['qty']
                idxs = df_prods_current[df_prods_current['商品名稱'] == prod_name].index
                if not idxs.empty:
                    target_idx = idxs[0]
                    raw_stock = df_prods_current.at[target_idx, '目前庫存']
                    try: current = int(raw_stock)
                    except: current = 0
                    new_stock = current - sold_qty
                    df_prods_current.at[target_idx, '目前庫存'] = new_stock
                    if new_stock <= 0:
                        out_of_stock_warnings.append(f"● {prod_name} (剩餘: {new_stock})")

            # --- 寫入 Excel (商品資料) ---
            with pd.ExcelWriter(FILE_NAME, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df_prods_current = df_prods_current.sort_values(by=['分類Tag', '商品名稱'], na_position='last')
                df_prods_current.to_excel(writer, sheet_name='商品資料', index=False)

            # --- 寫入 Excel (銷售紀錄) ---
            df_sales_new = pd.DataFrame(rows)
            

            # 在編號前面加上一個「'」(單引號)，這是 Excel 強制字串的暗號
            df_sales_new['訂單編號'] = df_sales_new['訂單編號'].apply(lambda x: f"'{x}")


            excel_columns_order = [
                "訂單編號", "日期", "買家名稱", "交易平台", "寄送方式", "取貨地點",
                "商品名稱", "數量", "單價(售)", "單價(進)", 
                "總銷售額", "總成本", "分攤手續費", "扣費項目", "總淨利", "毛利率"
            ]
            
            # 如果 DataFrame 有多餘或缺少欄位，這裡會自動對齊
            df_sales_new = df_sales_new[excel_columns_order]
            # ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★

# --- 寫入 Excel (商品資料分頁 - 更新庫存) ---
            with pd.ExcelWriter(FILE_NAME, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df_prods_current = df_prods_current.sort_values(by=['分類Tag', '商品名稱'], na_position='last')
                df_prods_current.to_excel(writer, sheet_name='商品資料', index=False)

            # --- 寫入 Excel (將新訂單寫入「訂單追蹤」而非銷售紀錄) ---
            df_sales_new = pd.DataFrame(rows)
            
            # 強制指定欄位順序 (確保 Excel 格式正確)
            excel_columns_order = [
                "訂單編號", "日期", "買家名稱", "交易平台", "寄送方式", "取貨地點",
                "商品名稱", "數量", "單價(售)", "單價(進)", 
                "總銷售額", "總成本", "分攤手續費", "扣費項目", "總淨利", "毛利率"
            ]
            df_sales_new = df_sales_new[excel_columns_order]

            with pd.ExcelWriter(FILE_NAME, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                # 【修正點】：將 sheet_name 改為 SHEET_TRACKING
                try:
                    df_ex = pd.read_excel(FILE_NAME, sheet_name=SHEET_TRACKING)
                    start_row = len(df_ex) + 1
                    header = False
                except:
                    # 如果分頁是空的或不存在
                    start_row = 0
                    header = True
                
                df_sales_new.to_excel(writer, sheet_name=SHEET_TRACKING, index=False, header=header, startrow=start_row)

            # --- 更新介面資料 ---
            self.products_df = df_prods_current
            self.update_sales_prod_list()
            self.update_mgmt_prod_list()
            
            # 【新增】：儲存後立刻重新讀取追蹤列表，讓緩衝區出現新資料
            self.load_tracking_data() 

            msg = f"訂單 {order_id} 已送至「訂單追蹤」緩衝區！\n庫存已預先扣除。"
            if out_of_stock_warnings:
                msg += "\n\n⚠️ 注意！以下商品已售完或庫存不足：\n" + "\n".join(out_of_stock_warnings)
            
            messagebox.showinfo("成功", msg)

            # 清空購物車欄位
            self.cart_data = []
            for i in self.tree.get_children(): self.tree.delete(i)
            self.update_totals()
            self.var_cust_name.set("")
            self.var_cust_loc.set("")
            self.var_sel_stock_info.set("--")

        except PermissionError: 
            messagebox.showerror("錯誤", "Excel 檔案未關閉，無法寫入！")
        except KeyError as e:
            messagebox.showerror("錯誤", f"欄位名稱不符，請檢查 Excel 標題: {str(e)}")
        except Exception as e: 
            messagebox.showerror("錯誤", f"發生未預期錯誤: {str(e)}")

    def update_mgmt_prod_list(self, event=None):
        search_term = self.var_mgmt_search.get().lower()
        self.listbox_mgmt.delete(0, tk.END)
        if not self.products_df.empty:
            for index, row in self.products_df.iterrows():
                p_name = str(row['商品名稱'])
                p_tag = str(row['分類Tag']) if pd.notna(row['分類Tag']) else "無"
                try: p_stock = int(row['目前庫存'])
                except: p_stock = 0
                display_str = f"[{p_tag}] {p_name} (庫存: {p_stock})"
                
                if search_term in p_name.lower() or search_term in p_tag.lower():
                    self.listbox_mgmt.insert(tk.END, display_str)

    def on_mgmt_prod_select(self, event):
        selection = self.listbox_mgmt.curselection()
        if selection:
            display_str = self.listbox_mgmt.get(selection[0])
            try:
                temp = display_str.rsplit(" (庫存:", 1)[0]
                selected_name = temp.split("]", 1)[1].strip() if "]" in temp else temp
            except:
                selected_name = display_str

            record = self.products_df[self.products_df['商品名稱'] == selected_name]
            if not record.empty:
                row = record.iloc[0]
                self.var_upd_name.set(row['商品名稱'])
                self.var_upd_tag.set(row['分類Tag'] if pd.notna(row['分類Tag']) else "")
                self.var_upd_cost.set(row['預設成本'])
                try: current_stock = int(row['目前庫存'])
                except: current_stock = 0
                self.var_upd_stock.set(current_stock)
                self.var_upd_time.set(row['最後更新時間'] if pd.notna(row['最後更新時間']) else "未知")

    def submit_new_product(self):
        name = self.var_add_name.get().strip()
        cost = self.var_add_cost.get()
        tag = self.var_add_tag.get().strip()
        stock = self.var_add_stock.get() 

        now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
        new_row = pd.DataFrame([{
            "分類Tag": tag, "商品名稱": name, "預設成本": cost, 
            "目前庫存": stock, "最後更新時間": now_str,
            "初始上架時間": now_str, "最後進貨時間": now_str  # 初始化
        }])
        
        if not name:
            messagebox.showwarning("警告", "請輸入商品名稱")
            return
        if name in self.products_df['商品名稱'].values:
            messagebox.showwarning("已存在", f"商品「{name}」已存在。\n請使用右側更新功能。")
            return
        try:
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
            new_row = pd.DataFrame([{"分類Tag": tag, "商品名稱": name, "預設成本": cost, "目前庫存": stock, "最後更新時間": now_str}])
            df_old = pd.read_excel(FILE_NAME, sheet_name='商品資料')
            df_updated = pd.concat([df_old, new_row], ignore_index=True)
            df_updated = df_updated.sort_values(by=['分類Tag', '商品名稱'], na_position='last')

            with pd.ExcelWriter(FILE_NAME, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                 df_updated.to_excel(writer, sheet_name='商品資料', index=False)
            self.products_df = df_updated
            self.update_sales_prod_list() 
            self.update_mgmt_prod_list()  
            messagebox.showinfo("成功", f"已新增：{name} (庫存: {stock})")
            self.var_add_name.set("")
            self.var_add_cost.set(0)
            self.var_add_stock.set(0)
        except PermissionError: messagebox.showerror("錯誤", "Excel 未關閉！")

    def submit_update_product(self):
        name = self.var_upd_name.get()
        if not name: return
        
        new_tag = self.var_upd_tag.get().strip()
        new_cost = self.var_upd_cost.get()
        new_stock = self.var_upd_stock.get() 
        
        try:
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
            # 1. 讀取商品資料
            df_prods = pd.read_excel(FILE_NAME, sheet_name=SHEET_PRODUCTS)
            
            idx = df_prods[df_prods['商品名稱'] == name].index
            if not idx.empty:
                old_stock = df_prods.loc[idx, '目前庫存'].values[0]
                
                # 補齊舊資料欄位 (相容性)
                if "初始上架時間" not in df_prods.columns: df_prods["初始上架時間"] = df_prods["最後更新時間"]
                if "最後進貨時間" not in df_prods.columns: df_prods["最後進貨時間"] = df_prods["最後更新時間"]

                # 補貨判定邏輯
                if new_stock > old_stock:
                    df_prods.loc[idx, '最後進貨時間'] = now_str
                    print(f"檢測到商品 {name} 補貨，更新進貨時間。")
                
                # 更新欄位
                df_prods.loc[idx, '分類Tag'] = new_tag
                df_prods.loc[idx, '預設成本'] = new_cost
                df_prods.loc[idx, '目前庫存'] = new_stock 
                df_prods.loc[idx, '最後更新時間'] = now_str
                
                # --- [修正：保護分頁的完整存檔邏輯] ---
                # 讀取其他分頁資料，避免被刪除
                try:
                    with pd.ExcelFile(FILE_NAME) as xls:
                        df_sales = pd.read_excel(xls, sheet_name=SHEET_SALES)
                        df_track = pd.read_excel(xls, sheet_name=SHEET_TRACKING)
                        df_ret = pd.read_excel(xls, sheet_name=SHEET_RETURNS)
                        df_cfg = pd.read_excel(xls, sheet_name=SHEET_CONFIG)
                except Exception as e:
                    # 如果讀取失敗 (例如有些分頁還沒產生)，則建立空白 DataFrame
                    df_sales = df_track = df_ret = df_cfg = pd.DataFrame()

                # 一口氣全部寫回
                with pd.ExcelWriter(FILE_NAME, engine='openpyxl') as writer:
                    df_prods.to_excel(writer, sheet_name=SHEET_PRODUCTS, index=False)
                    # 依序把舊有的資料寫回去，保護它們不消失
                    if not df_sales.empty: df_sales.to_excel(writer, sheet_name=SHEET_SALES, index=False)
                    if not df_track.empty: df_track.to_excel(writer, sheet_name=SHEET_TRACKING, index=False)
                    if not df_ret.empty: df_ret.to_excel(writer, sheet_name=SHEET_RETURNS, index=False)
                    if not df_cfg.empty: df_cfg.to_excel(writer, sheet_name=SHEET_CONFIG, index=False)
                # ------------------------------------
                
                self.products_df = self.load_products() 
                self.update_mgmt_prod_list()
                self.var_upd_time.set(now_str) 
                messagebox.showinfo("成功", f"商品「{name}」資訊已更新！")
                
        except PermissionError: 
            messagebox.showerror("錯誤", "Excel 檔案未關閉，無法寫入！")
        except Exception as e:
            messagebox.showerror("錯誤", f"更新失敗: {e}")

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


if __name__ == "__main__":
    root = tk.Tk()
    style = ttk.Style()
    try:
        style.theme_use('vista') 
    except:
        pass 
    app = SalesApp(root)
    root.mainloop()

