#shopee-oms 4.5 完整版

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

SHEET_PURCHASES = '進貨紀錄'
SHEET_PUR_TRACKING = '進貨追蹤'
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
        """上傳檔案到指定資料夾，並維持最多 15 筆備份"""
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
            
            if len(items) > 15:
                # 取得第 15 筆之後的所有檔案 (即最舊的檔案們)
                files_to_delete = items[15:] 
                for old_file in files_to_delete:
                    file_id = old_file.get('id')
                    try:
                        self.service.files().delete(fileId=file_id).execute()
                        print(f"自動清理舊備份: {old_file.get('name')}")
                    except Exception as delete_error:
                        print(f"刪除舊檔失敗: {delete_error}")

            return True, f"備份成功！\n雲端檔名: {file_name}\n(系統已自動保留最新 30 筆紀錄)"
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
        self.root.title("蝦皮/網拍進銷存系統 (V4.0 完整版)")
        self.root.geometry("1280x850") 
        self.var_shop_name = tk.StringVar(value="商店") # 預設名稱


          # 可選擇隱藏的欄位(不能隱藏): 商品名稱, 預設成本, 目前庫存

        self.show_fields = {
            "商品編號": tk.BooleanVar(value=True),
            "分類Tag": tk.BooleanVar(value=True),
            "安全庫存": tk.BooleanVar(value=True),
            "商品連結": tk.BooleanVar(value=True),
            "商品備註": tk.BooleanVar(value=True)
        }

        # --- 字型設定 ---
        self.default_font_size = 11
        self.style = ttk.Style()
        self.setup_fonts(self.default_font_size)

        self.drive_manager = GoogleDriveSync()

        # --- 變數初始化 ---
        self.fee_lookup = {}
        self.var_ship_payer = tk.StringVar(value="買家付") # 預設買家付
        self.var_tax_type = tk.StringVar(value="無")
        self.var_ship_fee = tk.DoubleVar(value=0.0)
        self.var_after_type = tk.StringVar()  # 售後類型 (補寄/補貼/換貨/保固)
        self.var_extra_fee = tk.DoubleVar(value=0.0)     # 折扣/額外扣費
        self.var_after_cost = tk.DoubleVar(value=0.0) # 額外支出金額
        self.var_after_remark = tk.StringVar() # 售後備註
        self.var_view_after_status = tk.StringVar(value="無售後紀錄")



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

      
    
   

    def setup_fonts(self, size):
        default_font = font.nametofont("TkDefaultFont")
        default_font.configure(family="微軟正黑體", size=size)
        
        text_font = font.nametofont("TkTextFont")
        text_font.configure(family="微軟正黑體", size=size)

        self.style.configure(".", font=("微軟正黑體", size))
        # 關鍵：行高必須隨字體大小縮放，通常是字體大小的 2.5 到 3 倍
        self.style.configure("Treeview", rowheight=int(size * 2.5)) 
        self.style.configure("Treeview.Heading", font=("微軟正黑體", size, "bold"))
        self.style.configure("TLabelframe.Label", font=("微軟正黑體", size, "bold"))

    def change_font_size(self, event=None):
        try:
            new_size = int(self.var_font_size.get())
            # 1. 更新全局字體定義
            self.setup_fonts(new_size)
            
            # 2. 強制更新特定「標準 Tk」元件 (Listbox, Text, Entry)
            # 這些元件不會自動跟隨 ttk 樣式變化，需要手動配置
            new_font = ("微軟正黑體", new_size)
            bold_font = ("微軟正黑體", new_size, "bold")

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

            print(f"系統字體已統一調整為: {new_size}")
        except Exception as e:
            print(f"字體調整失敗: {e}")


    def load_system_settings(self):
        """ 從 Excel 載入永久保存的系統設定 """
        try:
            if os.path.exists(FILE_NAME):
                df_cfg = pd.read_excel(FILE_NAME, sheet_name=SHEET_CONFIG)
                # 尋找商家名稱設定
                shop_row = df_cfg[df_cfg['設定名稱'] == "SYSTEM_SHOP_NAME"]
                if not shop_row.empty:
                    # 我們將店名存在「費率百分比」這一欄（雖然欄名不符，但為了不更動 Excel 結構）
                    # 或者妳可以檢查是否有『參數值』這一欄，若無則彈性處理
                    saved_name = str(shop_row.iloc[0]['費率百分比'])
                    self.var_shop_name.set(saved_name)
        except Exception as e:
            print(f"載入商家名稱失敗: {e}")


    def save_system_settings(self):
        """ 將商家名稱永久存入 Excel """
        shop_name = self.var_shop_name.get().strip()
        if not shop_name:
            messagebox.showwarning("警告", "商家名稱不能為空")
            return

        try:
            # 1. 讀取現有設定
            df_cfg = pd.read_excel(FILE_NAME, sheet_name=SHEET_CONFIG)
            
            # 2. 更新或新增商家名稱列
            if "SYSTEM_SHOP_NAME" in df_cfg['設定名稱'].values:
                df_cfg.loc[df_cfg['設定名稱'] == "SYSTEM_SHOP_NAME", '費率百分比'] = shop_name
            else:
                new_row = pd.DataFrame([["SYSTEM_SHOP_NAME", shop_name, 0]], columns=df_cfg.columns)
                df_cfg = pd.concat([df_cfg, new_row], ignore_index=True)

            # 3. 使用萬用引擎存檔，保護其他分頁
            if self._universal_save({SHEET_CONFIG: df_cfg}):
                messagebox.showinfo("成功", "商家設定已永久保存！")
        except Exception as e:
            messagebox.showerror("錯誤", f"儲存設定失敗: {e}")


    def check_excel_file(self):
            cols_sales = ["訂單編號", "日期", "買家名稱", "交易平台", "寄送方式", "取貨地點", 
                      "商品名稱", "數量", "單價(售)", "單價(進)", "總銷售額", "總成本", 
                      "分攤手續費", "扣費項目", "總淨利", "毛利率", "稅額"]
            
            cols_purchase = [
            "進貨單號", "採購日期", "入庫日期", "供應商", "物流追蹤", 
            "商品名稱", "數量", "進貨單價", "進貨總額", "進項稅額", "備註"
        ]

            cols_prods = ["商品編號","分類Tag", "商品名稱", "預設成本", "目前庫存", 
                            "最後更新時間", "初始上架時間", "最後進貨時間", "安全庫存",
                            "商品連結", "商品備註"]

            cols_config = ["設定名稱", "費率百分比", "固定金額"]

            default_fees = [
                ["蝦皮一般 方案一", 14.5, 0],
                ["蝦皮活動 方案二", 8.0, 60], # 8% + 60元
                ]
            

            if not os.path.exists(FILE_NAME):
                try:
                    
                    with pd.ExcelWriter(FILE_NAME, engine='openpyxl') as writer:
                        pd.DataFrame(columns=cols_sales).to_excel(writer, sheet_name=SHEET_SALES, index=False)
                        pd.DataFrame(columns=cols_sales).to_excel(writer, sheet_name=SHEET_TRACKING, index=False)
                        pd.DataFrame(columns=cols_sales).to_excel(writer, sheet_name=SHEET_RETURNS, index=False)
                        # 建立進貨分頁
                        pd.DataFrame(columns=cols_purchase).to_excel(writer, sheet_name=SHEET_PURCHASES, index=False)         

                        df_prods = pd.DataFrame(columns=cols_prods)
                        df_prods.to_excel(writer, sheet_name=SHEET_PRODUCTS, index=False)
                        pd.DataFrame(columns=cols_config).to_excel(writer, sheet_name=SHEET_CONFIG, index=False)
                except Exception as e:
                    messagebox.showerror("錯誤", f"無法建立 Excel: {e}")
            else:
                # 檢查是否缺少進貨分頁
                try:
                    with pd.ExcelWriter(FILE_NAME, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                        if SHEET_PURCHASES not in writer.book.sheetnames:
                            pd.DataFrame(columns=cols_purchase).to_excel(writer, sheet_name=SHEET_PURCHASES, index=False)
                            pd.DataFrame(columns=cols_purchase).to_excel(writer, sheet_name=SHEET_PUR_TRACKING, index=False)
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
            if "商品編號" not in df.columns:
                df["商品編號"] = "" # 若沒有編號欄位，自動補空字串
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
        
        self.tab_about = ttk.Frame(tab_control)
        self.tab_purchase = ttk.Frame(tab_control) # [新增] 進貨分頁
        self.tab_pur_tracking = ttk.Frame(tab_control)
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
        """ 建立進貨管理介面 (優化後的搜尋清單版) """
        current_size = int(self.var_font_size.get())
        self.pur_cart_data = []
        self.var_pur_date = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        self.var_pur_supplier = tk.StringVar()
        self.var_pur_sel_name = tk.StringVar()
        self.var_pur_sel_qty = tk.IntVar(value=1)
        self.var_pur_sel_cost = tk.DoubleVar(value=0.0)
        self.var_pur_tax_enabled = tk.BooleanVar(value=False)

        paned = ttk.PanedWindow(self.tab_purchase, orient=tk.HORIZONTAL)
        paned.pack(fill="both", expand=True, padx=10, pady=10)

        # --- 左側：輸入資訊 ---
        left_frame = ttk.LabelFrame(paned, text="1. 填寫採購單", padding=10)
        paned.add(left_frame, weight=1)

        ttk.Label(left_frame, text="採購日期:").pack(anchor="w")
        ttk.Entry(left_frame, textvariable=self.var_pur_date).pack(fill="x", pady=2)

        ttk.Label(left_frame, text="供應商:").pack(anchor="w")
        ttk.Entry(left_frame, textvariable=self.var_pur_supplier).pack(fill="x", pady=2)
        
        ttk.Separator(left_frame).pack(fill="x", pady=10)
        
        # --- 改良版搜尋區 ---
        ttk.Label(left_frame, text="🔍 搜尋商品名稱:", font=("微軟正黑體", current_size, "bold")).pack(anchor="w")
        self.ent_pur_search = ttk.Entry(left_frame)
        self.ent_pur_search.pack(fill="x", pady=2)
        self.ent_pur_search.bind('<KeyRelease>', self.update_pur_prod_list_by_search)

        # 商品列表框
        list_frame_pur = ttk.Frame(left_frame)
        list_frame_pur.pack(fill="both", expand=True, pady=5)
        self.list_pur_prod = tk.Listbox(list_frame_pur, height=6, font=("微軟正黑體", current_size))
        self.list_pur_prod.pack(side="left", fill="both", expand=True)
        
        sc_pur = ttk.Scrollbar(list_frame_pur, orient="vertical", command=self.list_pur_prod.yview)
        self.list_pur_prod.configure(yscrollcommand=sc_pur.set)
        sc_pur.pack(side="right", fill="y")
        self.list_pur_prod.bind('<<ListboxSelect>>', self.on_pur_list_select)

        # 顯示當前選中 (唯讀)
        ttk.Label(left_frame, text="已選商品:").pack(anchor="w")
        ttk.Entry(left_frame, textvariable=self.var_pur_sel_name, state="readonly", foreground="blue").pack(fill="x", pady=2)

        # 金額與數量
        f_row = ttk.Frame(left_frame)
        f_row.pack(fill="x", pady=5)
        
        ttk.Label(f_row, text="進貨單價:").grid(row=0, column=0, sticky="w")
        ttk.Entry(f_row, textvariable=self.var_pur_sel_cost, width=12).grid(row=0, column=1, padx=5)
        
        ttk.Label(f_row, text="數量:").grid(row=0, column=2, sticky="w")
        ttk.Entry(f_row, textvariable=self.var_pur_sel_qty, width=8).grid(row=0, column=3, padx=5)
        
        ttk.Checkbutton(left_frame, text="此筆有含 5% 營業稅", variable=self.var_pur_tax_enabled).pack(anchor="w", pady=5)
        
        ttk.Button(left_frame, text="➕ 加入採購清單", command=self.add_to_pur_cart).pack(fill="x", pady=10)

        # --- 右側：購物車預覽 ---
        right_frame = ttk.LabelFrame(paned, text="2. 本次採購明細 (待送出)", padding=10)
        paned.add(right_frame, weight=2)
        
        pur_cols = ("商品名稱", "採購數量", "進貨單價", "進項稅額", "小計(含稅)")
        self.tree_pur_cart = ttk.Treeview(right_frame, columns=pur_cols, show='headings', height=10)
        for c in pur_cols:
            self.tree_pur_cart.heading(c, text=c)
            # 根據內容調整寬度
            if c == "商品名稱":
                self.tree_pur_cart.column(c, width=180, anchor="w") # 商品名稱給寬一點
            elif c == "小計(含稅)":
                self.tree_pur_cart.column(c, width=100, anchor="e")
            else:
                self.tree_pur_cart.column(c, width=80, anchor="center")
                
        self.tree_pur_cart.pack(fill="both", expand=True)
        
        btn_area = ttk.Frame(right_frame)
        btn_area.pack(fill="x", pady=10)
        ttk.Button(btn_area, text="➖ 移除項目", command=self.remove_from_pur_cart).pack(side="left", padx=5)
        ttk.Button(btn_area, text="🚀 送出採購單", command=self.submit_purchase_batch).pack(side="right", padx=5)

        # 初始化載入清單
        self.update_pur_prod_list()

    def update_pur_prod_list(self):
        """ 初始化/重新載入進貨商品清單 """
        if hasattr(self, 'list_pur_prod') and not self.products_df.empty:
            self.list_pur_prod.delete(0, tk.END)
            for name in self.products_df['商品名稱'].tolist():
                self.list_pur_prod.insert(tk.END, name)

    def update_pur_prod_list_by_search(self, event=None):
        """ 進貨搜尋框：顯示 [編號] 商品名稱，並支援編號搜尋 """
        query = self.ent_pur_search.get().lower()
        self.list_pur_prod.delete(0, tk.END)
        
        if not self.products_df.empty:
            for index, row in self.products_df.iterrows():
                p_name = str(row['商品名稱'])
                sku = str(row.get('商品編號', ''))
                
                # 處理編號顯示邏輯
                sku_display = f"[{sku}] " if sku and sku != "nan" and sku.strip() != "" else ""
                
                # 搜尋邏輯：檢查 關鍵字 是否出現在 名稱 或 編號 中
                if query in p_name.lower() or query in sku.lower():
                    self.list_pur_prod.insert(tk.END, f"{sku_display}{p_name}")

    def on_pur_list_select(self, event):
        selection = self.list_pur_prod.curselection()
        if selection:
            raw_text = self.list_pur_prod.get(selection[0])
            
            # --- 拆解邏輯 ---
            # 如果文字裡面有 "]"，名稱通常在最後一個 "]" 之後
            if "]" in raw_text:
                selected_name = raw_text.split("]")[-1].strip()
            else:
                selected_name = raw_text
                
            self.var_pur_sel_name.set(selected_name)

            record = self.products_df[self.products_df['商品名稱'] == selected_name]
            if not record.empty:
                current_cost = record.iloc[0]['預設成本']
                self.var_pur_sel_cost.set(current_cost)



    def submit_purchase_batch(self):
        """ 提交採購：確保欄位名稱與 Excel 標題完全一致 """
        if not self.pur_cart_data: return
        supplier = self.var_pur_supplier.get().strip()
        pur_id = "I" + datetime.now().strftime("%Y%m%d%H%M%S")
        
        try:
            with pd.ExcelFile(FILE_NAME) as xls:
                df_history = pd.read_excel(xls, sheet_name=SHEET_PURCHASES)
                df_tracking = pd.read_excel(xls, sheet_name=SHEET_PUR_TRACKING)
            
            new_entries = []
            for item in self.pur_cart_data:
                # 注意：這裡的 Key 必須與 Excel 標題一致
                new_entries.append({
                    "進貨單號": f"'{pur_id}",
                    "採購日期": self.var_pur_date.get(),
                    "入庫日期": "",  
                    "供應商": supplier if supplier else "未填",
                    "物流追蹤": "待發貨", # <--- 這裡要固定叫做「物流追蹤」
                    "商品名稱": item['name'],
                    "數量": item['qty'],
                    "進貨單價": item['cost'],
                    "進貨總額": item['total'],
                    "進項稅額": item['tax'],
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
            
        print("已從暫存清單移除商品")


    def load_purchase_tracking(self):
        """ 載入待收貨清單：精準填入 8 個欄位資料 """
        # 清空 UI 列表
        for i in self.tree_pur_track.get_children(): 
            self.tree_pur_track.delete(i)
            
        try:
            if not os.path.exists(FILE_NAME): return
            df = pd.read_excel(FILE_NAME, sheet_name=SHEET_PUR_TRACKING)
            if df.empty: return

            for idx, row in df.iterrows():
                # 按順序填入 values:
                # 0:單號, 1:供應商, 2:商品名稱, 3:數量, 4:單價, 5:稅額, 6:運費, 7:物流
                self.tree_pur_track.insert("", "end", text=str(idx), values=(
                    str(row.get('進貨單號', '')).replace("'", ""),
                    row.get('供應商', '未填'),
                    row.get('商品名稱', '未知'),
                    row.get('數量', 0),
                    row.get('進貨單價', 0),
                    row.get('海關稅金', 0), # 稅金放在索引 5
                    row.get('分攤運費', 0), # 運費放在索引 6
                    row.get('物流追蹤', '待發貨') # 物流放在索引 7
                ))
        except Exception as e:
            print(f"載入追蹤清單出錯: {e}")

    def setup_pur_tracking_tab(self):
        """ 建立在途貨物追蹤：增加獨立的運費欄位 """
        frame = self.tab_pur_tracking
        
        top_frame = ttk.Frame(frame, padding=5)
        top_frame.pack(fill="x")
        ttk.Label(top_frame, text="🚚 運輸中貨物管理 (可補填 稅金、運費、物流單號)", foreground="blue").pack(side="left")
        ttk.Button(top_frame, text="🔄 刷新列表", command=self.load_purchase_tracking).pack(side="right")

        # --- 更新欄位：增加到 8 個 ---
        cols_pur_track = ("單號", "供應商", "商品名稱", "數量", "單價", "稅額", "運費", "物流狀態/單號")
        
        self.tree_pur_track = ttk.Treeview(frame, columns=cols_pur_track, show='headings', height=15)
        
        for c in cols_pur_track:
            self.tree_pur_track.heading(c, text=c)
            # 針對不同欄位設定寬度
            if c == "商品名稱":
                self.tree_pur_track.column(c, width=180, anchor="w")
            elif c in ["稅額", "運費"]:
                self.tree_pur_track.column(c, width=70, anchor="center")
            elif c == "物流狀態/單號":
                self.tree_pur_track.column(c, width=150, anchor="center")
            else:
                self.tree_pur_track.column(c, width=80, anchor="center")
        
        self.tree_pur_track.pack(fill="both", expand=True, padx=10)

        # 下方按鈕區不變...
        btn_ctrl = ttk.Frame(frame, padding=10)
        btn_ctrl.pack(fill="x")
        ttk.Button(btn_ctrl, text="✏️ 補充運費/稅金/物流號", command=self.action_update_pur_logistics).pack(side="left", padx=5)
        ttk.Button(btn_ctrl, text="✅ 確認收貨入庫", command=self.action_confirm_inbound).pack(side="left", padx=5)
        ttk.Button(btn_ctrl, text="❌ 標記遺失/取消", command=self.action_cancel_purchase).pack(side="left", padx=5)

        self.load_purchase_tracking()


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
        """ 核心分析邏輯 V4.2:修正消失問題，並列出近 10 日明細 """
        if not hasattr(self, 'tree_time_stats') or not hasattr(self, 'tree_prod_stats'): return
        
        # 1. 清空舊介面
        for i in self.tree_time_stats.get_children(): self.tree_time_stats.delete(i)
        for i in self.tree_prod_stats.get_children(): self.tree_prod_stats.delete(i)
        
        if not os.path.exists(FILE_NAME): return

        try:
            # 2. 一次性讀取銷售與商品分頁
            with pd.ExcelFile(FILE_NAME) as xls:
                df_sales = pd.read_excel(xls, sheet_name=SHEET_SALES)
                df_prods = pd.read_excel(xls, sheet_name=SHEET_PRODUCTS)

            if df_sales.empty: return

            # --- [關鍵步驟 A]：清洗資料與填充留白 ---
            # 將完全空白的儲存格轉為真正的空值 (NaN)，ffill 才會生效
            df_sales = df_sales.replace(r'^\s*$', pd.NA, regex=True)
            
            # 針對 Excel 美觀留白處進行向下填充
            fill_cols = ['訂單編號', '日期', '買家名稱', '交易平台']
            for col in fill_cols:
                if col in df_sales.columns:
                    df_sales[col] = df_sales[col].ffill()

            # 只保留「有商品名稱」的列，避免算到 Excel 底部的空行
            df_sales = df_sales.dropna(subset=['商品名稱'])

            # 轉換數字欄位，出錯則填 0
            num_cols = ['總銷售額', '總成本', '數量', '總淨利']
            for col in num_cols:
                if col in df_sales.columns:
                    df_sales[col] = pd.to_numeric(df_sales[col], errors='coerce').fillna(0)

            # 處理日期 (強制轉換，失敗的會變成 NaT)
            df_sales['日期'] = pd.to_datetime(df_sales['日期'], errors='coerce')
            
            # 處理毛利率 (轉換為數字方便平均運算)
            df_sales['毛利率_數值'] = pd.to_numeric(df_sales['毛利率'].astype(str).str.replace('%', ''), errors='coerce').fillna(0)

            # --- [關鍵步驟 B]：左側月份與每日統計 (修正消失點) ---
            # 建立一個乾淨的有日期的 DataFrame 用於時間統計
            df_time = df_sales.dropna(subset=['日期']).copy()

            if not df_time.empty:
                # 1. 月份匯總
                df_time['月份'] = df_time['日期'].dt.strftime('%Y-%m')
                monthly_group = df_time.groupby('月份').agg({
                    '總銷售額': 'sum',
                    '總淨利': 'sum',
                    '訂單編號': 'nunique', # 計算不重複單數
                    '數量': 'sum'
                }).sort_index(ascending=False)

                # 更新頂部看板數字 (本月)
                latest_m = monthly_group.index[0]
                self.lbl_month_sales.config(text=f"本月({latest_m}) 營收: ${monthly_group.iloc[0]['總銷售額']:,.0f}")
                self.lbl_month_profit.config(text=f"本月({latest_m}) 淨利: ${monthly_group.iloc[0]['總淨利']:,.0f}")

                # 填入左側表格 (月份部分)
                for m, row in monthly_group.iterrows():
                    self.tree_time_stats.insert("", "end", values=(
                        f"{m} (月)", 
                        f"${row['總銷售額']:,.0f}", 
                        f"${row['總淨利']:,.0f}", 
                        f"{int(row['訂單編號'])} 單"
                    ))

                # 插入分隔線
                self.tree_time_stats.insert("", "end", values=("--- 近10日明細 ---", "", "", ""))

                # 2. 每日明細 (修正為近 10 日)
                df_time['日期字串'] = df_time['日期'].dt.strftime('%Y-%m-%d')
                daily_group = df_time.groupby('日期字串').agg({
                    '總銷售額': 'sum',
                    '總淨利': 'sum',
                    '訂單編號': 'nunique'
                }).sort_index(ascending=False).head(10) # 這裡改為 10

                for d, row in daily_group.iterrows():
                    self.tree_time_stats.insert("", "end", values=(
                        d, 
                        f"${row['總銷售額']:,.0f}", 
                        f"${row['總淨利']:,.0f}", 
                        f"{int(row['訂單編號'])} 單"
                    ))

            # --- [關鍵步驟 C]：右側商品排行 (銷售速度) ---
            try:
                # 1. 統一清洗名稱 (避免空格造成 Map 失敗)
                df_prods['商品名稱'] = df_prods['商品名稱'].astype(str).str.strip()
                df_sales['商品名稱'] = df_sales['商品名稱'].astype(str).str.strip()

                # 2. 處理商品分頁的上架時間
                start_col = "初始上架時間"
                if start_col not in df_prods.columns:
                    df_prods[start_col] = pd.NA
                
                # 強制轉換日期格式
                df_prods[start_col] = pd.to_datetime(df_prods[start_col], errors='coerce')
                
                # 建立名稱對應上架日的地圖
                start_date_map = df_prods.set_index('商品名稱')[start_col].to_dict()

                # 3. 備援邏輯：從銷售紀錄抓取「每個商品的第一筆成交日」
                # 這是為了預防 Excel 上架時間漏填
                first_sale_map = df_sales.groupby('商品名稱')['日期'].min().to_dict()

                # 4. 聚合銷售數據
                prod_group = df_sales.groupby('商品名稱').agg({
                    '毛利率_數值': 'mean',
                    '總淨利': 'sum',
                    '數量': 'sum'
                }).reset_index()

                now = pd.Timestamp.now()

                def calculate_velocity(row):
                    p_name = row['商品名稱']
                    total_qty = row['數量']
                    
                    # 優先序 A: Excel 填寫的初始上架時間
                    st_date = start_date_map.get(p_name)
                    
                    # 優先序 B: 若 A 缺失，使用該商品在系統中的第一筆銷售日
                    if pd.isna(st_date):
                        st_date = first_sale_map.get(p_name)
                    
                    # 優先序 C: 若連銷售日都抓不到(理論上不會)，預設為 30 天前 (避免暴增)
                    if pd.isna(st_date):
                        st_date = now - pd.Timedelta(days=30)

                    # 計算天數差 (精確到小數點)
                    delta = now - st_date
                    days_diff = delta.total_seconds() / 86400 # 轉換為總天數
                    
                    # 限制最小分母為 1 天 (避免剛上架 1 小時賣 1 個就被算成時速 24 也就是日速 24)
                    velocity = total_qty / max(days_diff, 1)
                    return round(velocity, 2)

                # 執行速度計算
                prod_group['velocity'] = prod_group.apply(calculate_velocity, axis=1)

                # 5. 排序邏輯
                sort_mode = self.var_prod_sort_by.get()
                sort_map = {
                    "平均毛利率": '毛利率_數值', 
                    "總銷量排行": '數量', 
                    "總獲利排行": '總淨利', 
                    "銷售速度排行": 'velocity'
                }
                prod_group = prod_group.sort_values(sort_map.get(sort_mode, 'velocity'), ascending=False)

                # 6. 填入右側表格
                for _, row in prod_group.iterrows():
                    self.tree_prod_stats.insert("", "end", values=(
                        row['商品名稱'], 
                        f"{row['毛利率_數值']:.1f}%", 
                        f"${row['總淨利']:,.0f}", 
                        int(row['數量']), 
                        f"{row['velocity']} 件/日"
                    ))

            except Exception as e:
                print(f"商品排行計算出錯: {e}")

        except Exception as e:
            import traceback
            print("分析功能報錯：")
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
        """ 建立採購建議與評估分頁 """
        frame = self.tab_procurement # 記得在 create_tabs 加入此分頁
        
        # --- 頂部：評估參數控制區 ---
        ctrl_frame = ttk.LabelFrame(frame, text="⚙️ 採購評估參數 (手動微調)", padding=10)
        ctrl_frame.pack(fill="x", padx=10, pady=5)

        # 參數 A: 銷售速度閾值 (只看每天賣超過 X 件的商品)
        ttk.Label(ctrl_frame, text="1. 銷售速度大於:").grid(row=0, column=0, padx=5)
        self.var_filter_velocity = tk.DoubleVar(value=0.1) # 預設每天賣 0.1 件才報警
        ttk.Entry(ctrl_frame, textvariable=self.var_filter_velocity, width=8).grid(row=0, column=1)
        ttk.Label(ctrl_frame, text="件/日").grid(row=0, column=2, padx=5)

        # 參數 B: 安全庫存加權 (如果您想在旺季多備一點貨，可以設為 1.5 倍)
        ttk.Label(ctrl_frame, text="2. 安全庫存係數:").grid(row=0, column=3, padx=15)
        self.var_safety_multiplier = tk.DoubleVar(value=1.0)
        ttk.Entry(ctrl_frame, textvariable=self.var_safety_multiplier, width=8).grid(row=0, column=4)
        
        ttk.Label(ctrl_frame, text="3. 預計備貨天數:").grid(row=0, column=6, padx=15)
        self.var_days_to_cover = tk.IntVar(value=30) # 預設一次買 30 天份
        ttk.Entry(ctrl_frame, textvariable=self.var_days_to_cover, width=8).grid(row=0, column=7)
        ttk.Label(ctrl_frame, text="天").grid(row=0, column=8, padx=5)


        ttk.Button(ctrl_frame, text="🔄 重新生成採購建議", command=self.generate_procurement_report).grid(row=0, column=9, padx=20)

        # --- 中間：建議清單 ---
        list_frame = ttk.LabelFrame(frame, text="📋 建議採購商品清單 (基於銷售表現與庫存缺口)", padding=10)
        list_frame.pack(fill="both", expand=True, padx=10, pady=5)

        cols = ("品名", "目前庫存", "安全值", "銷售速度", "缺貨狀態", "建議採購量")
        self.tree_procure = ttk.Treeview(list_frame, columns=cols, show='headings', height=20)
        
        # 設定欄位 ID 順序與寬度
        widths = {"品名": 200, "目前庫存": 80, "安全值": 80, "銷售速度": 100, "缺貨狀態": 100, "建議採購量": 120}
        for c in cols:
            self.tree_procure.heading(c, text=c)
            self.tree_procure.column(c, width=widths[c], anchor="center")
        
        self.tree_procure.pack(fill="both", expand=True)
        
        # 狀態標記 (紅字)
        self.tree_procure.tag_configure('urgent', foreground='red')
        self.tree_procure.tag_configure('warning', foreground='orange')

    def generate_procurement_report(self):
        """ 核心計算邏輯：增加資料清洗與補零邏輯，防止 NaN 錯誤 """
        if not hasattr(self, 'tree_procure'): return
        for i in self.tree_procure.get_children(): self.tree_procure.delete(i)
        
        try:
            # 1. 讀取資料
            if not os.path.exists(FILE_NAME): return
            with pd.ExcelFile(FILE_NAME) as xls:
                df_sales = pd.read_excel(xls, sheet_name=SHEET_SALES)
                df_prods = pd.read_excel(xls, sheet_name=SHEET_PRODUCTS)
            
            if df_prods.empty: return

            # --- [關鍵修正：資料清洗] ---
            # 將數值欄位強制轉換為數字，如果原本是空白或文字，會變成 NaN，接著用 .fillna(0) 全部補 0
            num_cols = ['目前庫存', '安全庫存', '預設成本']
            for col in num_cols:
                if col in df_prods.columns:
                    df_prods[col] = pd.to_numeric(df_prods[col], errors='coerce').fillna(0)
                else:
                    df_prods[col] = 0.0 # 如果根本沒這一欄，直接補 0
            
            df_sales['數量'] = pd.to_numeric(df_sales['數量'], errors='coerce').fillna(0)
            # ---------------------------

            now = pd.Timestamp.now()
            # 處理初始上架時間 (如果空白就用現在時間)
            start_col = "初始上架時間"
            if start_col not in df_prods.columns:
                df_prods[start_col] = df_prods.get("最後更新時間", now)
            
            df_prods['start_dt'] = pd.to_datetime(df_prods[start_col], errors='coerce').fillna(now)
            
            # 獲取各商品總銷量
            qty_sum = df_sales.groupby('商品名稱')['數量'].sum()
            
            # 讀取介面參數 (加 try-except 防止介面輸入非數字)
            try:
                v_threshold = float(self.var_filter_velocity.get()) # 速度門檻
                s_multiplier = float(self.var_safety_multiplier.get()) # 安全係數
                cover_days = float(self.var_days_to_cover.get()) # 備貨天數

            except:
                v_threshold = 0.1
                s_multiplier = 1.0
                cover_days = 30.0 # 預設備貨天數

            for _, row in df_prods.iterrows():
                p_name = str(row['商品名稱'])
                curr_stock = float(row['目前庫存'])
                base_safety = float(row['安全庫存'])
                
                # A. 計算目前的日均銷量 (Velocity)
                total_sold = float(qty_sum.get(p_name, 0))
                days_since_start = (now - row['start_dt']).days
                velocity = total_sold / max(days_since_start, 1)

                # B. 計算目標庫存量
                # 目標 = (每天賣幾件 * 準備賣幾天) + 加權後的安全存量
                target_inventory = (velocity * cover_days) + (base_safety * s_multiplier)
                
                # C. 計算建議採購量 (無條件進位，因為商品沒有 0.5 件)
                import math
                raw_suggest = target_inventory - curr_stock
                suggest_qty = math.ceil(max(raw_suggest, 0))

                # D. 判定顯示狀態
                status = ""
                tag = ""
                
                # 只有符合以下條件才出現在清單：
                # 1. 庫存告急 (低於安全存量)
                # 2. 帳面超賣 (負數)
                # 3. 且銷售速度達到您的門檻 (或是超賣必補)
                
                if curr_stock < 0:
                    status = "⚠️ 帳面超賣"; tag = 'urgent'
                elif curr_stock <= (base_safety * s_multiplier) and velocity >= v_threshold:
                    status = "🔴 需補貨"; tag = 'urgent'
                elif curr_stock <= (base_safety * s_multiplier) and (base_safety > 0):
                    status = "🟡 庫存偏低"; tag = 'warning'
                else:
                    continue # 庫存還很足夠，不用採購

                self.tree_procure.insert("", "end", values=(
                    p_name, 
                    int(curr_stock), 
                    round(base_safety * s_multiplier, 1), 
                    f"{round(velocity, 2)}件/日", 
                    status, 
                    int(suggest_qty) # 這裡現在是根據「備貨天數」算出的科學數值
                ), tags=(tag,))
                
        except Exception as e:
            import traceback
            messagebox.showerror("評估失敗", f"錯誤原因: {str(e)}\n\n詳細資訊已印在終端機")
            traceback.print_exc()

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
            salt = globals().get('SECRET_SALT', "redmaple") # 確保 Salt 一致
            raw_string = user_id + salt
            # --- 這裡改成 sha256 ---
            expected_code = hashlib.sha256(raw_string.encode()).hexdigest()[:8].upper()
        except:
            raw_string = user_id + "redmaple"
            expected_code = hashlib.sha256(raw_string.encode()).hexdigest()[:8].upper()

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
                
            expected_code = hashlib.sha256(raw_string.encode()).hexdigest()[:8].upper()
            
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

        ttk.Button(right_frame, text="(x) 移除", command=self.remove_from_cart).pack(anchor="e", pady=2)

        fee_frame = ttk.LabelFrame(right_frame, text="費用與折扣", padding=10)
        fee_frame.pack(fill="x", pady=5)
        
        # 第一排：平台費率
        f1 = ttk.Frame(fee_frame)
        f1.pack(fill="x")
        ttk.Label(f1, text="平台費率:").pack(side="left")
        self.combo_fee_rate = ttk.Combobox(f1, textvariable=self.var_fee_rate_str, state="readonly", width=28)
        self.combo_fee_rate.pack(side="left", padx=5)
        self.combo_fee_rate.bind('<<ComboboxSelected>>', self.on_fee_option_selected)

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
            print(f"載入追蹤清單失敗: {e}")

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
            self._universal_save({ SHEET_TRACKING: df })
            messagebox.showinfo("成功", "商品已刪除"); self.load_tracking_data()
        except Exception as e: messagebox.showerror("錯誤", f"刪除失敗: {e}")


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

    def _save_all_sheets(self, df_target, target_sheet_name):
        """ 通用輔助函式：儲存單一變動分頁並保護其他所有分頁 """
        try:
            # 先讀取所有現有的 Sheet 內容
            with pd.ExcelFile(FILE_NAME) as xls:
                all_sheets = {sn: pd.read_excel(xls, sheet_name=sn) for sn in xls.sheet_names}
            
            # 更新目標 Sheet
            all_sheets[target_sheet_name] = df_target
            
            # 全部寫回
            with pd.ExcelWriter(FILE_NAME, engine='openpyxl') as writer:
                for sn, df in all_sheets.items():
                    df.to_excel(writer, sheet_name=sn, index=False)
        except Exception as e:
            messagebox.showerror("存檔錯誤", str(e))


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
                    print(f"售後扣庫存：{prod_name} 由 {old_stock} -> {old_stock-1}")

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
        sel = self.tree_sales_edit.selection()
        if not sel: return
        
        item = self.tree_sales_edit.item(sel[0])
        idx = int(item['text']) 

        try:
            df = pd.read_excel(FILE_NAME, sheet_name=SHEET_SALES)
            row = df.iloc[idx]
            
            # 更新訂單詳情
            self.var_view_oid.set(str(row.get('訂單編號', '')).replace("'", ""))
            self.var_view_buyer.set(str(row.get('買家名稱', '')))
            self.var_view_ship.set(str(row.get('寄送方式', '')))
            self.var_view_item.set(str(row.get('商品名稱', '')))
            self.var_view_tax.set(f"${row.get('稅額', 0)}")
            
            # --- [即時顯示售後狀態] ---
            # 抓取「扣費項目」欄位
            current_after_note = str(row.get('扣費項目', '')).strip()
            if current_after_note == "" or current_after_note == "nan":
                self.var_view_after_status.set("目前無售後紀錄")
            else:
                self.var_view_after_status.set(current_after_note)
            
        except Exception as e:
            print(f"讀取詳情失敗: {e}")


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



        # 商家名稱輸入
        ttk.Label(font_frame, text="商家名稱:").pack(side="left", padx=5)
        ent_shop = ttk.Entry(font_frame, textvariable=self.var_shop_name, width=20)
        ent_shop.pack(side="left", padx=5)
        
        # --- 新增：儲存按鈕 ---
        btn_save_cfg = ttk.Button(font_frame, text="💾 儲存設定", command=self.save_system_settings)
        btn_save_cfg.pack(side="left", padx=5)

        ttk.Label(font_frame, text="(調整後需重啟或切換分頁生效)", foreground="gray").pack(side="right", padx=10)
        spin = ttk.Spinbox(font_frame, from_=10, to=20, textvariable=self.var_font_size, width=5, command=self.change_font_size)
        spin.pack(side="right", padx=5)
        ttk.Label(font_frame, text="字型大小 (10-20):").pack(side="right", padx=5)



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

        ttk.Label(ctrl_frame, text="固定金額 ($):").pack(anchor="w")
        self.ent_fee_fixed = ttk.Entry(ctrl_frame, width=15)
        self.ent_fee_fixed.insert(0, "0") # 預設為 0
        self.ent_fee_fixed.pack(pady=5)

        ttk.Button(ctrl_frame, text="➕ 新增/更新", command=self.action_add_custom_fee).pack(fill="x", pady=5)
        ttk.Button(ctrl_frame, text="🗑️ 刪除選取", command=self.action_delete_custom_fee).pack(fill="x", pady=5)
        ttk.Label(ctrl_frame, text="*修改後銷售頁面\n選單會同步更新", foreground="gray", font=("", 9)).pack(pady=10)

        field_cfg_frame = ttk.LabelFrame(main_frame, text="👁️ 商品資料欄位顯示設定 (勾選欲使用的功能)", padding=15)
        field_cfg_frame.pack(fill="x", pady=10)

        # 建立兩排勾選框
        row_f = ttk.Frame(field_cfg_frame)
        row_f.pack(fill="x")

        for i, (label, var) in enumerate(self.show_fields.items()):
            # 點擊勾選框時，即時觸發介面刷新
            chk = ttk.Checkbutton(row_f, text=label, variable=var, 
                                command=self.refresh_product_ui_layout)
            chk.pack(side="left", padx=15, pady=5)

        ttk.Label(field_cfg_frame, text="* 隱藏欄位不會刪除資料，僅是在輸入與編輯介面中暫時收起。", 
                foreground="gray", font=("", 9)).pack(anchor="w")

        

        # 載入初始費率資料
        self.refresh_fee_tree()

    def refresh_fee_tree(self):
            if hasattr(self, 'fee_tree'):
                for i in self.fee_tree.get_children(): self.fee_tree.delete(i)
            
            self.fee_lookup = {} # 清空舊資料
            
            try:
                df = pd.read_excel(FILE_NAME, sheet_name=SHEET_CONFIG)
                fee_options = ["自訂手動輸入"]
                
                for _, row in df.iterrows():
                    name = str(row['設定名稱']).strip()
                    perc = float(row['費率百分比'])
                    fixed = float(row.get('固定金額', 0))
                    
                    # --- 核心改動：存入對照表 ---
                    display_str = f"{name} ({perc}% + ${fixed})" if fixed > 0 else f"{name} ({perc}%)"
                    self.fee_lookup[display_str] = (perc, fixed) # 用「顯示字串」當 Key
                    
                    fee_options.append(display_str)
                    if hasattr(self, 'fee_tree'):
                        self.fee_tree.insert("", "end", values=(name, perc, fixed))
                
                if hasattr(self, 'combo_fee_rate'):
                    self.combo_fee_rate['values'] = fee_options
                    # 預設選取第一個有效費率
                    if len(fee_options) > 1:
                        self.combo_fee_rate.set(fee_options[1])
                    else:
                        self.combo_fee_rate.set("自訂手動輸入")
            except:
                pass

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
                    df = pd.read_excel(FILE_NAME, sheet_name=SHEET_CONFIG)
                    
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
            save_success = self._universal_save({SHEET_CONFIG: df})
            
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
        
        lbl_version = ttk.Label(header_frame, text="Version 4.3 (採購決策優化版)", font=("Consolas", 11), foreground="gray")
        lbl_version.pack(anchor="center")

        # --- 中間：功能簡介與開發者資訊 ---
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill="both", expand=True)

        # 左側：核心功能
        left_box = ttk.LabelFrame(content_frame, text="🚀 系統核心價值", padding=15)
        left_box.pack(side="left", fill="both", expand=True, padx=10)
        
        features = [
            "● 全自動加權平均成本 (WAC) 計算",
            "● 支援內含營業稅 (5%) 自動回推",
            "● 進貨追蹤與結案緩衝區雙重機制",
            "● 智慧採購評估系統 (銷售速率/備貨天數)",
            "● 歷史訂單自動日期排序與資料保護",
            "● 支援雲端 Google Drive 自動替換備份"
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

        # --- 底部：更新日誌 ---
        log_frame = ttk.LabelFrame(main_frame, text="📝 更新日誌", padding=10)
        log_frame.pack(fill="x", pady=20)
        
        log_text = tk.Text(log_frame, height=5, font=("微軟正黑體", 10), bg="#F8F9FA", relief="flat")
        log_text.pack(fill="x")
        
        logs = (
            "[2026-02-08] V4.3: 引入採購需求分析模組、優化銷售速率計算邏輯。\n"
            "[2026-02-05] V4.2: 進貨與銷售端同步支援『內含營業稅』回推計算。\n"
            "[2026-02-02] V4.1: 進貨管理全面單據化，支援批次入庫與加權成本公式。\n"
            "[2026-01-31] V4.0: 移除原生 Excel 依賴，轉向資料庫邏輯架構 (V4 Hybrid)。"
        )
        log_text.insert("1.0", logs)
        log_text.config(state="disabled") 

        # 版權宣告
        lbl_copyright = ttk.Label(main_frame, text="© 2026 redmaple. All Rights Reserved.", foreground="#CED4DA")
        lbl_copyright.pack(side="bottom", pady=5)
    # ---------------- 邏輯功能區 ----------------

    def action_cancel_purchase(self):
        """ 標記遺失或取消：從『進貨紀錄』與『進貨追蹤』中同時刪除該筆資料 """
        sel = self.tree_pur_track.selection()
        if not sel: return
        
        item = self.tree_pur_track.item(sel[0])
        idx_in_track = int(item['text'])
        pur_id = item['values'][0]
        p_name = item['values'][2]

        if not messagebox.askyesno("取消確認", f"確定要【完全刪除】單號 {pur_id} 的這筆進貨嗎？\n(這將同時移除進貨紀錄與追蹤清單)"):
            return

        try:
            with pd.ExcelFile(FILE_NAME) as xls:
                df_tracking = pd.read_excel(xls, sheet_name=SHEET_PUR_TRACKING)
                df_history = pd.read_excel(xls, sheet_name=SHEET_PURCHASES)
                # 其餘分頁
                others = {sn: pd.read_excel(xls, sheet_name=sn) for sn in xls.sheet_names if sn not in [SHEET_PUR_TRACKING, SHEET_PURCHASES]}

            # 1. 從追蹤分頁刪除 (根據 index)
            df_tracking.drop(idx_in_track, inplace=True)

            # 2. 從進貨紀錄分頁刪除 (根據單號與品名)
            clean_id = str(pur_id).replace("'", "")
            mask = (df_history['進貨單號'].astype(str).str.contains(clean_id)) & (df_history['商品名稱'] == p_name)
            df_history = df_history[~mask]

            # 3. 寫回所有資料
            with pd.ExcelWriter(FILE_NAME, engine='openpyxl') as writer:
                df_tracking.to_excel(writer, sheet_name=SHEET_PUR_TRACKING, index=False)
                df_history.to_excel(writer, sheet_name=SHEET_PURCHASES, index=False)
                for sn, df in others.items(): df.to_excel(writer, sheet_name=sn, index=False)

            messagebox.showinfo("成功", f"進貨紀錄已完全移除。")
            self.load_purchase_tracking()
        except Exception as e:
            messagebox.showerror("錯誤", f"取消失敗: {e}")

    def action_confirm_inbound(self):
        """ [修正版] 確認收貨：解決日期格式 float64 報錯，並精準計算落地成本 """
        sel = self.tree_pur_track.selection()
        if not sel: 
            messagebox.showwarning("提示", "請先選擇要入庫的項目")
            return
        
        # 1. 從介面取得 8 個欄位的數值 (對齊索引)
        item = self.tree_pur_track.item(sel[0])
        idx_in_track_df = int(item['text']) 
        vals = item['values'] 

        pur_id = str(vals[0]).replace("'", "") # 單號
        p_name = vals[2]                       # 商品名稱
        new_qty = int(vals[3])                 # 數量
        new_price = float(vals[4])             # 進貨單價
        customs_tax = float(vals[5])           # 稅額 (索引 5)
        ship_fee = float(vals[6])              # 運費 (索引 6)

        if not messagebox.askyesno("確認入庫", f"商品: {p_name}\n即將入庫 {new_qty} 件。\n(含運費 ${ship_fee}, 稅金 ${customs_tax})\n\n系統將自動更新庫存並攤平平均成本。"):
            return

        try:
            today_str = datetime.now().strftime("%Y-%m-%d")

            # 2. 讀取 Excel 內容
            with pd.ExcelFile(FILE_NAME) as xls:
                df_prods = pd.read_excel(xls, sheet_name=SHEET_PRODUCTS)
                df_tracking = pd.read_excel(xls, sheet_name=SHEET_PUR_TRACKING)
                df_history = pd.read_excel(xls, sheet_name=SHEET_PURCHASES)

            # --- [核心修正：解決 float64 格式報錯] ---
            # 強制將入庫日期轉為 object (字串) 格式，避免 Pandas 報錯
            if '入庫日期' in df_history.columns:
                df_history['入庫日期'] = df_history['入庫日期'].astype(object).fillna("")
            
            # 確保運費與稅額欄位存在且為數值
            for col in ['分攤運費', '海關稅金']:
                if col not in df_history.columns: df_history[col] = 0.0
                df_history[col] = pd.to_numeric(df_history[col], errors='coerce').fillna(0.0)

            # 3. 【計算落地成本 (Landed Cost)】
            # 本批次總投入 = (數量 * 進價) + 運費 + 稅金
            current_batch_total_cost = (new_qty * new_price) + ship_fee + customs_tax
            
            if p_name in df_prods['商品名稱'].values:
                p_idx = df_prods[df_prods['商品名稱'] == p_name].index[0]
                
                # 取得舊庫存與舊成本
                old_stock = float(df_prods.at[p_idx, '目前庫存']) if pd.notna(df_prods.at[p_idx, '目前庫存']) else 0
                old_cost = float(df_prods.at[p_idx, '預設成本']) if pd.notna(df_prods.at[p_idx, '預設成本']) else 0
                
                total_qty = old_stock + new_qty
                
                # 加權平均成本公式
                if total_qty > 0:
                    if old_stock <= 0:
                        # 原本沒貨或超賣，直接以本次總成本攤平
                        weighted_cost = current_batch_total_cost / new_qty
                    else:
                        # 公式：(舊庫存總值 + 本批總值) / 總數量
                        weighted_cost = ((old_stock * old_cost) + current_batch_total_cost) / total_qty
                    
                    # A. 更新商品庫存與「落地」成本
                    df_prods.at[p_idx, '預設成本'] = round(weighted_cost, 2)
                    df_prods.at[p_idx, '目前庫存'] = total_qty
                    df_prods.at[p_idx, '最後進貨時間'] = today_str
                    df_prods.at[p_idx, '最後更新時間'] = datetime.now().strftime("%Y-%m-%d %H:%M")

            # 4. 【同步更新進貨紀錄總帳】
            clean_id = str(pur_id).replace("'", "")
            # 建立暫時過濾欄位避免修改到原始編號
            df_history['tmp_id'] = df_history['進貨單號'].astype(str).str.replace("'", "").str.strip()
            mask = (df_history['tmp_id'] == clean_id) & (df_history['商品名稱'] == p_name)
            
            if not df_history[mask].empty:
                df_history.loc[mask, '入庫日期'] = today_str
                df_history.loc[mask, '備註'] = "已完成入庫"
                df_history.loc[mask, '分攤運費'] = ship_fee
                df_history.loc[mask, '海關稅金'] = customs_tax
            
            df_history.drop(columns=['tmp_id'], inplace=True)

            # 5. 【移除追蹤清單】
            df_tracking.drop(idx_in_track_df, inplace=True)

            # 6. 【萬用引擎存檔】
            save_success = self._universal_save({
                SHEET_PRODUCTS: df_prods,
                SHEET_PUR_TRACKING: df_tracking,
                SHEET_PURCHASES: df_history
            })

            if save_success:
                messagebox.showinfo("成功", f"【入庫完成】\n商品: {p_name}\n庫存已補至: {total_qty}\n平均成本(含運費稅金): ${round(weighted_cost, 2)}")
                self.load_purchase_tracking() 
                self.products_df = self.load_products() 
                self.update_sales_prod_list() 

        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("入庫失敗", f"發生錯誤: {str(e)}")



    def update_pur_prod_list(self):
        """ 同步商品資料管理裡的商品名稱到進貨列表 (修正版) """
        # 檢查 list_pur_prod 是否存在，避免 Attribute Error
        if hasattr(self, 'list_pur_prod') and not self.products_df.empty:
            names = self.products_df['商品名稱'].tolist()
            # 清空目前的列表
            self.list_pur_prod.delete(0, tk.END)
            # 將商品名稱逐一放入列表框
            for name in names:
                self.list_pur_prod.insert(tk.END, name)

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
        self.ent_pur_search.delete(0, tk.END) # 清空搜尋框
        self.update_pur_prod_list() # 恢復完整列表

    def remove_from_pur_cart(self):
        """ 移除選中項目 """
        sel = self.tree_pur_cart.selection()
        if not sel: return
        for item in sel:
            idx = self.tree_pur_cart.index(item)
            del self.pur_cart_data[idx]
            self.tree_pur_cart.delete(item)
        
        total_sum = sum(item['total'] for item in self.pur_cart_data)
        self.lbl_pur_total.config(text=f"本次進貨總額: ${total_sum:,.0f}")





    def submit_purchase(self):
        """ 提交進貨：更新庫存、更新成本、記錄進貨單 """
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
            # 1. 讀取所有分頁
            with pd.ExcelFile(FILE_NAME) as xls:
                df_prods = pd.read_excel(xls, sheet_name=SHEET_PRODUCTS)
                df_pur = pd.read_excel(xls, sheet_name=SHEET_PURCHASES)
                # 讀取其他分頁以防遺失
                df_sales = pd.read_excel(xls, sheet_name=SHEET_SALES)
                df_track = pd.read_excel(xls, sheet_name=SHEET_TRACKING)
                df_ret = pd.read_excel(xls, sheet_name=SHEET_RETURNS)
                df_cfg = pd.read_excel(xls, sheet_name=SHEET_CONFIG)

            # 2. 更新商品庫存與成本
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
            if name in df_prods['商品名稱'].values:
                idx = df_prods[df_prods['商品名稱'] == name].index[0]
                df_prods.at[idx, '目前庫存'] += qty
                df_prods.at[idx, '預設成本'] = cost # 進貨價格自動更新成本
                df_prods.at[idx, '最後更新時間'] = now_str
                df_prods.at[idx, '最後進貨時間'] = now_str
            else:
                messagebox.showerror("錯誤", f"找不到商品「{name}」，請先到商品管理新增。")
                return

            # 3. 建立進貨紀錄
            new_pur = pd.DataFrame([{
                "進貨單號": f"'{pur_id}", # 強制字串
                "進貨日期": date_str,
                "供應商": supplier,
                "物流追蹤編號": logistics,
                "商品名稱": name,
                "數量": qty,
                "進貨單價": cost,
                "進貨總額": qty * cost,
                "備註": ""
            }])
            df_pur = pd.concat([df_pur, new_pur], ignore_index=True)

            # 4. 一次性寫回
            with pd.ExcelWriter(FILE_NAME, engine='openpyxl') as writer:
                df_prods.to_excel(writer, sheet_name=SHEET_PRODUCTS, index=False)
                df_pur.to_excel(writer, sheet_name=SHEET_PURCHASES, index=False)
                df_sales.to_excel(writer, sheet_name=SHEET_SALES, index=False)
                df_track.to_excel(writer, sheet_name=SHEET_TRACKING, index=False)
                df_ret.to_excel(writer, sheet_name=SHEET_RETURNS, index=False)
                df_cfg.to_excel(writer, sheet_name=SHEET_CONFIG, index=False)

            messagebox.showinfo("成功", f"進貨單 {pur_id} 已入庫！\n庫存已自動增加 {qty}。")
            
            # 清除輸入並刷新
            self.var_pur_qty.set(1); self.var_pur_cost.set(0.0); self.var_pur_logistics.set("")
            self.load_purchase_data()
            self.products_df = df_prods # 同步介面數據
            self.update_sales_prod_list() # 更新銷售頁面庫存顯示
            
        except Exception as e:
            messagebox.showerror("錯誤", f"進貨作業失敗: {e}")

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


    def action_update_pur_logistics(self):
        """ 彈出視窗：修正讀取索引 """
        sel = self.tree_pur_track.selection()
        if not sel: return
        
        item = self.tree_pur_track.item(sel[0])
        idx = int(item['text'])
        vals = item['values'] # 取得 8 個欄位的陣列
        
        pur_id = str(vals[0])
        p_name = vals[2]

        win = tk.Toplevel(self.root)
        win.title("更新物流與附加成本")
        win.geometry("350x400")
        
        # 抓取目前的舊資料
        old_tax = vals[5]
        old_ship = vals[6]
        old_logi = vals[7]

        ttk.Label(win, text=f"單號: {pur_id}", foreground="gray").pack(pady=5)
        ttk.Label(win, text=f"商品: {p_name}", font=("", 10, "bold")).pack(pady=5)

        ttk.Label(win, text="1. 物流單號:").pack(anchor="w", padx=30)
        var_logi = tk.StringVar(value=old_logi)
        ttk.Entry(win, textvariable=var_logi).pack(fill="x", padx=30)

        ttk.Label(win, text="2. 分攤運費 ($):").pack(anchor="w", padx=30, pady=(10,0))
        var_ship = tk.DoubleVar(value=old_ship)
        ttk.Entry(win, textvariable=var_ship).pack(fill="x", padx=30)

        ttk.Label(win, text="3. 海關稅金/加稅 ($):").pack(anchor="w", padx=30, pady=(10,0))
        var_tax = tk.DoubleVar(value=old_tax)
        ttk.Entry(win, textvariable=var_tax).pack(fill="x", padx=30)

        # 存檔按鈕邏輯保持不變，但確保讀取的是這三個變數...
        def save_logic():
            try:
                with pd.ExcelFile(FILE_NAME) as xls:
                    df_track = pd.read_excel(xls, sheet_name=SHEET_PUR_TRACKING)
                    df_hist = pd.read_excel(xls, sheet_name=SHEET_PURCHASES)

                for df in [df_track, df_hist]:
                    if '分攤運費' not in df.columns: df['分攤運費'] = 0
                    if '海關稅金' not in df.columns: df['海關稅金'] = 0
                    
                    m = (df['進貨單號'].astype(str).str.contains(pur_id)) & (df['商品名稱'] == p_name)
                    df.loc[m, '物流追蹤'] = var_logi.get()
                    df.loc[m, '分攤運費'] = var_ship.get()
                    df.loc[m, '海關稅金'] = var_tax.get()

                if self._universal_save({SHEET_PUR_TRACKING: df_track, SHEET_PURCHASES: df_hist}):
                    messagebox.showinfo("成功", "資料已更新")
                    self.load_purchase_tracking()
                    win.destroy()
            except Exception as e: messagebox.showerror("錯誤", str(e))

        ttk.Button(win, text="💾 儲存修改", command=save_logic).pack(pady=25)

        def save_and_close():
            try:
                with pd.ExcelFile(FILE_NAME) as xls:
                    df_track = pd.read_excel(xls, sheet_name=SHEET_PUR_TRACKING)
                    df_hist = pd.read_excel(xls, sheet_name=SHEET_PURCHASES)

                # 更新資料
                for df in [df_track, df_hist]:
                    # 這裡要確保 Excel 有這兩個欄位
                    if '分攤運費' not in df.columns: df['分攤運費'] = 0
                    if '海關稅金' not in df.columns: df['海關稅金'] = 0
                    
                    # 匹配單號與商品
                    m = (df['進貨單號'].astype(str).str.contains(pur_id)) & (df['商品名稱'] == p_name)
                    df.loc[m, '物流追蹤'] = var_logi.get()
                    df.loc[m, '分攤運費'] = var_ship.get()
                    df.loc[m, '海關稅金'] = var_tax.get()

                self._universal_save({SHEET_PUR_TRACKING: df_track, SHEET_PURCHASES: df_hist})
                messagebox.showinfo("成功", "附加成本已更新")
                self.load_purchase_tracking()
                win.destroy()
            except Exception as e:
                messagebox.showerror("錯誤", str(e))

        ttk.Button(win, text="💾 儲存並更新", command=save_and_close).pack(pady=20)

    
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
            self._universal_save({ SHEET_TRACKING: df })
            
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
        """ 退貨單一商品 (修正存檔格式) """
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

            info = self._get_full_order_info(df_track, order_id)
            row_to_move = df_track.loc[[idx]].copy()
            for col, val in info.items(): row_to_move[col] = val
            row_to_move['備註'] = reason

            # 補位邏輯
            is_header = pd.notna(df_track.at[idx, '日期']) and str(df_track.at[idx, '日期']) != ""
            if is_header:
                others = df_track[(df_track['訂單編號'] == order_id) & (df_track.index != idx)].index.tolist()
                if others:
                    new_h = others[0]
                    for col in info.keys(): df_track.at[new_h, col] = df_track.at[idx, col]

            df_track.drop(idx, inplace=True)
            try: df_returns = pd.read_excel(FILE_NAME, sheet_name=SHEET_RETURNS)
            except: df_returns = pd.DataFrame()
            df_returns = pd.concat([df_returns, row_to_move], ignore_index=True)

            # ---【關鍵修正：使用大括號字典傳參】---
            success = self._universal_save({
                SHEET_TRACKING: df_track, 
                SHEET_RETURNS: df_returns
            })
            
            if success:
                messagebox.showinfo("成功", f"商品「{prod_name}」已移至退貨紀錄。")
                self.load_tracking_data(); self.load_returns_data()
        except Exception as e: messagebox.showerror("錯誤", str(e))


    
    def action_track_complete_order(self):
        """ 完成訂單/整筆結案 (修正存檔格式) """
        sel = self.tree_track.selection()
        if not sel: return
        item = self.tree_track.item(sel[0]); order_id = str(item['values'][0]).replace("'", "")

        if not messagebox.askyesno("結案確認", f"確定訂單 [{order_id}] 已完成？"): return

        try:
            df_track = pd.read_excel(FILE_NAME, sheet_name=SHEET_TRACKING)
            df_track['訂單編號'] = df_track['訂單編號'].astype(str).str.replace(r'^\'', '', regex=True).str.replace(r'\.0$', '', regex=True)
            
            try: df_sales = pd.read_excel(FILE_NAME, sheet_name=SHEET_SALES)
            except: df_sales = pd.DataFrame()

            mask = df_track['訂單編號'] == order_id
            rows_to_finish = df_track[mask].copy()
            info = self._get_full_order_info(df_track, order_id)
            for col, val in info.items(): rows_to_finish[col] = val

            df_sales_combined = pd.concat([df_sales, rows_to_finish], ignore_index=True)
            df_track_new = df_track[~mask]

            # ---【關鍵修正：使用大括號字典傳參】---
            success = self._universal_save({
                SHEET_TRACKING: df_track_new, 
                SHEET_SALES: df_sales_combined
            })
            
            if success:
                messagebox.showinfo("成功", f"訂單 {order_id} 已結案！")
                self.load_tracking_data(); self.calculate_analysis_data()
        except Exception as e: messagebox.showerror("錯誤", str(e))

    def _universal_save(self, updates_dict):
        """ 強化版萬用存檔引擎：防止分頁消失，自動保護所有分頁 """
        try:
            all_data = {}
            # 1. 先讀取目前 Excel 裡「所有的」分頁內容
            if os.path.exists(FILE_NAME):
                with pd.ExcelFile(FILE_NAME) as xls:
                    # 遍歷 Excel 檔案裡實際存在的每一個分頁名稱
                    for sn in xls.sheet_names:
                        all_data[sn] = pd.read_excel(xls, sheet_name=sn)
            
            # 2. 將本次有變動的分頁「覆蓋」進字典中
            for sheet_name, df in updates_dict.items():
                all_data[sheet_name] = df

            # 3. 處理數據格式（防止科學記號、處理日期）
            for sn, df in all_data.items():
                if df is None or df.empty: continue
                
                # 保護 ID 欄位
                for id_col in ['訂單編號', '進貨單號']:
                    if id_col in df.columns:
                        df[id_col] = df[id_col].apply(lambda x: f"'{str(x).replace('\'','')}" if pd.notna(x) and str(x).strip() != "" else x)

            # 4. 寫回 Excel (使用 replace 模式確保分頁不丟失)
            with pd.ExcelWriter(FILE_NAME, engine='openpyxl') as writer:
                # 按照我們定義的標準順序排列分頁
                standard_order = [SHEET_PRODUCTS, SHEET_SALES, SHEET_TRACKING, SHEET_PURCHASES, SHEET_PUR_TRACKING, SHEET_RETURNS, SHEET_CONFIG]
                
                # 先寫入標準分頁
                for sn in standard_order:
                    if sn in all_data:
                        all_data[sn].to_excel(writer, sheet_name=sn, index=False)
                
                # 如果還有其他不在標準列表裡的分頁，也補寫回去
                for sn, df in all_data.items():
                    if sn not in standard_order:
                        df.to_excel(writer, sheet_name=sn, index=False)
            
            return True
        except PermissionError:
            messagebox.showerror("存檔失敗", "Excel 檔案正被開啟中，請先關閉 Excel 後再按存檔！")
            return False
        except Exception as e:
            messagebox.showerror("嚴重錯誤", f"存檔引擎故障: {str(e)}")
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
        """ 銷售搜尋框：顯示 [編號][分類] 名稱 (庫存)，並支援編號搜尋 """
        search_term = self.var_search.get().lower()
        self.listbox_sales.delete(0, tk.END)
        
        if not self.products_df.empty:
            for index, row in self.products_df.iterrows():
                p_name = str(row['商品名稱'])
                
                # --- 容錯處理：處理空編號 ---
                raw_sku = row.get('商品編號', '')
                # 如果是 pandas 的 NaN 或 None，轉為空字串
                sku = str(raw_sku) if pd.notna(raw_sku) else ""
                sku = sku if sku.lower() != "nan" else ""
                
                p_tag = str(row['分類Tag']) if pd.notna(row['分類Tag']) else "無"
                
                try: p_stock = int(row['目前庫存'])
                except: p_stock = 0
                
                # 顯示字串：不含編號
                display_str = f"[{p_tag}] {p_name} (庫存: {p_stock})"
                
                # 搜尋邏輯：如果沒編號，sku.lower() 就會是空字串，不會匹配到關鍵字，這很安全
                if (search_term in p_name.lower() or 
                    search_term in p_tag.lower() or 
                    search_term in sku.lower()):
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
        selected_text = self.combo_fee_rate.get()
        match = re.search(r"\((\d+\.?\d*)%\)", selected_text)
        if match: self.update_totals()
        elif "自訂" in selected_text: self.combo_fee_rate.set("") 
        self.update_totals()


    def update_totals_event(self, event): self.update_totals()
    
    
    def update_totals(self):
        try:
            # 1. 基礎商品總額與成本
            t_sales = sum(i['total_sales'] for i in self.cart_data)
            t_cost = sum(i['total_cost'] for i in self.cart_data)
            
            # --- [保持原本的費率對照表邏輯，不變動] ---
            selection = self.var_fee_rate_str.get()
            rate = 0.0
            fixed_fee = 0.0
            if selection in self.fee_lookup:
                rate, fixed_fee = self.fee_lookup[selection]
            else:
                try:
                    rate = float(selection.replace("%", ""))
                except:
                    rate = 0.0

            # ---------------------------------------

            # 2. 獲取新增的 運費 與 扣費(折扣)
            try: 
                ship_fee = float(self.var_ship_fee.get())  # 賣家負擔的運費
            except: 
                ship_fee = 0.0

            try: 
                extra_deduct = float(self.var_extra_fee.get()) # 折扣或額外扣費
            except: 
                extra_deduct = 0.0

            payer = self.var_ship_payer.get()
            
            # 3. 計算各項支出
             # 1. 平台手續費 (只算商品的抽成)
            platform_fee = (t_sales * (rate/100)) + fixed_fee
            
            # 2. 利潤計算 (不論誰付，只要是「賣家付」，淨利就要扣掉這筆成本)
            # 淨利 = 商品總價 - 成本 - 平台費 - 折扣 - (如果是賣家付則扣除運費)
            profit = t_sales - t_cost - platform_fee - extra_deduct
            if payer == "賣家付":
                profit -= ship_fee
            
            # 3. 預估入帳 (你從平台或買家手中拿到的錢)
            # 如果買家付運費，且該運費是「代收」性質（如賣貨便、賣家宅配）：
            # 你會拿到：商品錢 + 運費 - 平台費 - 折扣
            if payer == "買家付":
                income = t_sales + ship_fee - platform_fee - extra_deduct
            else:
                income = t_sales - platform_fee - extra_deduct

            # --- 更新 UI ---
            self.lbl_gross.config(text=f"商品小計: ${t_sales:,.0f}")
            payer_color = "red" if payer == "賣家付" else "black"
            self.lbl_fee.config(text=f"手續費: -${platform_fee:,.0f} | 運費({payer}): ${ship_fee:,.0f} | 折扣: -${extra_deduct:,.0f}")
            self.lbl_income.config(text=f"實收/撥款總額: ${income:,.1f}")
            self.lbl_profit.config(text=f"本單純利: ${profit:,.1f}", foreground="green" if profit > 0 else "red")

            return t_sales, platform_fee, 0
        except Exception as e:
            print(f"計算出錯: {e}")
            return 0, 0, 0
        
    
        
    def submit_order(self):
        """ 修正版：送出訂單至追蹤區，確保不覆蓋舊有資料 """
        if not self.cart_data: return
        
        def clean_text(text):
            if not text: return ""
            return text.replace("\n", "").replace("\r", "").strip()

        if self.var_enable_cust.get():
            cust_name = self.var_cust_name.get().strip()
            if not cust_name or cust_name == "":
                messagebox.showerror("欄位缺失", "您已勾選『填寫來源與顧客』，請務必輸入『買家名稱』！")
                # 將焦點移回輸入框，方便使用者補填
                self.entry_cust_name.focus()
                return
            
            # 其餘資訊抓取
            cust_loc = self.var_cust_loc.get().strip()
            ship_method = self.var_ship_method.get()
            platform_name = self.var_platform.get()
        else:
            cust_name = "未提供" ; cust_loc = "未提供" ; ship_method = "未提供" ; platform_name = "零售/現場"
            
        date_str = self.var_date.get().strip()
        now = datetime.now()
        order_id = now.strftime("%Y%m%d%H%M%S") 

        t_sales, t_fee, t_tax = self.update_totals() 
        fee_tag = self.var_fee_tag.get()
        try: extra_val = float(self.var_extra_fee.get())
        except: extra_val = 0
        if extra_val > 0 and not fee_tag: fee_tag = "其他"
        elif extra_val == 0: fee_tag = ""

        try:
            rows = []
            out_of_stock_warnings = [] 
            
            # 1. 讀取目前的商品資料 (用於更新庫存)
            df_prods_current = pd.read_excel(FILE_NAME, sheet_name=SHEET_PRODUCTS)

            # 2. 準備本次新訂單的資料列
            for i, item in enumerate(self.cart_data):
                if i == 0:
                    row_date, row_platform, row_buyer, row_ship, row_loc = date_str, platform_name, cust_name, ship_method, cust_loc
                else:
                    row_date = row_platform = row_buyer = row_ship = row_loc = ""

                ratio = item['total_sales'] / t_sales if t_sales > 0 else 0
                alloc_fee = t_fee * ratio
                alloc_tax = t_tax * ratio 
                
                net = item['total_sales'] - item['total_cost'] - alloc_fee - alloc_tax
                margin_pct = (net / item['total_sales']) * 100 if item['total_sales'] > 0 else 0.0

                rows.append({
                    "訂單編號": order_id,
                    "商品編號": item.get('sku', ''), # 這裡把 sku 存進 Excel
                    "日期": row_date, "買家名稱": row_buyer, "交易平台": row_platform,  
                    "寄送方式": row_ship, "取貨地點": row_loc,
                    "商品名稱": item['name'], "數量": item['qty'], 
                    "單價(售)": item['unit_price'], "單價(進)": item['unit_cost'],
                    "總銷售額": item['total_sales'], "總成本": item['total_cost'], 
                    "分攤手續費": round(alloc_fee, 2), "扣費項目": fee_tag, 
                    "總淨利": round(net, 2), "毛利率": round(margin_pct, 1), "稅額": round(alloc_tax, 2)
                })

                # 庫存扣除邏輯
                prod_name = item['name']
                sold_qty = item['qty']
                idxs = df_prods_current[df_prods_current['商品名稱'] == prod_name].index
                if not idxs.empty:
                    target_idx = idxs[0]
                    curr_stock = df_prods_current.at[target_idx, '目前庫存']
                    df_prods_current.at[target_idx, '目前庫存'] = curr_stock - sold_qty
                    if (curr_stock - sold_qty) <= 0:
                        out_of_stock_warnings.append(f"● {prod_name}")

            # 3. 【核心修正點】：讀取「訂單追蹤」中原本就有的資料，並與新訂單合併
            try:
                df_track_existing = pd.read_excel(FILE_NAME, sheet_name=SHEET_TRACKING)
            except:
                df_track_existing = pd.DataFrame()

            df_sales_new_batch = pd.DataFrame(rows)
            # 強制補上單引號保護編號
            df_sales_new_batch['訂單編號'] = df_sales_new_batch['訂單編號'].apply(lambda x: f"'{x}")

            # 合併新舊追蹤資料
            df_track_combined = pd.concat([df_track_existing, df_sales_new_batch], ignore_index=True)

            # 確保欄位順序正確
            excel_columns_order = ["訂單編號", "日期", "買家名稱", "交易平台", "寄送方式", "取貨地點",
                                  "商品名稱", "數量", "單價(售)", "單價(進)", "總銷售額", "總成本", 
                                  "分攤手續費", "扣費項目", "總淨利", "毛利率", "稅額"]
            df_track_combined = df_track_combined[excel_columns_order]

            # 4. 調用全能存檔引擎：一次更新商品與追蹤表，保護其他分頁
            save_success = self._universal_save({
                SHEET_PRODUCTS: df_prods_current, 
                SHEET_TRACKING: df_track_combined
            })

            if save_success:
                self.products_df = df_prods_current
                self.update_sales_prod_list()
                self.update_mgmt_prod_list()
                self.load_tracking_data() 
                messagebox.showinfo("成功", f"訂單 {order_id} 已成功加入追蹤區！")

                # 清空介面
                self.cart_data = []
                for i in self.tree.get_children(): self.tree.delete(i)
                self.update_totals()
                self.var_cust_name.set(""); self.var_cust_loc.set(""); self.var_sel_stock_info.set("--")

        except Exception as e: 
            messagebox.showerror("錯誤", f"發生未預期錯誤: {str(e)}")

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
        selection = self.listbox_mgmt.curselection()
        if selection:
            display_str = self.listbox_mgmt.get(selection[0])
            try:
                temp = display_str.rsplit(" (庫存:", 1)[0]
                selected_name = temp.split("]", 1)[1].strip() if "]" in temp else temp
            except:
                return

            record = self.products_df[self.products_df['商品名稱'] == selected_name]
            if not record.empty:
                row = record.iloc[0]
                
                # --- 核心修正：定義一個清理函數來處理 NaN ---
                def clean_val(val, default=""):
                    if pd.isna(val): return default
                    return val

                # 確保填入 UI 的資料不會出現 "NaN" 字樣
                self.var_upd_sku.set(clean_val(row.get('商品編號', '')))
                self.var_upd_name.set(clean_val(row['商品名稱']))
                self.var_upd_tag.set(clean_val(row.get('分類Tag', '')))
                self.var_upd_url.set(clean_val(row.get('商品連結', '')))
                self.var_upd_remarks.set(clean_val(row.get('商品備註', '')))
                
                # 數值欄位若為 NaN 則設為 0
                self.var_upd_safety.set(int(clean_val(row.get('安全庫存', 0), 0)))
                self.var_upd_stock.set(int(clean_val(row['目前庫存'], 0)))
                self.var_upd_cost.set(float(clean_val(row['預設成本'], 0.0)))
                self.var_upd_time.set(clean_val(row['最後更新時間'], "無資料"))

    def submit_new_product(self):
        """ 建立新商品：URL 與 備註改為選填 """
        name = self.var_add_name.get().strip()
        if not name:
            messagebox.showwarning("警告", "『商品名稱』為必填項目！")
            return
        
        try:
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
            # 讀取 URL 與 備註，如果為空則填入 "無"
            url = self.var_add_url.get().strip()
            remarks = self.var_add_remarks.get().strip()

            new_row = {
                "商品編號": self.var_add_sku.get().strip().upper(), # 自動轉大寫
                "分類Tag": self.var_add_tag.get().strip() if self.var_add_tag.get() else "未分類",
                "商品名稱": name,
                "預設成本": 0.0,
                "目前庫存": 0,
                "最後更新時間": now_str,
                "初始上架時間": now_str,
                "最後進貨時間": "",
                "安全庫存": self.var_add_safety.get(),
                "商品連結": url if url else "無",     # 選填
                "商品備註": remarks if remarks else "無" # 選填
            }
            
            df_new = pd.concat([self.products_df, pd.DataFrame([new_row])], ignore_index=True)
            
            # ---【核心修正：使用字典呼叫萬用引擎】---
            if self._universal_save({SHEET_PRODUCTS: df_new}):
                self.products_df = df_new
                self.update_mgmt_prod_list()
                self.update_pur_prod_list()
                messagebox.showinfo("成功", f"商品「{name}」已建檔！")
                # 清空輸入
                self.var_add_name.set(""); self.var_add_url.set(""); self.var_add_remarks.set("")
        except Exception as e:
            messagebox.showerror("錯誤", f"建檔失敗: {e}")

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
                    print(f"檢測到商品 {name} 補貨，更新進貨時間。")
                
                # --- [更新資料列] ---
                df_prods.loc[idx, '商品編號'] = self.var_upd_sku.get()
                df_prods.loc[idx, '分類Tag'] = self.var_upd_tag.get()
                df_prods.loc[idx, '商品名稱'] = self.var_upd_name.get()
                df_prods.loc[idx, '預設成本'] = new_cost
                df_prods.loc[idx, '目前庫存'] = new_stock
                df_prods.loc[idx, '安全庫存'] = new_safety
                df_prods.loc[idx, '商品連結'] = self.var_upd_url.get()
                df_prods.loc[idx, '商品備註'] = self.var_upd_remarks.get()
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

