#shopee-oms 3.2 測試版

import tkinter as tk
from tkinter import ttk, messagebox, font
import pandas as pd
from datetime import datetime, timedelta  # 引入 timedelta 來處理時區加減
import os
import re
import pickle
import threading 
import hashlib

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
CREDENTIALS_FILE = 'credentials.json' 
TOKEN_FILE = 'token.json'             
SCOPES = ['https://www.googleapis.com/auth/drive.file'] 

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
        """上傳檔案到指定資料夾"""
        if not self.is_authenticated: return False, "尚未登入 Google 帳號"
        if not self.folder_id: self.folder_id = self.get_or_create_folder()

        try:
            timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
            file_name = f"[系統備份] {os.path.basename(filepath).replace('.xlsx', '')}_{timestamp}.xlsx"
            
            file_metadata = {
                'name': file_name,
                'parents': [self.folder_id] 
            }
            media = MediaFileUpload(filepath, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            
            file = self.service.files().create(body=file_metadata, media_body=media, fields='id').execute()
            return True, f"備份成功！\n雲端檔名: {file_name}\n位置: {BACKUP_FOLDER_NAME}"
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
        self.root.title("蝦皮/網拍進銷存系統 (V3.6 時區修正版)")
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
        if not os.path.exists(FILE_NAME):
            try:
                with pd.ExcelWriter(FILE_NAME, engine='openpyxl') as writer:
                    cols_sales = [
                        "日期", "交易平台", "買家名稱", "寄送方式", "取貨地點", 
                        "商品名稱", "數量", "單價(售)", "單價(進)", 
                        "總銷售額", "總成本", "分攤手續費", "扣費項目", "總淨利", "毛利率"
                    ]
                    df_sales = pd.DataFrame(columns=cols_sales)
                    df_sales.to_excel(writer, sheet_name='銷售紀錄', index=False)
                    
                    cols_prods = ["分類Tag", "商品名稱", "預設成本", "目前庫存", "最後更新時間"]
                    df_prods = pd.DataFrame(columns=cols_prods)
                    df_prods.loc[0] = ["範例分類", "範例商品A", 100, 10, datetime.now().strftime("%Y-%m-%d %H:%M")]
                    df_prods.to_excel(writer, sheet_name='商品資料', index=False)
            except Exception as e:
                messagebox.showerror("錯誤", f"無法建立 Excel 檔案: {e}")

    def load_products(self):
        try:
            df = pd.read_excel(FILE_NAME, sheet_name='商品資料')
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
        self.tab_products = ttk.Frame(tab_control)
        self.tab_backup = ttk.Frame(tab_control) 
        self.tab_about = ttk.Frame(tab_control)
        
        tab_control.add(self.tab_sales, text='銷售輸入 & 庫存')
        tab_control.add(self.tab_products, text='商品資料管理')
        tab_control.add(self.tab_backup, text='☁️ 雲端備份還原') 
        tab_control.add(self.tab_about, text='設定與關於')
        
        tab_control.pack(expand=1, fill="both")
        
        self.setup_sales_tab()
        self.setup_product_tab()
        self.setup_backup_tab() 
        self.setup_about_tab()

    # ================= 備份還原頁面 =================
    def setup_backup_tab(self):
        frame = ttk.Frame(self.tab_backup, padding=20)
        frame.pack(fill="both", expand=True)

        auth_frame = ttk.LabelFrame(frame, text="1. Google 帳號連結", padding=15)
        auth_frame.pack(fill="x", pady=10)
        
        self.lbl_auth_status = ttk.Label(auth_frame, text="狀態: 尚未連結", foreground="red")
        self.lbl_auth_status.pack(side="left", padx=10)
        
        self.btn_login = ttk.Button(auth_frame, text="登入 Google 帳號", command=self.start_login_thread)
        self.btn_login.pack(side="right")

        op_frame = ttk.LabelFrame(frame, text="2. 檔案備份與還原 (自動存入「蝦皮進銷存系統_備份」)", padding=15)
        op_frame.pack(fill="both", expand=True, pady=10)

        up_frame = ttk.Frame(op_frame)
        up_frame.pack(fill="x", pady=5)
        ttk.Label(up_frame, text="將目前的 Excel 檔案備份到雲端 (建議每日執行):").pack(side="left")
        
        self.btn_upload = ttk.Button(up_frame, text="⬆️ 上傳備份", command=self.start_upload_thread)
        self.btn_upload.pack(side="right")

        ttk.Separator(op_frame, orient="horizontal").pack(fill="x", pady=15)

        ttk.Label(op_frame, text="3. 歷史備份紀錄 (雙擊項目可還原):").pack(anchor="w")
        
        cols = ("檔名", "備份時間")
        self.tree_backup = ttk.Treeview(op_frame, columns=cols, show='headings', height=10)
        self.tree_backup.heading("檔名", text="備份檔名")
        self.tree_backup.column("檔名", width=400)
        self.tree_backup.heading("備份時間", text="建立時間 (已轉為台灣時間)")
        self.tree_backup.column("備份時間", width=200)
        self.tree_backup.pack(fill="both", expand=True, pady=5)
        
        self.tree_backup.bind("<Double-1>", self.action_restore_backup)

        self.btn_refresh = ttk.Button(op_frame, text="🔄 重新整理列表", command=self.start_list_thread)
        self.btn_refresh.pack(fill="x", pady=5)


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

        # === 核心驗證邏輯 ===
        # 這裡的 SALT 必須跟您的生成器完全一樣
        SECRET_SALT = "My_Super_Secret_Salt_Key_2026"
        
        # 軟體自己算一次正確答案
        raw_string = user_id + SECRET_SALT
        expected_code = hashlib.md5(raw_string.encode()).hexdigest()[:8].upper()
        
        # 比對客戶輸入的 跟 算出來的 是否一致
        if input_code == expected_code:
            self.is_vip = True
            messagebox.showinfo("成功", "VIP 功能已解鎖！\n請接著進行 Google 帳號登入。")
            
            # 解鎖按鈕
            self.btn_login.config(state="normal")
            self.lbl_auth_status.config(text="狀態: 尚未連結 (請點擊登入)", foreground="red")
            if self.drive_manager.is_authenticated:
                 self.btn_upload.config(state="normal")
                 
            # (進階) 這裡可以把 user_id 和 code 存到一個本地文件 config.ini
            # 下次打開程式自動讀取並驗證，不用每次都輸入
        else:
            messagebox.showerror("錯誤", "啟用碼錯誤或是帳號不符！\n請聯繫開發者獲取正確授權。")

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
        
        f2 = ttk.Frame(fee_frame)
        f2.pack(fill="x", pady=5)
        
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

    def setup_about_tab(self):
        frame = ttk.Frame(self.tab_about, padding=40)
        frame.pack(expand=True, fill="both")

        font_frame = ttk.LabelFrame(frame, text="介面顯示設定 (字體放大)", padding=15)
        font_frame.pack(fill="x", pady=10)
        
        ttk.Label(font_frame, text="調整字型大小 (10-20):").pack(side="left", padx=5)
        spin = ttk.Spinbox(font_frame, from_=10, to=20, textvariable=self.var_font_size, width=5, command=self.change_font_size)
        spin.pack(side="left", padx=5)
        spin.bind('<KeyRelease>', self.change_font_size)
        
        ttk.Label(font_frame, text="(調整後表格行高會自動變更)", foreground="gray").pack(side="left", padx=10)


        ttk.Label(frame, text="關於本軟體", font=("微軟正黑體", 20, "bold")).pack(pady=10)
        intro_text = "本系統專為個人賣家設計，整合進銷存管理與蝦皮費用試算。\n\n[新增功能]\n1. Google 雲端備份 (多執行緒不卡頓)\n2. 自動建立專屬備份資料夾\n3. 字體大小調整 (長輩友善)\n4. 備份時間自動修正為台灣時間"
        ttk.Label(frame, text=intro_text, font=("微軟正黑體", 12), justify="center").pack(pady=20)
        
        contact_frame = ttk.LabelFrame(frame, text="聯絡資訊", padding=20)
        contact_frame.pack(fill="x", padx=50, pady=10)
        ttk.Label(contact_frame, text="程式設計者: redmaple", font=("微軟正黑體", 11)).pack(anchor="w", pady=5)
        ttk.Label(contact_frame, text="聯絡信箱: az062596216@gmail.com", font=("微軟正黑體", 11)).pack(anchor="w", pady=5)
        
        ttk.Label(frame, text="Version 3.6 (Timezone Fix)", foreground="gray").pack(side="bottom", pady=20)

    # ---------------- 邏輯功能區 ----------------

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
            t_sales = sum(i['total_sales'] for i in self.cart_data)
            t_cost = sum(i['total_cost'] for i in self.cart_data)
            
            raw_rate = self.var_fee_rate_str.get()
            rate = 0.0
            try: rate = float(raw_rate)
            except ValueError:
                match = re.search(r"\((\d+\.?\d*)%\)", raw_rate)
                rate = float(match.group(1)) if match else 0.0

            try: extra = float(self.var_extra_fee.get())
            except: extra = 0.0
            
            fee = (t_sales * (rate/100)) + extra
            income = t_sales - fee
            profit = income - t_cost
            
            self.lbl_gross.config(text=f"總金額: ${t_sales:,.0f}")
            self.lbl_fee.config(text=f"扣費: -${fee:,.1f}")
            self.lbl_income.config(text=f"預估入帳: ${income:,.1f}")
            self.lbl_profit.config(text=f"實收淨利: ${profit:,.1f}")
            return t_sales, fee
        except: return 0, 0

    def submit_order(self):
        if not self.cart_data: return
        
        cust_name = self.var_cust_name.get() if self.var_enable_cust.get() else ""
        cust_loc = self.var_cust_loc.get() if self.var_enable_cust.get() else ""
        ship_method = self.var_ship_method.get() if self.var_enable_cust.get() else ""
        platform_name = self.var_platform.get() if self.var_enable_cust.get() else "" 
        
        t_sales, t_fee = self.update_totals()
        fee_tag = self.var_fee_tag.get()
        try: extra_val = float(self.var_extra_fee.get())
        except: extra_val = 0
        if extra_val > 0 and not fee_tag: fee_tag = "其他"
        elif extra_val == 0: fee_tag = ""

        try:
            rows = []
            date_str = self.var_date.get()
            out_of_stock_warnings = [] 

            df_prods_current = pd.read_excel(FILE_NAME, sheet_name='商品資料')

            for i, item in enumerate(self.cart_data):
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
                
                margin_pct = 0.0
                if item['total_sales'] > 0:
                    margin_pct = (net / item['total_sales']) * 100
                
                rows.append({
                    "日期": row_date, 
                    "交易平台": row_platform, 
                    "買家名稱": row_buyer, 
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
                    "毛利率": f"{margin_pct:.1f}%"
                })

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

            with pd.ExcelWriter(FILE_NAME, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df_prods_current = df_prods_current.sort_values(by=['分類Tag', '商品名稱'], na_position='last')
                df_prods_current.to_excel(writer, sheet_name='商品資料', index=False)

            df_sales_new = pd.DataFrame(rows)
            with pd.ExcelWriter(FILE_NAME, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                try:
                    df_ex = pd.read_excel(FILE_NAME, sheet_name='銷售紀錄')
                    start_row = len(df_ex) + 1
                    header = False
                except:
                    start_row = 0
                    header = True
                df_sales_new.to_excel(writer, sheet_name='銷售紀錄', index=False, header=header, startrow=start_row)

            self.products_df = df_prods_current
            self.update_sales_prod_list()
            self.update_mgmt_prod_list()

            msg = "訂單已儲存！庫存已更新。"
            if out_of_stock_warnings:
                msg += "\n\n⚠️ 注意！以下商品已售完或庫存不足：\n" + "\n".join(out_of_stock_warnings)
            
            messagebox.showinfo("成功", msg)

            self.cart_data = []
            for i in self.tree.get_children(): self.tree.delete(i)
            self.update_totals()
            self.var_cust_name.set("")
            self.var_cust_loc.set("")
            self.var_ship_method.set("")
            self.var_sel_stock_info.set("--")

        except PermissionError: messagebox.showerror("錯誤", "Excel 檔案未關閉，無法寫入！")
        except Exception as e: messagebox.showerror("錯誤", f"發生未預期錯誤: {str(e)}")

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
            df_old = pd.read_excel(FILE_NAME, sheet_name='商品資料')
            idx = df_old[df_old['商品名稱'] == name].index
            if not idx.empty:
                df_old.loc[idx, '分類Tag'] = new_tag
                df_old.loc[idx, '預設成本'] = new_cost
                df_old.loc[idx, '目前庫存'] = new_stock 
                df_old.loc[idx, '最後更新時間'] = now_str
                df_old = df_old.sort_values(by=['分類Tag', '商品名稱'], na_position='last')

                with pd.ExcelWriter(FILE_NAME, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                     df_old.to_excel(writer, sheet_name='商品資料', index=False)
                self.products_df = df_old
                self.update_sales_prod_list() 
                self.update_mgmt_prod_list()
                self.var_upd_time.set(now_str) 
                messagebox.showinfo("成功", f"已更新：{name} (目前庫存: {new_stock})")
        except PermissionError: messagebox.showerror("錯誤", "Excel 未關閉！")

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
