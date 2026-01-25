import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from datetime import datetime
import os
import re

# 設定 Excel 檔案名稱
FILE_NAME = 'sales_data.xlsx'

# 台灣縣市列表
TAIWAN_CITIES = [
    "基隆市", "臺北市", "新北市", "桃園市", "新竹市", "新竹縣", "苗栗縣",
    "臺中市", "彰化縣", "南投縣", "雲林縣", "嘉義市", "嘉義縣", "臺南市",
    "高雄市", "屏東縣", "宜蘭縣", "花蓮縣", "臺東縣", "澎湖縣", "金門縣", "連江縣",
    "海外", "面交"
]

# 交易平台列表 (來源)
PLATFORM_OPTIONS = [
    "蝦皮購物", "賣貨便(7-11)", "好賣家(全家)", "旋轉拍賣", 
    "官方網站", "Facebook社團", "IG", "PChome", "Momo", "實體店面/面交"
]

# 寄送方式列表 (純物流)
SHIPPING_METHODS = [
    "7-11", "全家", "萊爾富", "OK超商", "蝦皮店到店", 
    "蝦皮店到店-隔日到貨", "蝦皮店到宅",
    "黑貓宅急便", "新竹物流", "郵局掛號", "賣家宅配", "面交/自取"
]

# 蝦皮 2026/1/1 後新版手續費方案
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

class SalesApp:
    def __init__(self, root):
        self.root = root
        self.root.title("蝦皮/網拍進銷存系統 (OMS + 庫存管理 + 多平台排序版)")
        self.root.geometry("1280x850") 

        # --- 變數初始化 ---
        self.var_date = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        self.var_search = tk.StringVar()
        
        # 商品選擇暫存
        self.var_sel_name = tk.StringVar()
        self.var_sel_cost = tk.DoubleVar(value=0)
        self.var_sel_price = tk.DoubleVar(value=0)
        self.var_sel_qty = tk.IntVar(value=1)
        self.var_sel_stock_info = tk.StringVar(value="--") 
        
        # 訂單費用
        self.var_fee_rate_str = tk.StringVar() 
        self.var_extra_fee = tk.DoubleVar(value=0.0)
        self.var_fee_tag = tk.StringVar()

        # 顧客與平台資料
        self.var_enable_cust = tk.BooleanVar(value=False)
        self.var_platform = tk.StringVar() 
        self.var_cust_name = tk.StringVar()
        self.var_cust_loc = tk.StringVar()
        self.var_ship_method = tk.StringVar()

        # 購物車
        self.cart_data = []

        # --- 後台管理變數 ---
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

        # 檢查 Excel & 載入資料
        self.check_excel_file()
        self.products_df = self.load_products()
        
        # 建立 UI
        self.create_tabs()

    def check_excel_file(self):
        if not os.path.exists(FILE_NAME):
            try:
                with pd.ExcelWriter(FILE_NAME, engine='openpyxl') as writer:
                    # 銷售紀錄表
                    cols_sales = [
                        "日期", "交易平台", "買家名稱", "寄送方式", "取貨地點", 
                        "商品名稱", "數量", "單價(售)", "單價(進)", 
                        "總銷售額", "總成本", "分攤手續費", "扣費項目", "總淨利", "毛利率"
                    ]
                    df_sales = pd.DataFrame(columns=cols_sales)
                    df_sales.to_excel(writer, sheet_name='銷售紀錄', index=False)
                    
                    # 商品資料表
                    cols_prods = ["分類Tag", "商品名稱", "預設成本", "目前庫存", "最後更新時間"]
                    df_prods = pd.DataFrame(columns=cols_prods)
                    # 範例資料
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
            
            # [新增] 讀取時自動排序，確保 UI 顯示整齊
            df = df.sort_values(by=['分類Tag', '商品名稱'], na_position='last')
            return df
        except:
            return pd.DataFrame(columns=["分類Tag", "商品名稱", "預設成本", "目前庫存", "最後更新時間"])

    def create_tabs(self):
        tab_control = ttk.Notebook(self.root)
        self.tab_sales = ttk.Frame(tab_control)
        self.tab_products = ttk.Frame(tab_control)
        self.tab_about = ttk.Frame(tab_control)
        
        tab_control.add(self.tab_sales, text='銷售輸入 & 庫存扣除')
        tab_control.add(self.tab_products, text='商品資料 & 庫存管理')
        tab_control.add(self.tab_about, text='關於開發者')
        
        tab_control.pack(expand=1, fill="both")
        
        self.setup_sales_tab()
        self.setup_product_tab()
        self.setup_about_tab()

    # ================= 1. 銷售輸入頁面 =================
    def setup_sales_tab(self):
        # Top: Info
        top_frame = ttk.LabelFrame(self.tab_sales, text="訂單基本資料", padding=10)
        top_frame.pack(fill="x", padx=10, pady=5)

        # 第一排：日期、啟用開關
        r1 = ttk.Frame(top_frame)
        r1.pack(fill="x", pady=2)
        ttk.Label(r1, text="訂單日期:").pack(side="left")
        ttk.Entry(r1, textvariable=self.var_date, width=12).pack(side="left", padx=5)
        
        chk = ttk.Checkbutton(r1, text="填寫訂單來源與顧客資料", variable=self.var_enable_cust, command=self.toggle_cust_info)
        chk.pack(side="left", padx=20)

        # 第二排：平台、買家 (使用 Grid 排版比較整齊)
        self.cust_frame = ttk.Frame(top_frame)
        self.cust_frame.pack(fill="x", pady=5)
        
        # 交易平台輸入
        ttk.Label(self.cust_frame, text="交易平台:").grid(row=0, column=0, sticky="w", padx=2)
        self.combo_platform = ttk.Combobox(self.cust_frame, textvariable=self.var_platform, values=PLATFORM_OPTIONS, state="readonly", width=14)
        self.combo_platform.grid(row=0, column=1, padx=5)
        self.combo_platform.set("蝦皮購物") # 預設值

        ttk.Label(self.cust_frame, text="買家名稱(ID):").grid(row=0, column=2, sticky="w", padx=10)
        self.entry_cust_name = ttk.Entry(self.cust_frame, textvariable=self.var_cust_name, width=15)
        self.entry_cust_name.grid(row=0, column=3, padx=5)

        # 第三排：物流、地點
        ttk.Label(self.cust_frame, text="物流方式:").grid(row=1, column=0, sticky="w", padx=2, pady=5)
        self.combo_ship = ttk.Combobox(self.cust_frame, textvariable=self.var_ship_method, values=SHIPPING_METHODS, state="readonly", width=14)
        self.combo_ship.grid(row=1, column=1, padx=5, pady=5)
        self.combo_ship.bind("<<ComboboxSelected>>", self.on_ship_method_change)

        ttk.Label(self.cust_frame, text="取貨縣市:").grid(row=1, column=2, sticky="w", padx=10, pady=5)
        self.combo_loc = ttk.Combobox(self.cust_frame, textvariable=self.var_cust_loc, values=TAIWAN_CITIES, width=13)
        self.combo_loc.grid(row=1, column=3, padx=5, pady=5)
        self.combo_loc.bind('<KeyRelease>', self.filter_cities)

        self.toggle_cust_info()

        # Middle: Split View
        paned = ttk.PanedWindow(self.tab_sales, orient=tk.HORIZONTAL)
        paned.pack(fill="both", expand=True, padx=10, pady=5)

        # Left: Product Select
        left_frame = ttk.LabelFrame(paned, text="加入商品 (即時庫存查詢)", padding=10)
        paned.add(left_frame, weight=1)

        ttk.Label(left_frame, text="搜尋商品:").pack(anchor="w")
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

        # Details
        detail_frame = ttk.Frame(left_frame)
        detail_frame.pack(fill="x", pady=5)
        
        grid_opts = {'sticky': 'w', 'padx': 2, 'pady': 2}
        ttk.Label(detail_frame, text="已選商品:").grid(row=0, column=0, **grid_opts)
        ttk.Entry(detail_frame, textvariable=self.var_sel_name, state='readonly').grid(row=0, column=1, sticky="ew")
        
        # 顯示庫存量
        ttk.Label(detail_frame, text="目前庫存:").grid(row=1, column=0, **grid_opts)
        lbl_stock = ttk.Label(detail_frame, textvariable=self.var_sel_stock_info, foreground="blue", font=("bold", 10))
        lbl_stock.grid(row=1, column=1, sticky="w", padx=2)

        ttk.Label(detail_frame, text="售價(單):").grid(row=2, column=0, **grid_opts)
        ttk.Entry(detail_frame, textvariable=self.var_sel_price).grid(row=2, column=1, sticky="ew")

        ttk.Label(detail_frame, text="購買數量:").grid(row=3, column=0, **grid_opts)
        ttk.Entry(detail_frame, textvariable=self.var_sel_qty).grid(row=3, column=1, sticky="ew")


        ttk.Label(detail_frame, text="成本(單):").grid(row=4, column=0, **grid_opts)
        ttk.Entry(detail_frame, textvariable=self.var_sel_cost).grid(row=4, column=1, sticky="ew")

        ttk.Button(detail_frame, text="加入清單 ->", command=self.add_to_cart).grid(row=5, column=0, columnspan=2, pady=10, sticky="ew")

        # Right: Cart
        right_frame = ttk.LabelFrame(paned, text="訂單內容 (送出後自動扣庫存)", padding=10)
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

        ttk.Button(right_frame, text="(x) 移除選中項目", command=self.remove_from_cart).pack(anchor="e", pady=2)

        # === 費用設定 ===
        fee_frame = ttk.LabelFrame(right_frame, text="手續費與其他扣款 (2026新制)", padding=10)
        fee_frame.pack(fill="x", pady=5)
        
        f1 = ttk.Frame(fee_frame)
        f1.pack(fill="x")
        ttk.Label(f1, text="平台手續費率:").pack(side="left")
        
        self.combo_fee_rate = ttk.Combobox(f1, textvariable=self.var_fee_rate_str, values=SHOPEE_FEE_OPTIONS, width=28)
        self.combo_fee_rate.pack(side="left", padx=5)
        self.combo_fee_rate.set("一般賣家-平日 (14.5%)") # 預設值
        self.combo_fee_rate.bind('<<ComboboxSelected>>', self.on_fee_option_selected)
        self.combo_fee_rate.bind('<KeyRelease>', self.update_totals_event)
        
        f2 = ttk.Frame(fee_frame)
        f2.pack(fill="x", pady=5)
        
        tag_opts = ["", "活動費", "運費補貼", "補償金額", "私人預定", "補寄補貼", "固定成本(如包材/出貨)"]
        self.combo_tag = ttk.Combobox(f2, textvariable=self.var_fee_tag, values=tag_opts, state="readonly", width=12)
        self.combo_tag.pack(side="left")
        self.combo_tag.set("扣費原因")

        ttk.Label(f2, text=" 金額$").pack(side="left", padx=2)
        e_extra = ttk.Entry(f2, textvariable=self.var_extra_fee, width=8)
        e_extra.pack(side="left")
        e_extra.bind('<KeyRelease>', self.update_totals_event)
        
        ttk.Label(f2, text="(如:負擔運費60)", foreground="gray", font=("微軟正黑體", 8)).pack(side="left", padx=2)

        # Summary
        sum_frame = ttk.Frame(right_frame, relief="groove", padding=5)
        sum_frame.pack(fill="x", side="bottom")
        
        self.lbl_gross = ttk.Label(sum_frame, text="總金額: $0",font=("bold", 11))
        self.lbl_gross.pack(anchor="w")
        self.lbl_fee = ttk.Label(sum_frame, text="扣費: $0", foreground="blue", font=("bold", 11))
        self.lbl_fee.pack(anchor="w")
        self.lbl_profit = ttk.Label(sum_frame, text="實收淨利: $0", foreground="green", font=("bold", 12))
        self.lbl_profit.pack(anchor="w")
        self.lbl_income = ttk.Label(sum_frame, text="預估入帳: $0", foreground="#ff0800", font=("bold", 12))
        self.lbl_income.pack(anchor="w")


        ttk.Button(sum_frame, text="✔ 確認送出並寫入 Excel", command=self.submit_order).pack(fill="x", pady=5)


    # ================= 2. 商品管理頁面 =================
    def setup_product_tab(self):
        paned = ttk.PanedWindow(self.tab_products, orient=tk.HORIZONTAL)
        paned.pack(fill="both", expand=True, padx=10, pady=10)

        # === 左側：新增商品 ===
        frame_add = ttk.LabelFrame(paned, text="【新增】新商品入庫", padding=15)
        paned.add(frame_add, weight=1)

        ttk.Label(frame_add, text="1. 選擇或輸入分類Tag:", font=("bold", 10)).pack(anchor="w", pady=(0,5))
        self.combo_add_tag = ttk.Combobox(frame_add, textvariable=self.var_add_tag)
        self.combo_add_tag.pack(fill="x", pady=5)
        self.combo_add_tag.bind('<Button-1>', self.load_existing_tags)

        ttk.Label(frame_add, text="2. 商品名稱:", font=("bold", 10)).pack(anchor="w", pady=(10,5))
        ttk.Entry(frame_add, textvariable=self.var_add_name).pack(fill="x", pady=5)

        ttk.Label(frame_add, text="3. 進貨成本:", font=("bold", 10)).pack(anchor="w", pady=(10,5))
        ttk.Entry(frame_add, textvariable=self.var_add_cost).pack(fill="x", pady=5)
        
        # [新增] 初始庫存
        ttk.Label(frame_add, text="4. 初始庫存數量:", font=("bold", 10)).pack(anchor="w", pady=(10,5))
        ttk.Entry(frame_add, textvariable=self.var_add_stock).pack(fill="x", pady=5)

        ttk.Button(frame_add, text="+ 新增至資料庫", command=self.submit_new_product).pack(fill="x", pady=20)

        # === 右側：更新商品 ===
        frame_upd = ttk.LabelFrame(paned, text="【更新】維護既有商品 (含補貨)", padding=15)
        paned.add(frame_upd, weight=1)

        ttk.Label(frame_upd, text="搜尋商品關鍵字:", font=("bold", 10)).pack(anchor="w")
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

        ttk.Label(edit_frame, text="商品名稱 (不可改):").grid(row=0, column=0, sticky="w")
        ttk.Entry(edit_frame, textvariable=self.var_upd_name, state="readonly").grid(row=0, column=1, sticky="ew", padx=5)

        ttk.Label(edit_frame, text="分類Tag:").grid(row=1, column=0, sticky="w", pady=5)
        self.combo_upd_tag = ttk.Combobox(edit_frame, textvariable=self.var_upd_tag, width=18)
        self.combo_upd_tag.grid(row=1, column=1, sticky="ew", padx=5, pady=5)
        self.combo_upd_tag.bind('<Button-1>', self.load_existing_tags)

        ttk.Label(edit_frame, text="成本調整:").grid(row=2, column=0, sticky="w", pady=5)
        ttk.Entry(edit_frame, textvariable=self.var_upd_cost).grid(row=2, column=1, sticky="ew", padx=5, pady=5)
        
        # [新增] 修改庫存
        ttk.Label(edit_frame, text="目前庫存(補貨):").grid(row=3, column=0, sticky="w", pady=5)
        ttk.Entry(edit_frame, textvariable=self.var_upd_stock).grid(row=3, column=1, sticky="ew", padx=5, pady=5)

        ttk.Label(edit_frame, text="上次更新:").grid(row=4, column=0, sticky="w")
        ttk.Label(edit_frame, textvariable=self.var_upd_time, foreground="gray").grid(row=4, column=1, sticky="w", padx=5)

        btn_frame = ttk.Frame(edit_frame)
        btn_frame.grid(row=5, column=0, columnspan=2, pady=10, sticky="ew")
        
        ttk.Button(btn_frame, text="💾 儲存變更", command=self.submit_update_product).pack(side="left", fill="x", expand=True, padx=(0, 5))
        ttk.Button(btn_frame, text="🗑️ 刪除商品", command=self.delete_product).pack(side="left", fill="x", expand=True, padx=(5, 0))

        self.update_mgmt_prod_list()

    # ================= 3. 關於開發者頁面 =================
    def setup_about_tab(self):
        frame = ttk.Frame(self.tab_about, padding=40)
        frame.pack(expand=True, fill="both")
        ttk.Label(frame, text="關於本軟體", font=("微軟正黑體", 20, "bold")).pack(pady=10)
        intro_text = "本系統專為個人賣家設計，整合進銷存管理與蝦皮費用試算功能。"
        ttk.Label(frame, text=intro_text, font=("微軟正黑體", 12), justify="center").pack(pady=20)
        contact_frame = ttk.LabelFrame(frame, text="聯絡開發者", padding=20)
        contact_frame.pack(fill="x", padx=50, pady=10)
        ttk.Label(contact_frame, text="程式設計者: redmaple", font=("微軟正黑體", 11)).pack(anchor="w", pady=5)
        ttk.Label(contact_frame, text="聯絡信箱: az062596216@gmail.com", font=("微軟正黑體", 11)).pack(anchor="w", pady=5)
        license_frame = ttk.LabelFrame(frame, text="使用與授權聲明", padding=20)
        license_frame.pack(fill="x", padx=50, pady=10)
        license_text = "● 本軟體以開源 (Open Source) 精神發布，永久免費供個人使用。\n● 軟體按「現狀」提供，請務必定期備份 Excel 檔案。 \n● 開發者不對使用本軟體所產生的任何直接或間接損失負責。\n● 未經授權禁止商業販售本軟體。"
        ttk.Label(license_frame, text=license_text, font=("微軟正黑體", 10), foreground="#555", justify="left").pack(anchor="w")
        ttk.Label(frame, text="Version 3.2 (Product Sorting)", foreground="gray").pack(side="bottom", pady=20)

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
                try:
                    p_stock = int(row['目前庫存'])
                except:
                    p_stock = 0
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
                try:
                    stock = int(record.iloc[0]['目前庫存'])
                except:
                    stock = 0
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

            # 檢查庫存
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

    # 【核心功能】 送出訂單：包含資料留白、平台欄位、庫存修正、毛利、**商品排序**
    def submit_order(self):
        if not self.cart_data: return
        
        # 取得表單資料
        cust_name = self.var_cust_name.get() if self.var_enable_cust.get() else ""
        cust_loc = self.var_cust_loc.get() if self.var_enable_cust.get() else ""
        ship_method = self.var_ship_method.get() if self.var_enable_cust.get() else ""
        platform_name = self.var_platform.get() if self.var_enable_cust.get() else "" 
        
        t_sales, t_fee = self.update_totals()
        fee_tag = self.var_fee_tag.get()
        extra_val = 0
        try: extra_val = float(self.var_extra_fee.get())
        except: pass
        if extra_val > 0 and not fee_tag: fee_tag = "其他"
        elif extra_val == 0: fee_tag = ""

        try:
            # 1. 準備寫入銷售紀錄
            rows = []
            date_str = self.var_date.get()
            out_of_stock_warnings = [] 

            # 讀取最新的商品資料
            df_prods_current = pd.read_excel(FILE_NAME, sheet_name='商品資料')

            for i, item in enumerate(self.cart_data):
                # 資料留白邏輯 (第一筆顯示，其餘留白)
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

                # 費用分攤計算
                ratio = item['total_sales'] / t_sales if t_sales > 0 else 0
                alloc_fee = t_fee * ratio
                net = item['total_sales'] - item['total_cost'] - alloc_fee
                
                # 計算毛利率
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

                # --- 庫存扣除邏輯 (含 Bug 修正) ---
                prod_name = item['name']
                sold_qty = item['qty']
                
                idxs = df_prods_current[df_prods_current['商品名稱'] == prod_name].index
                
                if not idxs.empty:
                    target_idx = idxs[0]
                    raw_stock = df_prods_current.at[target_idx, '目前庫存']
                    try:
                        current = int(raw_stock)
                    except (ValueError, TypeError):
                        current = 0
                        
                    new_stock = current - sold_qty
                    df_prods_current.at[target_idx, '目前庫存'] = new_stock
                    
                    if new_stock <= 0:
                        out_of_stock_warnings.append(f"● {prod_name} (剩餘: {new_stock})")

            # 3. 寫入 Excel
            with pd.ExcelWriter(FILE_NAME, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                # 【新增】寫入商品表前，依分類+名稱排序
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

            # 4. 更新記憶體
            self.products_df = df_prods_current
            self.update_sales_prod_list()
            self.update_mgmt_prod_list()

            # 5. 結果通知
            msg = "訂單已儲存！庫存已更新。"
            if out_of_stock_warnings:
                msg += "\n\n⚠️ 注意！以下商品已售完或庫存不足：\n" + "\n".join(out_of_stock_warnings)
            
            messagebox.showinfo("成功", msg)

            # Reset
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
                
                try:
                    current_stock = int(row['目前庫存'])
                except (ValueError, TypeError):
                    current_stock = 0
                    
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
            
            # 【新增】排序
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
                
                # 【新增】排序
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
