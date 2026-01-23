import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from datetime import datetime
import os
import webbrowser # 用於開啟超連結(如果未來需要)

# 設定 Excel 檔案名稱
FILE_NAME = 'sales_data.xlsx'

# 台灣縣市列表
TAIWAN_CITIES = [
    "基隆市", "臺北市", "新北市", "桃園市", "新竹市", "新竹縣", "苗栗縣",
    "臺中市", "彰化縣", "南投縣", "雲林縣", "嘉義市", "嘉義縣", "臺南市",
    "高雄市", "屏東縣", "宜蘭縣", "花蓮縣", "臺東縣", "澎湖縣", "金門縣", "連江縣",
    "海外", "面交"
]

# 寄送方式列表
SHIPPING_METHODS = [
    "7-11", "全家", "蝦皮店到店", "蝦皮店到店-隔日到貨", "蝦皮店到宅",
    "黑貓宅急便", "新竹物流", "郵局掛號", "賣貨便(7-11)", "好賣家(全家)", "面交"
]

class SalesApp:
    def __init__(self, root):
        self.root = root
        self.root.title("蝦皮/網拍銷售記錄系統 (OMS 完整版)")
        self.root.geometry("1200x800") 

        # --- 變數初始化 ---
        self.var_date = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        self.var_search = tk.StringVar()
        
        # 商品選擇暫存 (銷售頁面用)
        self.var_sel_name = tk.StringVar()
        self.var_sel_cost = tk.DoubleVar(value=0)
        self.var_sel_price = tk.DoubleVar(value=0)
        self.var_sel_qty = tk.IntVar(value=1)
        
        # 訂單費用
        self.var_fee_rate = tk.DoubleVar(value=0.0)
        self.var_extra_fee = tk.DoubleVar(value=0.0)
        self.var_fee_tag = tk.StringVar()

        # 顧客資料
        self.var_enable_cust = tk.BooleanVar(value=False)
        self.var_cust_name = tk.StringVar()
        self.var_cust_loc = tk.StringVar()
        self.var_ship_method = tk.StringVar()

        # 購物車
        self.cart_data = []

        # --- 後台管理變數 ---
        # 左側：新增用
        self.var_add_tag = tk.StringVar()
        self.var_add_name = tk.StringVar()
        self.var_add_cost = tk.DoubleVar(value=0)
        
        # 右側：更新用
        self.var_mgmt_search = tk.StringVar() # 搜尋框
        self.var_upd_tag = tk.StringVar()
        self.var_upd_name = tk.StringVar() # 唯讀，作為Key
        self.var_upd_cost = tk.DoubleVar(value=0)
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
                        "日期", "買家名稱", "寄送方式", "取貨地點", 
                        "商品名稱", "數量", "單價(售)", "單價(進)", 
                        "總銷售額", "總成本", "分攤手續費", "扣費項目", "總淨利"
                    ]
                    df_sales = pd.DataFrame(columns=cols_sales)
                    df_sales.to_excel(writer, sheet_name='銷售紀錄', index=False)
                    
                    # 商品資料表
                    cols_prods = ["分類Tag", "商品名稱", "預設成本", "最後更新時間"]
                    df_prods = pd.DataFrame(columns=cols_prods)
                    df_prods.loc[0] = ["範例分類", "範例商品A", 100, datetime.now().strftime("%Y-%m-%d %H:%M")]
                    df_prods.to_excel(writer, sheet_name='商品資料', index=False)
            except Exception as e:
                messagebox.showerror("錯誤", f"無法建立 Excel 檔案: {e}")

    def load_products(self):
        try:
            df = pd.read_excel(FILE_NAME, sheet_name='商品資料')
            if "分類Tag" not in df.columns: df["分類Tag"] = ""
            return df
        except:
            return pd.DataFrame(columns=["分類Tag", "商品名稱", "預設成本", "最後更新時間"])

    def create_tabs(self):
        tab_control = ttk.Notebook(self.root)
        self.tab_sales = ttk.Frame(tab_control)
        self.tab_products = ttk.Frame(tab_control)
        self.tab_about = ttk.Frame(tab_control) # 新增關於頁面
        
        tab_control.add(self.tab_sales, text='銷售輸入 & 訂單')
        tab_control.add(self.tab_products, text='商品資料庫管理')
        tab_control.add(self.tab_about, text='關於開發者')
        
        tab_control.pack(expand=1, fill="both")
        
        self.setup_sales_tab()
        self.setup_product_tab()
        self.setup_about_tab()

    # ================= 1. 銷售輸入頁面 (維持原樣) =================
    def setup_sales_tab(self):
        # Top: Info
        top_frame = ttk.LabelFrame(self.tab_sales, text="訂單基本資料", padding=10)
        top_frame.pack(fill="x", padx=10, pady=5)

        r1 = ttk.Frame(top_frame)
        r1.pack(fill="x", pady=2)
        ttk.Label(r1, text="訂單日期:").pack(side="left")
        ttk.Entry(r1, textvariable=self.var_date, width=12).pack(side="left", padx=5)

        chk = ttk.Checkbutton(r1, text="填寫顧客/寄送資料", variable=self.var_enable_cust, command=self.toggle_cust_info)
        chk.pack(side="left", padx=20)

        self.cust_frame = ttk.Frame(top_frame)
        self.cust_frame.pack(fill="x", pady=5)
        
        ttk.Label(self.cust_frame, text="買家名稱(ID):").pack(side="left")
        self.entry_cust_name = ttk.Entry(self.cust_frame, textvariable=self.var_cust_name, width=15)
        self.entry_cust_name.pack(side="left", padx=5)

        ttk.Label(self.cust_frame, text="寄送方式:").pack(side="left")
        self.combo_ship = ttk.Combobox(self.cust_frame, textvariable=self.var_ship_method, values=SHIPPING_METHODS, state="readonly", width=18)
        self.combo_ship.pack(side="left", padx=5)
        self.combo_ship.bind("<<ComboboxSelected>>", self.on_ship_method_change)

        ttk.Label(self.cust_frame, text="取貨縣市:").pack(side="left")
        self.combo_loc = ttk.Combobox(self.cust_frame, textvariable=self.var_cust_loc, values=TAIWAN_CITIES, width=10)
        self.combo_loc.pack(side="left", padx=5)
        self.combo_loc.bind('<KeyRelease>', self.filter_cities)

        self.toggle_cust_info()

        # Middle: Split View
        paned = ttk.PanedWindow(self.tab_sales, orient=tk.HORIZONTAL)
        paned.pack(fill="both", expand=True, padx=10, pady=5)

        # Left: Product Select
        left_frame = ttk.LabelFrame(paned, text="加入商品", padding=10)
        paned.add(left_frame, weight=1)

        ttk.Label(left_frame, text="搜尋商品 (名稱/分類):").pack(anchor="w")
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
        
        ttk.Label(detail_frame, text="數量:").grid(row=1, column=0, **grid_opts)
        ttk.Entry(detail_frame, textvariable=self.var_sel_qty).grid(row=1, column=1, sticky="ew")

        ttk.Label(detail_frame, text="蝦皮售價(單):").grid(row=2, column=0, **grid_opts)
        ttk.Entry(detail_frame, textvariable=self.var_sel_price).grid(row=2, column=1, sticky="ew")

        ttk.Label(detail_frame, text="進貨成本(單):").grid(row=3, column=0, **grid_opts)
        ttk.Entry(detail_frame, textvariable=self.var_sel_cost).grid(row=3, column=1, sticky="ew")

        ttk.Button(detail_frame, text="加入清單 ->", command=self.add_to_cart).grid(row=4, column=0, columnspan=2, pady=10, sticky="ew")

        # Right: Cart
        right_frame = ttk.LabelFrame(paned, text="訂單內容與結算", padding=10)
        paned.add(right_frame, weight=2)

        cols = ("商品名稱", "數量", "單價", "總計")
        self.tree = ttk.Treeview(right_frame, columns=cols, show='headings', height=8)
        self.tree.heading("商品名稱", text="商品名稱")
        self.tree.column("商品名稱", width=120)
        self.tree.heading("數量", text="數量")
        self.tree.column("數量", width=40, anchor="center")
        self.tree.heading("單價", text="售價")
        self.tree.column("單價", width=60, anchor="e")
        self.tree.heading("總計", text="小計")
        self.tree.column("總計", width=70, anchor="e")
        self.tree.pack(fill="both", expand=True)

        ttk.Button(right_frame, text="(x) 移除選中項目", command=self.remove_from_cart).pack(anchor="e", pady=2)

        # Fees
        fee_frame = ttk.LabelFrame(right_frame, text="手續費與其他扣款", padding=10)
        fee_frame.pack(fill="x", pady=5)
        
        f1 = ttk.Frame(fee_frame)
        f1.pack(fill="x")
        ttk.Label(f1, text="手續費率 (%):").pack(side="left")
        # 這裡加入提示文字
        ttk.Label(f1, text="(預設蝦皮手續費為14.5%)", foreground="gray", font=("微軟正黑體", 9)).pack(side="right", padx=2)

        e_rate = ttk.Entry(f1, textvariable=self.var_fee_rate, width=5)
        e_rate.pack(side="left", padx=5)
        

        e_rate.bind('<KeyRelease>', self.update_totals_event)

        f2 = ttk.Frame(fee_frame)
        f2.pack(fill="x", pady=2)
        tag_opts = ["", "活動費", "運費補貼", "補償金額", "私人預定", "補寄補貼"]
        self.combo_tag = ttk.Combobox(f2, textvariable=self.var_fee_tag, values=tag_opts, state="readonly", width=10)
        self.combo_tag.pack(side="left")
        ttk.Label(f2, text="$").pack(side="left", padx=2)
        e_extra = ttk.Entry(f2, textvariable=self.var_extra_fee, width=6)
        e_extra.pack(side="left")
        e_extra.bind('<KeyRelease>', self.update_totals_event)

        # Summary
        sum_frame = ttk.Frame(right_frame, relief="groove", padding=5)
        sum_frame.pack(fill="x", side="bottom")

        self.lbl_gross = ttk.Label(sum_frame, text="總金額: $0", font=("微軟正黑體", 10))
        self.lbl_gross.pack(anchor="w")
        self.lbl_fee = ttk.Label(sum_frame, text="扣費: $0", foreground="blue", font=("微軟正黑體", 10))
        self.lbl_fee.pack(anchor="w")
        self.lbl_income = ttk.Label(sum_frame, text="預估入帳: $0", foreground="red", font=("微軟正黑體", 12))
        self.lbl_income.pack(anchor="w")
        self.lbl_profit = ttk.Label(sum_frame, text="實收淨利: $0", foreground="green", font=("微軟正黑體", 12))
        self.lbl_profit.pack(anchor="w")

        ttk.Button(sum_frame, text="✔ 確認送出並寫入 Excel", command=self.submit_order).pack(fill="x", pady=5)

    # ================= 2. 商品管理頁面 (新增/更新 分離版) =================
    def setup_product_tab(self):
        # 使用 PanedWindow 切割左右
        paned = ttk.PanedWindow(self.tab_products, orient=tk.HORIZONTAL)
        paned.pack(fill="both", expand=True, padx=10, pady=10)

        # === 左側：新增商品專區 ===
        frame_add = ttk.LabelFrame(paned, text="【新增】新商品入庫", padding=15)
        paned.add(frame_add, weight=1)

        ttk.Label(frame_add, text="1. 選擇或輸入分類Tag:", font=("bold", 10)).pack(anchor="w", pady=(0,5))
        self.combo_add_tag = ttk.Combobox(frame_add, textvariable=self.var_add_tag)
        self.combo_add_tag.pack(fill="x", pady=5)
        self.combo_add_tag.bind('<Button-1>', self.load_existing_tags)

        ttk.Label(frame_add, text="2. 商品名稱:", font=("bold", 10)).pack(anchor="w", pady=(10,5))
        ttk.Entry(frame_add, textvariable=self.var_add_name).pack(fill="x", pady=5)

        ttk.Label(frame_add, text="3. 預設進貨成本:", font=("bold", 10)).pack(anchor="w", pady=(10,5))
        ttk.Entry(frame_add, textvariable=self.var_add_cost).pack(fill="x", pady=5)

        ttk.Button(frame_add, text="+ 新增至資料庫", command=self.submit_new_product).pack(fill="x", pady=20)
        ttk.Label(frame_add, text="※ 若商品已存在，請使用右側更新功能", foreground="gray", wraplength=300).pack()

        # === 右側：更新商品專區 ===
        frame_upd = ttk.LabelFrame(paned, text="【更新】維護既有商品", padding=15)
        paned.add(frame_upd, weight=1)

        # 搜尋區
        ttk.Label(frame_upd, text="搜尋商品關鍵字:", font=("bold", 10)).pack(anchor="w")
        e_search = ttk.Entry(frame_upd, textvariable=self.var_mgmt_search)
        e_search.pack(fill="x", pady=5)
        e_search.bind('<KeyRelease>', self.update_mgmt_prod_list)

        # 列表區
        list_frame = ttk.Frame(frame_upd)
        list_frame.pack(fill="both", expand=True, pady=5)
        self.listbox_mgmt = tk.Listbox(list_frame, height=10)
        sb = ttk.Scrollbar(list_frame, orient="vertical", command=self.listbox_mgmt.yview)
        self.listbox_mgmt.configure(yscrollcommand=sb.set)
        self.listbox_mgmt.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.listbox_mgmt.bind('<<ListboxSelect>>', self.on_mgmt_prod_select)

        # 編輯區
        edit_frame = ttk.LabelFrame(frame_upd, text="編輯選中商品", padding=10)
        edit_frame.pack(fill="x", pady=10)

        # 顯示商品名稱 (唯讀，確保 Key 不變)
        ttk.Label(edit_frame, text="商品名稱 (不可改):").grid(row=0, column=0, sticky="w")
        ttk.Entry(edit_frame, textvariable=self.var_upd_name, state="readonly").grid(row=0, column=1, sticky="ew", padx=5)

        ttk.Label(edit_frame, text="分類Tag:").grid(row=1, column=0, sticky="w", pady=5)
        self.combo_upd_tag = ttk.Combobox(edit_frame, textvariable=self.var_upd_tag, width=18)
        self.combo_upd_tag.grid(row=1, column=1, sticky="ew", padx=5, pady=5)
        self.combo_upd_tag.bind('<Button-1>', self.load_existing_tags)

        ttk.Label(edit_frame, text="成本調整:").grid(row=2, column=0, sticky="w", pady=5)
        ttk.Entry(edit_frame, textvariable=self.var_upd_cost).grid(row=2, column=1, sticky="ew", padx=5, pady=5)
        
        ttk.Label(edit_frame, text="上次更新:").grid(row=3, column=0, sticky="w")
        ttk.Label(edit_frame, textvariable=self.var_upd_time, foreground="gray").grid(row=3, column=1, sticky="w", padx=5)

        ttk.Button(edit_frame, text="💾 儲存變更", command=self.submit_update_product).grid(row=4, column=0, columnspan=2, pady=10, sticky="ew")

        # 初始化列表
        self.update_mgmt_prod_list()

    # ================= 3. 關於開發者頁面 (新增) =================
    def setup_about_tab(self):
        frame = ttk.Frame(self.tab_about, padding=40)
        frame.pack(expand=True, fill="both")

        # 標題
        ttk.Label(frame, text="關於本軟體", font=("微軟正黑體", 20, "bold")).pack(pady=10)
        
        # 簡介
        intro_text = (
            "歡迎使用蝦皮/網拍銷售記錄系統 (OMS 完整版)！\n\n"
            "本系統為作者本人蝦皮多年銷售經驗設計，旨在簡化每日記帳與訂單管理流程。\n"
            "希望透過輕量化的工具，協助您更有效率地掌握營收狀況。\n\n"
            "如有任何建議或問題，歡迎隨時聯絡我！\n"
        )
        ttk.Label(frame, text=intro_text, font=("微軟正黑體", 12), justify="center").pack(pady=20)

        # 聯絡資訊區塊
        contact_frame = ttk.LabelFrame(frame, text="聯絡開發者", padding=20)
        contact_frame.pack(fill="x", padx=50, pady=10)
        
        ttk.Label(contact_frame, text="程式設計者: 紅楓 ", font=("微軟正黑體", 11)).pack(anchor="w", pady=5)
        ttk.Label(contact_frame, text="聯絡信箱: az062596216@gmail.com", font=("微軟正黑體", 11)).pack(anchor="w", pady=5)
        
        # 開源聲明區塊
        license_frame = ttk.LabelFrame(frame, text="使用與授權聲明", padding=20)
        license_frame.pack(fill="x", padx=50, pady=10)

        license_text = (
            "● 本軟體以開源 (Open Source) 精神發布，永久免費供個人使用。\n"
            "● 禁止將本軟體進行打包販售、營利或做為商業課程教材。\n"
            "● 軟體按「現狀」提供，開發者不對因使用本軟體造成的資料遺失負責，請務必定期備份 Excel 檔案。"
        )
        ttk.Label(license_frame, text=license_text, font=("微軟正黑體", 10), foreground="#555", justify="left").pack(anchor="w")

        # 版本號
        ttk.Label(frame, text="Version 2.1 (OMS Edition)", foreground="gray").pack(side="bottom", pady=20)

    # ---------------- 邏輯功能區 ----------------

    # --- 共用邏輯 ---
    def load_existing_tags(self, event=None):
        if not self.products_df.empty and "分類Tag" in self.products_df.columns:
            tags = self.products_df["分類Tag"].dropna().unique().tolist()
            # 更新所有下拉選單
            self.combo_add_tag['values'] = tags
            self.combo_upd_tag['values'] = tags

    # --- 銷售頁面邏輯 ---
    def toggle_cust_info(self):
        state = "normal" if self.var_enable_cust.get() else "disabled"
        self.entry_cust_name.config(state=state)
        self.combo_ship.config(state="readonly" if state == "normal" else "disabled")
        self.combo_loc.config(state=state)

    def filter_cities(self, event):
        typed = self.var_cust_loc.get()
        if typed == '': self.combo_loc['values'] = TAIWAN_CITIES
        else: self.combo_loc['values'] = [i for i in TAIWAN_CITIES if typed in i]

    def on_ship_method_change(self, event):
        if self.var_ship_method.get() == "面交": self.var_cust_loc.set("面交")
        elif self.var_cust_loc.get() == "面交": self.var_cust_loc.set("")

    def update_sales_prod_list(self, event=None):
        search_term = self.var_search.get().lower()
        self.listbox_sales.delete(0, tk.END)
        if not self.products_df.empty:
            for index, row in self.products_df.iterrows():
                p_name = str(row['商品名稱'])
                p_tag = str(row['分類Tag']) if pd.notna(row['分類Tag']) else "無"
                display_str = f"[{p_tag}] {p_name}"
                if search_term in p_name.lower() or search_term in p_tag.lower():
                    self.listbox_sales.insert(tk.END, display_str)

    def on_sales_prod_select(self, event):
        selection = self.listbox_sales.curselection()
        if selection:
            display_str = self.listbox_sales.get(selection[0])
            selected_name = display_str.split("]", 1)[1].strip() if "]" in display_str else display_str
            self.var_sel_name.set(selected_name)
            self.var_sel_qty.set(1)
            record = self.products_df[self.products_df['商品名稱'] == selected_name]
            if not record.empty:
                self.var_sel_cost.set(record.iloc[0]['預設成本'])
                self.var_sel_price.set(0)

    def add_to_cart(self):
        name = self.var_sel_name.get()
        if not name: return
        try:
            qty = self.var_sel_qty.get()
            cost = self.var_sel_cost.get()
            price = self.var_sel_price.get()
            if qty <= 0: return
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

    def update_totals_event(self, event): self.update_totals()
    
    def update_totals(self):
        try:
            t_sales = sum(i['total_sales'] for i in self.cart_data)
            t_cost = sum(i['total_cost'] for i in self.cart_data)
            try: rate = float(self.var_fee_rate.get())
            except: rate = 0.0
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
        
        t_sales, t_fee = self.update_totals()
        fee_tag = self.var_fee_tag.get()
        extra_val = 0
        try: extra_val = float(self.var_extra_fee.get())
        except: pass
        if extra_val > 0 and not fee_tag: fee_tag = "其他"
        elif extra_val == 0: fee_tag = ""

        try:
            rows = []
            date_str = self.var_date.get()
            for item in self.cart_data:
                ratio = item['total_sales'] / t_sales if t_sales > 0 else 0
                alloc_fee = t_fee * ratio
                net = item['total_sales'] - item['total_cost'] - alloc_fee
                rows.append({
                    "日期": date_str, "買家名稱": cust_name, "寄送方式": ship_method, "取貨地點": cust_loc,
                    "商品名稱": item['name'], "數量": item['qty'], "單價(售)": item['unit_price'], "單價(進)": item['unit_cost'],
                    "總銷售額": item['total_sales'], "總成本": item['total_cost'], "分攤手續費": round(alloc_fee, 2),
                    "扣費項目": fee_tag, "總淨利": round(net, 2)
                })
            df_new = pd.DataFrame(rows)
            with pd.ExcelWriter(FILE_NAME, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                try:
                    df_ex = pd.read_excel(FILE_NAME, sheet_name='銷售紀錄')
                    start_row = len(df_ex) + 1
                    header = False
                except:
                    start_row = 0
                    header = True
                df_new.to_excel(writer, sheet_name='銷售紀錄', index=False, header=header, startrow=start_row)
            messagebox.showinfo("成功", "訂單已儲存！")
            self.cart_data = []
            for i in self.tree.get_children(): self.tree.delete(i)
            self.update_totals()
            self.var_cust_name.set("")
            self.var_cust_loc.set("")
            self.var_ship_method.set("")
        except PermissionError: messagebox.showerror("錯誤", "Excel 檔案未關閉！")
        except Exception as e: messagebox.showerror("錯誤", str(e))

    # --- 商品管理頁面邏輯 (新增/更新) ---
    
    # 1. 右側：更新列表搜尋
    def update_mgmt_prod_list(self, event=None):
        search_term = self.var_mgmt_search.get().lower()
        self.listbox_mgmt.delete(0, tk.END)
        if not self.products_df.empty:
            for index, row in self.products_df.iterrows():
                p_name = str(row['商品名稱'])
                p_tag = str(row['分類Tag']) if pd.notna(row['分類Tag']) else "無"
                display_str = f"[{p_tag}] {p_name}"
                if search_term in p_name.lower() or search_term in p_tag.lower():
                    self.listbox_mgmt.insert(tk.END, display_str)

    # 2. 右側：選擇要編輯的商品
    def on_mgmt_prod_select(self, event):
        selection = self.listbox_mgmt.curselection()
        if selection:
            display_str = self.listbox_mgmt.get(selection[0])
            selected_name = display_str.split("]", 1)[1].strip() if "]" in display_str else display_str
            
            # 填入編輯框
            record = self.products_df[self.products_df['商品名稱'] == selected_name]
            if not record.empty:
                row = record.iloc[0]
                self.var_upd_name.set(row['商品名稱'])
                self.var_upd_tag.set(row['分類Tag'] if pd.notna(row['分類Tag']) else "")
                self.var_upd_cost.set(row['預設成本'])
                self.var_upd_time.set(row['最後更新時間'] if pd.notna(row['最後更新時間']) else "未知")

    # 3. 左側：提交新商品
    def submit_new_product(self):
        name = self.var_add_name.get().strip()
        cost = self.var_add_cost.get()
        tag = self.var_add_tag.get().strip()
        
        if not name:
            messagebox.showwarning("警告", "請輸入商品名稱")
            return

        # 檢查是否重複
        if name in self.products_df['商品名稱'].values:
            messagebox.showwarning("已存在", f"商品「{name}」已存在於資料庫中。\n請使用右側「更新」功能來修改價格。")
            return

        # 寫入
        try:
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
            new_row = pd.DataFrame([{"分類Tag": tag, "商品名稱": name, "預設成本": cost, "最後更新時間": now_str}])
            df_old = pd.read_excel(FILE_NAME, sheet_name='商品資料')
            df_updated = pd.concat([df_old, new_row], ignore_index=True)
            
            with pd.ExcelWriter(FILE_NAME, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                 df_updated.to_excel(writer, sheet_name='商品資料', index=False)
            
            self.products_df = df_updated
            self.update_sales_prod_list() # 刷新銷售頁列表
            self.update_mgmt_prod_list()  # 刷新管理頁列表
            
            messagebox.showinfo("成功", f"已新增：{name}")
            self.var_add_name.set("")
            self.var_add_cost.set(0)
        except PermissionError: messagebox.showerror("錯誤", "Excel 未關閉！")

    # 4. 右側：提交更新
    def submit_update_product(self):
        name = self.var_upd_name.get() # 這是 Key，不能空的
        if not name:
            messagebox.showwarning("提示", "請先從列表選擇要編輯的商品")
            return
            
        new_tag = self.var_upd_tag.get().strip()
        new_cost = self.var_upd_cost.get()
        
        try:
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
            df_old = pd.read_excel(FILE_NAME, sheet_name='商品資料')
            
            # 找到該行索引
            idx = df_old[df_old['商品名稱'] == name].index
            if not idx.empty:
                df_old.loc[idx, '分類Tag'] = new_tag
                df_old.loc[idx, '預設成本'] = new_cost
                df_old.loc[idx, '最後更新時間'] = now_str
                
                with pd.ExcelWriter(FILE_NAME, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                     df_old.to_excel(writer, sheet_name='商品資料', index=False)
                
                self.products_df = df_old
                self.update_sales_prod_list() # 刷新所有相關列表
                self.update_mgmt_prod_list()
                self.var_upd_time.set(now_str) # 即時更新介面時間
                
                messagebox.showinfo("成功", f"已更新：{name}")
            else:
                messagebox.showerror("錯誤", "找不到原始資料，請重啟程式試試")
                
        except PermissionError: messagebox.showerror("錯誤", "Excel 未關閉！")

if __name__ == "__main__":
    root = tk.Tk()
    style = ttk.Style()
    # 【修改點 3】 使用 'vista' 主題 (Windows原生樣式) 以確保 Checkbutton 是打勾(✓)而不是叉(X)
    # 若在非 Windows 系統上可能會報錯，會自動退回預設
    try:
        style.theme_use('vista') 
    except:
        pass # 如果不支援 vista 主題就使用預設，預設通常也是打勾
    app = SalesApp(root)
    root.mainloop()