import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from decimal import Decimal

class ShippingDistributor(tk.Toplevel):
    def __init__(self, parent, app_instance):
        super().__init__(parent)
        self.app = app_instance
        
        # 1. 初始化視窗基本設定
        self.title("⚖️ 整單費用自動分攤")
        self.geometry("450x400")
        self.resizable(False, False)
        
        # 讓視窗出現在螢幕中央
        self.transient(parent) 
        self.grab_set()

        # 2. 取得選中的單號邏輯
        try:
            sel = self.app.tree_pur_track.selection()
            if not sel:
                messagebox.showwarning("提示", "請先選擇清單中的一個商品項目")
                self.destroy()
                return
                
            item = self.app.tree_pur_track.item(sel[0])
            # 確保抓到的是『進貨單號』
            self.target_pur_id = str(item['values'][0]).replace("'", "").strip()
            
            # 從主程式抓取正確的分頁名稱常數
            self.FILE_NAME = getattr(self.app, 'FILE_NAME', 'sales_data.xlsx')
            self.SHEET_TRACK = getattr(self.app, 'SHEET_PUR_TRACKING', '進貨追蹤')
            self.SHEET_HIST = getattr(self.app, 'SHEET_PURCHASES', '進貨紀錄')
            self.SHEET_PROD = getattr(self.app, 'SHEET_PRODUCTS', '商品資料')

            # 3. 執行 UI 繪製
            self._setup_ui()
            
        except Exception as e:
            messagebox.showerror("初始化失敗", f"無法開啟分攤視窗: {e}")
            self.destroy()

    def _setup_ui(self):
        """ 繪製輸入介面，確保使用 pack 填滿 """
        # 主容器
        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(fill="both", expand=True)

        # 顯示單號
        header_f = ttk.Frame(main_frame)
        header_f.pack(fill="x", pady=(0, 15))
        ttk.Label(header_f, text="正在處理單號：", font=("", 10)).pack(side="left")
        ttk.Label(header_f, text=self.target_pur_id, font=("Arial", 11, "bold"), foreground="blue").pack(side="left")

        # 輸入區域 (用 LabelFrame 包起來更清楚)
        input_box = ttk.LabelFrame(main_frame, text="請輸入整箱貨物總額", padding=15)
        input_box.pack(fill="x", pady=5)

        # 總運費
        row1 = ttk.Frame(input_box)
        row1.pack(fill="x", pady=5)
        ttk.Label(row1, text="總運費 ($):", width=12).pack(side="left")
        self.var_total_ship = tk.DoubleVar(value=0.0)
        ent_ship = ttk.Entry(row1, textvariable=self.var_total_ship)
        ent_ship.pack(side="left", fill="x", expand=True)
        ent_ship.focus_set() # 自動聚焦到運費

        # 總稅金
        row2 = ttk.Frame(input_box)
        row2.pack(fill="x", pady=5)
        ttk.Label(row2, text="總稅金 ($):", width=12).pack(side="left")
        self.var_total_tax = tk.DoubleVar(value=0.0)
        ttk.Entry(row2, textvariable=self.var_total_tax).pack(side="left", fill="x", expand=True)

        # 說明
        ttk.Label(main_frame, text="* 系統將讀取商品資料庫中的「單位權重」自動分攤", 
                  foreground="gray", font=("微軟正黑體", 9)).pack(pady=10)

        # 按鈕區
        btn_f = ttk.Frame(main_frame)
        btn_f.pack(fill="x", side="bottom", pady=10)
        
        ttk.Button(btn_f, text="✅ 開始自動分攤並存檔", 
                   command=self.calculate_and_save, style="Accent.TButton").pack(fill="x")
        
        ttk.Button(btn_f, text="取消", command=self.destroy).pack(fill="x", pady=5)

    def calculate_and_save(self):
        """ 執行權重分攤計算法 """
        try:
            # 使用主程式的 dec_round 工具確保精度
            t_ship = Decimal(str(self.var_total_ship.get()))
            t_tax = Decimal(str(self.var_total_tax.get()))
            
            with pd.ExcelFile(self.FILE_NAME) as xls:
                df_track = pd.read_excel(xls, sheet_name=self.SHEET_TRACK)
                df_hist = pd.read_excel(xls, sheet_name=self.SHEET_HIST)
                df_prods = pd.read_excel(xls, sheet_name=self.SHEET_PROD)

            # 確保權重欄位存在
            df_prods['單位權重'] = pd.to_numeric(df_prods.get('單位權重', 1.0), errors='coerce').fillna(1.0)
            weight_map = df_prods.set_index('商品名稱')['單位權重'].to_dict()

            # 篩選
            df_track['tmp_id'] = df_track['進貨單號'].astype(str).str.replace("'", "").str.strip()
            mask = df_track['tmp_id'] == self.target_pur_id
            
            if df_track[mask].empty:
                messagebox.showerror("錯誤", "找不到該訂單項目，請重新整理列表")
                return

            # 計算總重
            total_weight = Decimal("0.0")
            for _, row in df_track[mask].iterrows():
                q = Decimal(str(row.get('數量', 1)))
                w = Decimal(str(weight_map.get(str(row['商品名稱']).strip(), 1.0)))
                total_weight += (q * w)

            if total_weight <= 0: 
                total_weight = Decimal("1.0")

            # 分攤循環
            for idx in df_track[mask].index:
                p_name = str(df_track.at[idx, '商品名稱']).strip()
                q = Decimal(str(df_track.at[idx, '數量']))
                w = Decimal(str(weight_map.get(p_name, 1.0)))
                
                ratio = (q * w) / total_weight
                # 呼叫主程式的高精度四捨五入
                alloc_s = self.app.dec_round(t_ship * ratio)
                alloc_t = self.app.dec_round(t_tax * ratio)

                df_track.at[idx, '分攤運費'] = float(alloc_s)
                df_track.at[idx, '海關稅金'] = float(alloc_t)
                
                # 同步歷史表
                df_hist['tmp_id'] = df_hist['進貨單號'].astype(str).str.replace("'", "").str.strip()
                h_mask = (df_hist['tmp_id'] == self.target_pur_id) & (df_hist['商品名稱'].astype(str).str.strip() == p_name)
                if not df_hist[h_mask].empty:
                    df_hist.loc[h_mask, '分攤運費'] = float(alloc_s)
                    df_hist.loc[h_mask, '海關稅金'] = float(alloc_t)

            # 存檔
            df_track.drop(columns=['tmp_id'], inplace=True, errors='ignore')
            df_hist.drop(columns=['tmp_id'], inplace=True, errors='ignore')
            
            if self.app._universal_save({self.SHEET_TRACK: df_track, self.SHEET_HIST: df_hist}):
                messagebox.showinfo("成功", "運費與稅金分攤完成！")
                self.app.load_purchase_tracking()
                self.destroy()

        except Exception as e:
            messagebox.showerror("計算失敗", f"請檢查輸入數值是否正確: {e}")