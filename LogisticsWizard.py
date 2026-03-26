import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from datetime import datetime

class LogisticsWizard(tk.Toplevel):
    """ 
    物流維護(獨立模組版)
    整合：批次更新、跨單偵測、單號保護、時間紀錄
    """
    def __init__(self, parent, app_instance):
        super().__init__(parent)
        self.app = app_instance # 存取主程式實例 (SalesApp)
        self.tree = self.app.tree_pur_track
        
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("提示", "請先選擇商品項目")
            self.destroy()
            return

        self.title("🚛 物流維護")
        self.geometry("500x750")
        self.grab_set()

        # 從主程式獲取檔案名稱與分頁常數
        self.FILE_NAME = self.app.FILE_NAME if hasattr(self.app, 'FILE_NAME') else 'sales_data.xlsx'
        self.SHEET_TRACK = "進貨追蹤"
        self.SHEET_HIST = "進貨紀錄"

        # 1. 蒐集資料
        self.batch_list, self.unique_order_ids = self._collect_data(selected_items)
        self.is_batch = len(self.batch_list) > 1
        self.is_mixed_orders = len(self.unique_order_ids) > 1

        # 2. 繪製 UI
        self._setup_ui()

    def _collect_data(self, selected_items):
        """ 蒐集並清理選中項目的數據 """
        try:
            df_full = pd.read_excel(self.FILE_NAME, sheet_name=self.SHEET_TRACK)
        except Exception:
            df_full = pd.DataFrame()

        batch_list = []
        unique_ids = set()

        for item_id in selected_items:
            item_data = self.tree.item(item_id)
            vals = item_data['values']
            df_idx = int(item_data['text'])
            p_id = str(vals[0]).replace("'", "").strip()
            unique_ids.add(p_id)

            raw_logi = str(df_full.at[df_idx, '物流追蹤']).strip() if '物流追蹤' in df_full.columns else ""
            raw_remark = str(df_full.at[df_idx, '備註']).strip() if '備註' in df_full.columns else ""

            def clean(s):
                if s.lower() in ["nan", "none", "nat", ""]:
                    return ""
                return s.lstrip("'")

            batch_list.append({
                "df_idx": df_idx, 
                "p_name": str(vals[2]).strip(),
                "pur_id": p_id,
                "qty": int(vals[3]),
                "status": str(vals[7]),
                "logi_id": clean(raw_logi),
                "remark": clean(raw_remark)
            })
        return batch_list, unique_ids

    def _setup_ui(self):
        """ 繪製視窗介面 """
        # 頂部提示
        header = ttk.Frame(self, padding=10)
        header.pack(fill="x")

        if self.is_batch:
            ttk.Label(header, text=f"⚠️ 複選模式：已選取 {len(self.batch_list)} 筆項目", 
                      foreground="red", font=("", 12, "bold")).pack()
            if self.is_mixed_orders:
                ttk.Label(header, text="❗ 偵測到跨單編輯", foreground="#FF8C00", font=("", 10, "bold")).pack()
        else:
            ttk.Label(header, text=f"商品：{self.batch_list[0]['p_name']}", font=("", 11, "bold")).pack(anchor="w")

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=10)

        body = ttk.Frame(self, padding=20)
        body.pack(fill="both", expand=True)

        # 1. 數量與瑕疵 (僅單選)
        if not self.is_batch:
            q_f = ttk.LabelFrame(body, text="📦 數量與驗收", padding=10)
            q_f.pack(fill="x", pady=(0, 10))
            self.var_qty = tk.IntVar(value=self.batch_list[0]['qty'])
            self.var_defects = tk.IntVar(value=0)
            ttk.Label(q_f, text="實際收到數量:").pack(anchor="w")
            ttk.Entry(q_f, textvariable=self.var_qty).pack(fill="x")
            ttk.Label(q_f, text="瑕疵損壞數量:").pack(anchor="w", pady=(5,0))
            ttk.Entry(q_f, textvariable=self.var_defects).pack(fill="x")

        # 2. 物流更新區
        l_f = ttk.LabelFrame(body, text="🚚 物流狀態更新", padding=10)
        l_f.pack(fill="x", pady=10)
        
        ttk.Label(l_f, text="變更階段為:").pack(anchor="w")
        self.var_status = tk.StringVar(value=self.batch_list[0]['status'])
        cb = ttk.Combobox(l_f, textvariable=self.var_status, state="readonly")
        cb['values'] = ("待出貨", "廠商已發貨", "貨到集運倉", "集運倉已發貨", "抵達台灣海關", "國內配送中")
        cb.pack(fill="x", pady=5)

        # 單號保護
        self.var_skip_logi = tk.BooleanVar(value=False)
        self.chk_skip = ttk.Checkbutton(l_f, text="固定原有單號 (不更動)", 
                                        variable=self.var_skip_logi, command=self._toggle_ent)
        self.chk_skip.pack(anchor="w", pady=5)

        ttk.Label(l_f, text="物流單號:").pack(anchor="w")
        self.var_logi = tk.StringVar(value=self.batch_list[0]['logi_id'])
        self.ent_logi = ttk.Entry(l_f, textvariable=self.var_logi)
        self.ent_logi.pack(fill="x", pady=5)

        # 3. 備註
        r_f = ttk.LabelFrame(body, text="📝 備註事項", padding=10)
        r_f.pack(fill="x", pady=10)
        self.var_remark = tk.StringVar(value=self.batch_list[0]['remark'])
        ttk.Entry(r_f, textvariable=self.var_remark).pack(fill="x", pady=5)

        ttk.Button(body, text="🚀 執行更新並存檔", command=self.save, style="Accent.TButton").pack(pady=20, fill="x")

    def _toggle_ent(self):
        state = "disabled" if self.var_skip_logi.get() else "normal"
        self.ent_logi.config(state=state)

    def save(self):
        """ 執行存檔邏輯 """
        if self.is_mixed_orders and not self.var_skip_logi.get():
            if not messagebox.askyesno("二次確認", "您正在跨單編輯且未開啟單號保護，確定要覆蓋單號嗎？"):
                return

        try:
            today = datetime.now().strftime("%Y-%m-%d")
            # 透過 app_instance 讀取資料
            with pd.ExcelFile(self.FILE_NAME) as xls:
                df_track = pd.read_excel(xls, sheet_name=self.SHEET_TRACK)
                df_hist = pd.read_excel(xls, sheet_name=self.SHEET_HIST)

            # 數據清洗
            text_cols = ['物流狀態', '物流追蹤', '備註', '進貨單號', '商品名稱', '時間_廠商出貨', 
                         '時間_抵達集運倉', '時間_集運倉出貨', '時間_抵達台灣海關', '時間_國內配送中']
            for df in [df_track, df_hist]:
                for col in text_cols:
                    if col not in df.columns:
                        df[col] = ""
                df[text_cols] = df[text_cols].fillna("").astype(str).replace(['nan', 'NaN', 'None'], '')

            for item in self.batch_list:
                t_idx = item['df_idx']
                df_track.at[t_idx, '物流狀態'] = self.var_status.get()
                df_track.at[t_idx, '備註'] = self.var_remark.get().strip()

                if not self.var_skip_logi.get() and self.var_logi.get().strip():
                    df_track.at[t_idx, '物流追蹤'] = f"'{self.var_logi.get().strip()}"

                if not self.is_batch:
                    df_track.at[t_idx, '數量'] = self.var_qty.get()
                    u_price = pd.to_numeric(df_track.at[t_idx, '進貨單價'], errors='coerce')
                    df_track.at[t_idx, '進貨總額'] = self.var_qty.get() * u_price

                time_map = {"廠商已發貨": "時間_廠商出貨", "貨到集運倉": "時間_抵達集運倉", 
                            "集運倉已發貨": "時間_集運倉出貨", "抵達台灣海關": "時間_抵達台灣海關", "國內配送中": "時間_國內配送中"}
                col = time_map.get(self.var_status.get())
                if col: 
                    df_track.at[t_idx, col] = today

                # 同步歷史表
                h_mask = (df_hist['進貨單號'].str.replace("'", "").str.strip() == item['pur_id']) & \
                         (df_hist['商品名稱'].str.strip() == item['p_name'])
                if not df_hist[h_mask].empty:
                    df_hist.loc[h_mask, '物流狀態'] = self.var_status.get()
                    df_hist.loc[h_mask, '備註'] = self.var_remark.get().strip()
                    if col: 
                        df_hist.loc[h_mask, col] = today
                    if not self.var_skip_logi.get() and self.var_logi.get().strip():
                        df_hist.loc[h_mask, '物流追蹤'] = df_track.at[t_idx, '物流追蹤']

            # 呼叫主程式的萬用引擎存檔
            if self.app._universal_save({self.SHEET_TRACK: df_track, self.SHEET_HIST: df_hist}):
                messagebox.showinfo("成功", "物流資訊同步更新成功")
                self.app.load_purchase_tracking()
                self.destroy()
        except Exception as e:
            messagebox.showerror("錯誤", f"存檔失敗: {e}")