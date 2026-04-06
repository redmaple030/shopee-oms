import math
import os
from datetime import datetime

import pandas as pd
import tkinter as tk
from tkinter import messagebox, ttk


class ProcurementManager:
    """
    採購需求分析與庫存修正管理器。
    專門處理 ROP (補貨點) 模型運算及實體庫存盤點修正邏輯。
    """

    @staticmethod
    def generate_report(app):
        """
        核心運算：根據銷售速率、前置時間與安全權重生成採購建議清單。
        """
        if not hasattr(app, "tree_procure"):
            return

        # 清空舊列表
        for i in app.tree_procure.get_children():
            app.tree_procure.delete(i)

        if not os.path.exists(app.FILE_NAME):
            return

        try:
            with pd.ExcelFile(app.FILE_NAME) as xls:
                df_sales = pd.read_excel(xls, sheet_name="銷售紀錄")
                df_prods = pd.read_excel(xls, sheet_name="商品資料")

            if df_prods.empty:
                return

            # 資料清洗：將空值填補並轉換為正確型別
            df_prods["目前庫存"] = pd.to_numeric(df_prods["目前庫存"], errors="coerce").fillna(0)
            df_prods["安全庫存"] = pd.to_numeric(df_prods["安全庫存"], errors="coerce").fillna(0)
            df_sales["數量"] = pd.to_numeric(df_sales["數量"], errors="coerce").fillna(0)
            df_sales["日期"] = pd.to_datetime(df_sales["日期"], errors="coerce")

            now = pd.Timestamp.now()
            # 建立商品第一筆成交日期地圖 (作為銷售速率的分母參考)
            first_sale_map = df_sales.groupby("商品名稱")["日期"].min().to_dict()
            qty_sum = df_sales.groupby("商品名稱")["數量"].sum()

            # 讀取介面配置參數
            v_threshold = app.var_filter_velocity.get()
            cover_days = app.var_days_to_cover.get()
            s_multiplier = app.var_safety_multiplier.get()

            for _, row in df_prods.iterrows():
                p_name = str(row["商品名稱"])
                curr_stock = float(row["目前庫存"])
                base_safety = float(row["安全庫存"])

                # A. 計算銷售速率 (Velocity)
                st_date = pd.to_datetime(row.get("初始上架時間"), errors="coerce")
                if pd.isna(st_date):
                    st_date = first_sale_map.get(p_name, now)
                
                days_diff = max((now - st_date).days, 1)
                velocity = float(qty_sum.get(p_name, 0)) / days_diff

                # 執行嚴格速率過濾 (優先級最高)
                if velocity < v_threshold:
                    continue

                # B. 計算動態補貨點 (Reorder Point, ROP)
                # 補貨點 = (日均銷量 * 備貨天數) + (安全庫存 * 加權係數)
                reorder_point = (velocity * cover_days) + (base_safety * s_multiplier)

                # C. 判定缺貨狀態與視覺標籤
                is_needed = False
                status = ""
                row_tag = ""

                if curr_stock < 0:
                    status = "⚠️ 帳面超賣"
                    row_tag = "urgent"
                    is_needed = True
                elif curr_stock == 0:
                    status = "🚫 缺貨中"
                    row_tag = "urgent"
                    is_needed = True
                elif curr_stock <= reorder_point:
                    status = "🔴 需補貨"
                    row_tag = "urgent"
                    is_needed = True
                elif curr_stock <= (base_safety * s_multiplier) and base_safety > 0:
                    status = "🟡 庫存偏低"
                    row_tag = "warning"
                    is_needed = True

                # D. 寫入清單
                if is_needed:
                    suggest_qty = math.ceil(max(reorder_point - curr_stock, 0))
                    app.tree_procure.insert(
                        "",
                        "end",
                        values=(
                            p_name,
                            int(curr_stock),
                            round(reorder_point, 1),
                            f"{round(velocity, 2)}件/日",
                            status,
                            int(suggest_qty),
                        ),
                        tags=(row_tag,),
                    )

        except Exception as e:
            print(f"Procurement analysis error: {e}")



    @staticmethod
    def open_stock_correction(app):
        """
        [RPA 模式] 彈出視窗執行即時庫存盤點修正。
        """
        selected = app.tree_procure.selection()
        if not selected:
            return

        item_id = selected[0]
        item_vals = app.tree_procure.item(item_id)["values"]
        p_name = str(item_vals[0])
        old_stock = str(item_vals[1])

        # 建立彈出視窗
        win = tk.Toplevel(app.root)
        win.title(f"庫存校準 - {p_name}")
        win.geometry("350x280")
        win.resizable(False, False)
        win.grab_set()  # 視窗鎖定

        container = ttk.Frame(win, padding=20)
        container.pack(fill="both", expand=True)

        ttk.Label(container, text="商品名稱:", font=("", 10, "bold")).pack(anchor="w")
        ttk.Label(container, text=p_name, foreground="blue", wraplength=300).pack(
            anchor="w", pady=(2, 10)
        )

        ttk.Label(container, text=f"目前系統帳面庫存: {old_stock}").pack(anchor="w")
        ttk.Label(container, text="修正為實體盤點數量:", foreground="#d9534f").pack(
            anchor="w", pady=(10, 0)
        )

        # 數字輸入框
        var_new_stock = tk.IntVar(value=int(old_stock))
        entry = ttk.Entry(container, textvariable=var_new_stock, font=("Arial", 11))
        entry.pack(fill="x", pady=5)
        entry.focus_set()
        entry.selection_range(0, tk.END)

        def perform_save(event=None):
            try:
                new_val = var_new_stock.get()
                now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

                # 讀取商品分頁進行局部更新
                df_prods = pd.read_excel(app.FILE_NAME, sheet_name="商品資料")
                mask = df_prods["商品名稱"].astype(str).str.strip() == p_name.strip()
                
                if not df_prods[mask].empty:
                    idx = df_prods[mask].index[0]
                    df_prods.at[idx, "目前庫存"] = new_val
                    df_prods.at[idx, "最後更新時間"] = now_str

                    # 呼叫主程式萬用存檔引擎
                    if app._universal_save({"商品資料": df_prods}):
                        messagebox.showinfo("成功", f"【{p_name}】\n庫存已校準為: {new_val}")
                        # 觸發主程式數據刷新
                        app.products_df = app.load_products()
                        ProcurementManager.generate_report(app)
                        win.destroy()
                else:
                    messagebox.showerror("錯誤", "資料庫中找不到該商品，請重新整理。")

            except Exception as ex:
                messagebox.showerror("儲存失敗", f"請輸入正確數字格式。\n錯誤訊息: {ex}")

        # 按鈕區
        btn_f = ttk.Frame(container)
        btn_f.pack(fill="x", pady=15)

        ttk.Button(
            btn_f, text="✅ 執行盤點修正", command=perform_save, style="Accent.TButton"
        ).pack(side="left", expand=True, fill="x", padx=(0, 5))
        
        ttk.Button(btn_f, text="取消", command=win.destroy).pack(
            side="left", expand=True, fill="x"
        )

        # 綁定鍵盤 Enter 鍵
        win.bind("<Return>", perform_save)