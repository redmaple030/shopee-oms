from tkinter import messagebox
import pandas as pd
from decimal import Decimal

class RecallManager:
    """ 
    工作流回溯管理器 V1.0
    專門處理將『已送出但未結案』的數據抓回輸入頁面，並補償相關庫存或狀態。
    """

    @staticmethod
    def recall_sales_order(app):
        """ 將訂單追蹤區的資料退回至『銷售輸入』 """
        sel = app.tree_track.selection()
        if not sel:
            messagebox.showwarning("提示", "請先選擇要退回修改的訂單項目")
            return
            
        item = app.tree_track.item(sel[0])
        order_id = str(item['values'][0]).replace("'", "").strip()

        # 1. 檢查目標頁面是否有資料
        if app.cart_data:
            if not messagebox.askyesno("警告", "銷售輸入頁面已有未送出的資料，抓回將會覆蓋，是否繼續？"):
                return

        try:
            # 2. 讀取所需資料
            with pd.ExcelFile(app.FILE_NAME) as xls:
                df_track = pd.read_excel(xls, sheet_name='訂單追蹤')
                df_prods = pd.read_excel(xls, sheet_name='商品資料')
            
            # 格式化編號
            df_track['訂單編號'] = df_track['訂單編號'].astype(str).str.replace(r'^\'', '', regex=True).str.replace(r'\.0$', '', regex=True).str.strip()
            
            mask = df_track['訂單編號'] == order_id
            target_rows = df_track[mask].copy()
            if target_rows.empty: 
                return

            # --- 3. 庫存補償邏輯 ---
            restored_log = []
            for _, row in target_rows.iterrows():
                p_name = row['商品名稱']
                qty = int(row['數量'])
                p_idx = df_prods[df_prods['商品名稱'] == p_name].index
                if not p_idx.empty:
                    df_prods.at[p_idx[0], '目前庫存'] += qty
                    restored_log.append(f"{p_name}(+{qty})")

            # --- 4. 重新填充主程式購物車與介面 ---
            app.cart_data = []
            for i in app.tree.get_children(): 
                app.tree.delete(i)

            for _, row in target_rows.iterrows():
                s_price = Decimal(str(row.get('單價(售)', 0)))
                qty = Decimal(str(row.get('數量', 0)))
                c_price = Decimal(str(row.get('單價(進)', 0)))
                sku = str(row.get('商品編號', '')).replace("'", "")

                app.cart_data.append({
                    "sku": sku, "name": row['商品名稱'], "qty": int(qty),
                    "unit_cost": float(c_price), "unit_price": float(s_price),
                    "total_sales": float(s_price * qty), "total_cost": float(c_price * qty)
                })
                app.tree.insert("", "end", values=(sku if sku else "--", row['商品名稱'], int(qty), float(s_price), float(s_price * qty)))

            # --- 5. 帶回買家與平台資訊 ---
            header_info = app._get_full_order_info(df_track, order_id)
            app.var_enable_cust.set(True)
            app.toggle_cust_info() 
            app.var_cust_name.set(header_info.get('買家名稱', ''))
            app.var_cust_loc.set(header_info.get('取貨地點', ''))
            app.var_ship_method.set(header_info.get('寄送方式', ''))
            app.var_platform.set(header_info.get('交易平台', '蝦皮購物'))

            # --- 6. 存檔並移除追蹤紀錄 ---
            df_track_new = df_track[~mask]
            if app._universal_save({'訂單追蹤': df_track_new, '商品資料': df_prods}):
                app.products_df = df_prods
                app.update_sales_prod_list()
                app.load_tracking_data()

                # 跳轉分頁 (假設銷售輸入在 index 3)
                app.root.nametowidget(app.root.winfo_children()[0]).select(3) 
                app.update_totals()
                messagebox.showinfo("成功", f"訂單 {order_id} 已退回修改並重填回庫存。")

        except Exception as e:
            messagebox.showerror("撤回失敗", f"錯誤: {e}")

    @staticmethod
    def recall_purchase_order(app):
        """ 將進貨追蹤區的資料退回至『進貨管理』 """
        sel = app.tree_pur_track.selection()
        if not sel:
            messagebox.showwarning("提示", "請先選擇要退回的進貨單")
            return
            
        item = app.tree_pur_track.item(sel[0])
        pur_id = str(item['values'][0]).replace("'", "").strip()

        if app.pur_cart_data:
            if not messagebox.askyesno("警告", "採購清單已有資料，抓回將會覆蓋，是否繼續？"):
                return

        try:
            with pd.ExcelFile(app.FILE_NAME) as xls:
                df_pt = pd.read_excel(xls, sheet_name='進貨追蹤')
                df_hist = pd.read_excel(xls, sheet_name='進貨紀錄')

            df_pt['進貨單號'] = df_pt['進貨單號'].astype(str).str.replace("'", "").str.strip()
            mask = df_pt['進貨單號'] == pur_id
            target_rows = df_pt[mask].copy()

            if target_rows.empty: 
                return

            # --- 重新填充採購清單 ---
            app.pur_cart_data = []
            for i in app.tree_pur_cart.get_children(): 
                app.tree_pur_cart.delete(i)

            for _, row in target_rows.iterrows():
                cost = float(row['進貨單價'])
                qty = int(row['數量'])
                tax = float(row.get('進項稅額', 0))
                
                app.pur_cart_data.append({
                    "name": row['商品名稱'], "qty": qty, "cost": cost, "tax": tax, "total": cost * qty
                })
                app.tree_pur_cart.insert("", "end", values=(row['商品名稱'], qty, cost, tax, cost * qty))

            app.var_pur_supplier.set(target_rows.iloc[0]['供應商'])
            
            # --- 移除追蹤與歷史紀錄 (因為還沒入庫，這算撤銷下單) ---
            df_pt_new = df_pt[~mask]
            df_hist['tmp_id'] = df_hist['進貨單號'].astype(str).str.replace("'", "").str.strip()
            df_hist_new = df_hist[df_hist['tmp_id'] != pur_id].drop(columns=['tmp_id'])

            if app._universal_save({'進貨追蹤': df_pt_new, '進貨紀錄': df_hist_new}):
                app.root.nametowidget(app.root.winfo_children()[0]).select(0) # 跳轉到進貨
                app.update_pur_cart_total()
                app.load_purchase_tracking()
                messagebox.showinfo("成功", f"進貨單 {pur_id} 已退回編輯頁面。")

        except Exception as e:
            messagebox.showerror("撤回失敗", f"錯誤: {e}")