#出貨單生成插件

import os
import sys
import webbrowser
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk
import shutil

from main import resource_path # 用於檔案操作

# --- 設定區 ---
TEMP_FOLDER = "temp_print_files" # 存放出貨單的資料夾
EXPIRE_DAYS = 3                  # 檔案保留天數，超過則自動刪除


def resource_path(relative_path):
    """ 獲取資源的絕對路徑 (打包用) """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)




def manage_temp_folder():
    """ 檢查資料夾是否存在，並清理舊檔案 """
    # 1. 如果資料夾不存在，就建立它
    if not os.path.exists(TEMP_FOLDER):
        os.makedirs(TEMP_FOLDER)
        return

    # 2. 清理過期檔案
    now = datetime.now()
    try:
        for filename in os.listdir(TEMP_FOLDER):
            file_path = os.path.join(TEMP_FOLDER, filename)
            
            # 只處理檔案 (排除資料夾)
            if os.path.isfile(file_path):
                # 取得檔案最後修改時間
                file_time = datetime.fromtimestamp(os.path.getmtime(file_path))
                
                # 如果檔案時間早於 (現在 - EXPIRE_DAYS)，則刪除
                if now - file_time > timedelta(days=EXPIRE_DAYS):
                    os.remove(file_path)
                    print(f"系統清理舊出貨單: {filename}")
    except Exception as e:
        print(f"清理暫存資料夾出錯: {e}")

def show_shipping_dialog(parent, order_info, items_data):
    """ (其餘 Dialog 部分維持不變) """
    dialog = tk.Toplevel(parent)
    dialog.title("選擇出貨單尺寸")
    dialog.geometry("320x280")
    dialog.grab_set()

    ttk.Label(dialog, text="請選擇列印尺寸：", font=("", 11, "bold")).pack(pady=15)

    size_var = tk.StringVar(value="A4")
    ttk.Radiobutton(dialog, text="A4 標準報表 (每頁12項)", variable=size_var, value="A4").pack(pady=5, anchor="w", padx=60)
    ttk.Radiobutton(dialog, text="10x15cm 標籤 (每頁8項)", variable=size_var, value="Label").pack(pady=5, anchor="w", padx=60)

    try:
        dialog.iconbitmap(resource_path("main.ico"))
    except:
        pass
        

    def on_confirm():
        generate_shipping_html(order_info, items_data, size_var.get())
        dialog.destroy()

    ttk.Button(dialog, text="確認並預覽", command=on_confirm).pack(pady=20)

def generate_shipping_html(info, items, size_choice):
    """ 核心 HTML 產生邏輯 """
    
    # --- [新增] 在產生新檔前，先執行管理與清理 ---
    manage_temp_folder()

    shop_name = info.get('shop_name', '商家出貨單') # 取得店名
    order_id = datetime.now().strftime("%Y%m%d%H%M%S")
    
    # (中間計算邏輯與 HTML 模板部分不變，字體維持 10px，編號部分維持原設定)
    # ... (省略中間重複的 HTML 內容) ...

    # 計算金額與天數
    product_total = sum(i['total_sales'] for i in items)
    ship_fee = info.get('ship_fee', 0)
    payer = info.get('payer', "買家付")
    discount_amount = info.get('discount_amount', 0)
    discount_tag = info.get('discount_tag', "折扣")

    if payer == "買家付":
        final_paid = product_total + ship_fee - discount_amount
        display_ship = f"${ship_fee:,.0f}"
    else:
        final_paid = product_total - discount_amount
        display_ship = "免運 (賣家付)"

    font_px = "10px"
    if size_choice == "A4":
        limit, page_height, page_width, row_height = 12, "297mm", "210mm", "1.2cm"
    else:
        limit, page_height, page_width, row_height = 8, "150mm", "100mm", "1.0cm"

    chunks = [items[i:i + limit] for i in range(0, len(items), limit)]
    total_pages = len(chunks)
    all_pages_html = ""


    for page_idx, chunk in enumerate(chunks, 1):
        is_last_page = (page_idx == total_pages)
        table_rows = ""
        for item in chunk:
            sku = item.get('sku', '--')
            if not sku or str(sku).strip() == "": sku = "--"
            table_rows += f"""<tr style="height: {row_height};">
                <td style="text-align:center;">{sku}</td>
                <td style="padding-left: 8px;">{item['name']}</td>
                <td style="text-align:center;">{item['qty']}</td>
                <td style="text-align:right; padding-right: 8px;">${item['unit_price']:,.0f}</td>
            </tr>"""

        summary_section = ""
        if is_last_page:
            summary_section = f"""<div class="summary-box">
                <div class="sum-row">
                    <div class="sum-item"><span class="label">物流運費 ({payer})</span><span class="val">{display_ship}</span></div>
                    <div class="sum-item"><span class="label">{discount_tag}</span><span class="val">-${discount_amount:,.0f}</span></div>
                </div>
                <div class="sum-row" style="border-bottom:none; background:#eee !important;">
                    <div class="sum-item"><span class="label">商品總額</span><span class="val">${product_total:,.0f}</span></div>
                    <div class="sum-item" style="background:#ddd !important;">
                        <span class="label">買家應付總額</span><span class="val" style="font-size: 14px;">${final_paid:,.0f}</span>
                    </div>
                </div>
            </div>"""

        all_pages_html += f"""
        <div class="page">
            <div class="header">
                <!-- 顯示自定義店名 -->
                <div style="font-size: 14px; font-weight: bold; text-align: center;">{shop_name}</div>
                <h3 style="margin:2px 0; font-size: 11px; text-align: center; color: #555;">出貨明細 ({size_choice})</h3>
                <div style="text-align:right; font-size: 9px;">頁次：{page_idx} / {total_pages}</div>
            </div>

            <table class="info">
                <tr><td><b>買家：</b>{info['buyer']}</td><td style="text-align:right;">{info['date']}</td></tr>
                <tr><td><b>物流：</b>{info['ship_method']}</td><td style="text-align:right;">ID: {order_id}</td></tr>
            </table>
            <table class="item-table">
                <thead><tr><th width="20%">商品編號</th><th width="50%">商品名稱</th><th width="10%">數量</th><th width="20%">單價</th></tr></thead>
                <tbody>{table_rows}</tbody>
            </table>
            {summary_section}
            <div class="page-footer">-- 感謝您的購買！請錄影拆封保障權益 --</div>
        </div>"""

    # --- [關鍵修正：路徑結合] ---
    filename = f"Shipping_Note_{order_id}.html"
    file_full_path = os.path.join(TEMP_FOLDER, filename) # 存入 temp_print_files 資料夾

    html_content = f"""<!DOCTYPE html><html>
    <head><meta charset="UTF-8"><style>
        * {{ box-sizing: border-box; -webkit-print-color-adjust: exact; }}
        @page {{ size: {size_choice == 'A4' and 'A4' or '100mm 150mm'}; margin: 0; }}
        body {{ font-family: "微軟正黑體", sans-serif; margin: 0; padding: 0; background: #f0f0f0; font-size: {font_px}; }}
        .page {{ width: {page_width}; height: {page_height}; padding: {size_choice == 'A4' and '15mm' or '8mm'}; 
                background: white; margin: 10px auto; display: flex; flex-direction: column; 
                page-break-after: always; overflow: hidden; }}
        @media print {{ body {{ background: none; }} .page {{ margin: 0; border: none; }} }}
        .header {{ border-bottom: 1.5px solid #000; padding-bottom: 3px; margin-bottom: 8px; }}
        .info {{ width: 100%; margin-bottom: 8px; font-size: {font_px}; }}
        .item-table {{ width: 100%; border-collapse: collapse; table-layout: fixed; }}
        .item-table th, .item-table td {{ border: 1px solid #000; font-size: {font_px}; padding: 4px 2px; }}
        .item-table th {{ background-color: #eee !important; }}
        .summary-box {{ border: 1.5px solid #000; margin-top: 5px; }}
        .sum-row {{ display: flex; border-bottom: 1px solid #000; }}
        .sum-item {{ flex: 1; padding: 4px; border-right: 1px solid #000; }}
        .sum-item:last-child {{ border-right: none; }}
        .label {{ font-size: 9px; display: block; font-weight: bold; }}
        .val {{ font-size: {font_px}; font-weight: bold; }}
        .page-footer {{ margin-top: auto; text-align: center; font-size: 9px; padding-top: 5px; }}
    </style></head>
    <body onload="window.print()">{all_pages_html}</body></html>"""

    # 寫入指定的暫存資料夾
    with open(file_full_path, "w", encoding="utf-8") as f:
        f.write(html_content)
    
    # 開啟該路徑
    webbrowser.open(os.path.abspath(file_full_path))
