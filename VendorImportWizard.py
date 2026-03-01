import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from datetime import datetime
import sys
import os

try:
    from tksheet import Sheet
except ImportError:
    Sheet = None

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class VendorImportWizard(tk.Toplevel):
    def __init__(self, parent, save_callback):
        super().__init__(parent)
        self.title("廠商資料批次匯入精靈 (Excel版)")
        self.geometry("1200x850")
        self.save_callback = save_callback 
        self.import_raw_df = pd.DataFrame()

        try:
            self.iconbitmap(resource_path("main.ico"))
        except:
            pass
        
        # 廠商匯入必填欄位
        self.REQUIRED_FIELDS = ["廠商名稱"]
        
        self.grab_set()
        self.setup_ui()

    def setup_ui(self):
        header = ttk.Frame(self, padding=20)
        header.pack(fill="x")
        ttk.Label(header, text="Step 1: 開啟廠商 Excel 檔案", font=("微軟正黑體", 12, "bold")).pack(side="left")
        ttk.Button(header, text="📁 選擇檔案", command=self.load_file).pack(side="left", padx=10)
        self.lbl_path = ttk.Label(header, text="尚未選取檔案", foreground="gray")
        self.lbl_path.pack(side="left")

        paned = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        paned.pack(fill="both", expand=True, padx=20)

        # 左側預覽
        left_f = ttk.LabelFrame(paned, text="Step 2: 原始資料預覽", padding=5)
        paned.add(left_f, weight=3)
        if Sheet:
            self.sheet = Sheet(left_f, data=[[]], show_row_index=True)
            self.sheet.pack(fill="both", expand=True)
            self.sheet.enable_bindings()
        else:
            self.sheet = tk.Text(left_f, wrap="none")
            self.sheet.pack(fill="both", expand=True)

        # 右側匹配
        right_f = ttk.LabelFrame(paned, text="Step 3: 廠商欄位映射設定", padding=10)
        paned.add(right_f, weight=1)

        # 這裡對應您 SHEET_VENDORS 的所有欄位
        self.field_keys = [
            "廠商名稱", "通路", "統編", "聯絡人", "電話", "地址", "備註",
            "平均前置天數", "總到貨率", "總合格率", "綜合評等分數", "星等"
        ]
        self.vars = {k: tk.StringVar(value="(不匯入 / 留空)") for k in self.field_keys}

        container = ttk.Frame(right_f)
        container.pack(fill="both", expand=True)
        canvas = tk.Canvas(container, width=320)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        for label in self.field_keys:
            f = ttk.Frame(scroll_frame)
            f.pack(fill="x", pady=4)
            prefix = "⭐ " if label in self.REQUIRED_FIELDS else "  "
            ttk.Label(f, text=f"{prefix}{label}:", width=13).pack(side="left")
            cb = ttk.Combobox(f, textvariable=self.vars[label], state="readonly")
            cb.pack(side="left", fill="x", expand=True)
            setattr(self, f"cb_{label}", cb)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        ttk.Label(right_f, text="\n* 廠商名稱必須匹配，否則無法匯入", foreground="#d9534f", font=("", 9)).pack(anchor="w")

        footer = ttk.Frame(self, padding=20)
        footer.pack(fill="x")
        ttk.Button(footer, text="✅ 執行廠商資料匯入", command=self.execute_import, width=35).pack(side="right")
        ttk.Button(footer, text="❌ 取消", command=self.destroy).pack(side="right", padx=10)

    def load_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel 檔案", "*.xlsx *.xls")])
        if not path: return
        try:
            self.lbl_path.config(text=f"已載入: {os.path.basename(path)}", foreground="green")
            self.import_raw_df = pd.read_excel(path).fillna("")
            headers = self.import_raw_df.columns.tolist()
            
            if Sheet and isinstance(self.sheet, Sheet):
                self.sheet.set_sheet_data(self.import_raw_df.values.tolist())
                self.sheet.headers(headers)

            options = ["(不匯入 / 留空)"] + [f"列 {i}: {h}" for i, h in enumerate(headers)]
            for label in self.field_keys:
                cb = getattr(self, f"cb_{label}")
                cb['values'] = options
                cb.set("(不匯入 / 留空)")

                # --- 智慧自動匹配邏輯 ---
                for opt in options:
                    h_low = opt.lower()
                    if label in opt: cb.set(opt); break
                    if label == "廠商名稱" and ("商店" in h_low or "公司" in h_low or "店名" in h_low or "名稱" in h_low): cb.set(opt); break
                    if label == "通路" and ("來源" in h_low or "平台" in h_low): cb.set(opt); break
                    if label == "統編" and ("統一編號" in h_low or "tax" in h_low): cb.set(opt); break
                    if label == "聯絡人" and ("對口" in h_low or "負責人" in h_low): cb.set(opt); break
        except Exception as e:
            messagebox.showerror("錯誤", f"無法讀取檔案: {e}")

    def execute_import(self):
        if self.import_raw_df.empty: return
        mapping = {}
        for label, var in self.vars.items():
            val = var.get()
            if val != "(不匯入 / 留空)":
                mapping[label] = int(val.split(":")[0].replace("列 ", ""))

        if "廠商名稱" not in mapping:
            messagebox.showerror("錯誤", "您必須對應「廠商名稱」欄位！")
            return

        new_list = []
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

        for _, row in self.import_raw_df.iterrows():
            try:
                v_name = str(row.iloc[mapping["廠商名稱"]]).strip()
                if not v_name or v_name.lower() == "nan": continue

                def get_val(key, default=""):
                    if key in mapping:
                        v = row.iloc[mapping[key]]
                        return str(v).strip() if str(v).strip() != "" else default
                    return default

                def get_num(key, default=0):
                    if key in mapping:
                        val = pd.to_numeric(row.iloc[mapping[key]], errors='coerce')
                        return val if pd.notna(val) else default
                    return default

                # 建立廠商格式
                item = {
                    "廠商名稱": v_name,
                    "通路": get_val("通路", ""),
                    "統編": get_val("統編", ""),
                    "聯絡人": get_val("聯絡人", ""),
                    "電話": get_val("電話", ""),
                    "地址": get_val("地址", ""),
                    "備註": get_val("備註", ""),
                    "平均前置天數": get_num("平均前置天數", 0),
                    "總到貨率": get_val("總到貨率", "0%"),
                    "總合格率": get_val("總合格率", "0%"),
                    "綜合評等分數": get_num("綜合評等分數", 0),
                    "星等": get_num("星等", 5),
                    "最後更新": now_str
                }
                new_list.append(item)
            except: continue

        if not new_list:
            messagebox.showwarning("警告", "無有效資料可匯入")
            return

        if messagebox.askyesno("匯入確認", f"準備匯入 {len(new_list)} 筆廠商資料。是否繼續？"):
            if self.save_callback(new_list):
                messagebox.showinfo("成功", "廠商資料庫已更新")
                self.destroy()