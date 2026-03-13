import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from datetime import datetime
import sys
import os

# 嘗試匯入專業表格套件
try:
    from tksheet import Sheet
except ImportError:
    Sheet = None


def resource_path(relative_path):
    """ 獲取資源的絕對路徑 (打包用) """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


class ImportWizard(tk.Toplevel):
    def __init__(self, parent, save_callback):
        super().__init__(parent)
        self.title("商品資料批次匯入精靈(excel表格版)")
        self.geometry("1200x850")
        self.save_callback = save_callback 
        self.import_raw_df = pd.DataFrame()

        try:
            self.iconbitmap(resource_path("main.ico"))
        except:
            pass

        
        
        # ERP 核心必填欄位
        self.REQUIRED_FIELDS = ["商品名稱", "目前庫存", "預設成本"]
        
        self.grab_set()
        self.setup_ui()

    def setup_ui(self):
        # 頂部：檔案選取區
        header = ttk.Frame(self, padding=20)
        header.pack(fill="x")
        ttk.Label(header, text="Step 1: 開啟舊有的商品 Excel", font=("微軟正黑體", 12, "bold")).pack(side="left")
        ttk.Button(header, text="📁 選擇檔案", command=self.load_file).pack(side="left", padx=10)
        self.lbl_path = ttk.Label(header, text="尚未選取檔案", foreground="gray")
        self.lbl_path.pack(side="left")

        # 中間：雙欄佈局 (左表格預覽，右映射設定)
        paned = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        paned.pack(fill="both", expand=True, padx=20)

        # --- 左側：回歸本來的表格 UI ---
        left_f = ttk.LabelFrame(paned, text="Step 2: 原始資料預覽 (tksheet)", padding=5)
        paned.add(left_f, weight=3)
        
        if Sheet:
            self.sheet = Sheet(left_f, data=[[]], show_row_index=True)
            self.sheet.pack(fill="both", expand=True)
            self.sheet.enable_bindings()
        else:
            # 備援方案：若使用者未安裝 tksheet，才顯示文字
            self.sheet = tk.Text(left_f, wrap="none")
            self.sheet.pack(fill="both", expand=True)
            ttk.Label(left_f, text="建議安裝 tksheet 以獲得最佳表格體驗", foreground="red").pack()

        # --- 右側：欄位映射區 ---
        right_f = ttk.LabelFrame(paned, text="Step 3: ERP 欄位匹配設定", padding=10)
        paned.add(right_f, weight=1)

        self.field_keys = [
            "商品名稱", "商品編號", "分類Tag", "單位權重", 
            "目前庫存", "預設成本", "安全庫存", 
            "初始上架時間", "最後進貨時間", "商品連結", "商品備註"
        ]
        self.vars = {k: tk.StringVar(value="(不匯入 / 留空)") for k in self.field_keys}

        # 映射清單加入滾輪，防止欄位過多塞不下
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
            
            # 每一個對應欄位
            cb = ttk.Combobox(f, textvariable=self.vars[label], state="readonly")
            cb.pack(side="left", fill="x", expand=True)
            setattr(self, f"cb_{label}", cb)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        ttk.Label(right_f, text="\n* 每個欄位均可選擇「不匯入」", foreground="#d9534f", font=("", 9)).pack(anchor="w")

        # 底部：按鈕區
        footer = ttk.Frame(self, padding=20)
        footer.pack(fill="x")
        ttk.Button(footer, text="✅ 開始執行資料核對與匯入", command=self.execute_import, width=35).pack(side="right")
        ttk.Button(footer, text="❌ 取消", command=self.destroy).pack(side="right", padx=10)

    def load_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel 活頁簿", "*.xlsx"), ("舊版 Excel", "*.xls")])
        if not path: 
            return
        try:
            self.lbl_path.config(text=f"已載入: {path.split('/')[-1]}", foreground="green")
            self.import_raw_df = pd.read_excel(path).fillna("")
            headers = self.import_raw_df.columns.tolist()
            
            # 更新 Step 2 的表格資料
            if Sheet and isinstance(self.sheet, Sheet):
                self.sheet.set_sheet_data(self.import_raw_df.values.tolist())
                self.sheet.headers(headers)
            else:
                self.sheet.delete("1.0", tk.END)
                self.sheet.insert(tk.END, self.import_raw_df.to_string())

            # 更新 Step 3 的選單選項
            options = ["(不匯入 / 留空)"] + [f"列 {i}: {h}" for i, h in enumerate(headers)]
            
            for label in self.field_keys:
                cb = getattr(self, f"cb_{label}")
                cb['values'] = options
                cb.set("(不匯入 / 留空)") # 預設重設，防止舊緩存

                # --- 智慧自動匹配邏輯 ---
                for opt in options:
                    h_low = opt.lower()
                    if label in opt: cb.set(opt); break
                    if label == "商品編號" and ("編號" in h_low or "sku" in h_low or "位置" in h_low): cb.set(opt); break
                    if label == "單位權重" and ("g" in h_low or "重量" in h_low or "weight" in h_low): cb.set(opt); break
                    if label == "分類Tag" and ("分類" in h_low or "標籤" in h_low or "tag" in h_low): cb.set(opt); break

        except Exception as e:
            messagebox.showerror("讀取失敗", f"Excel 解析錯誤: {e}")

    def execute_import(self):
        if self.import_raw_df.empty: return

        # 1. 整理匹配對應表
        mapping = {}
        for label, var in self.vars.items():
            val = var.get()
            if val != "(不匯入 / 留空)":
                mapping[label] = int(val.split(":")[0].replace("列 ", ""))

        # 2. 核心欄位檢查
        missing = [f for f in self.REQUIRED_FIELDS if f not in mapping]
        if missing:
            messagebox.showerror("映射不全", f"您漏掉了 ERP 核心必填欄位：\n{', '.join(missing)}")
            return

        # 3. 逐行資料清洗與轉換
        new_list = []
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

        for _, row in self.import_raw_df.iterrows():
            try:
                p_name = str(row.iloc[mapping["商品名稱"]]).strip()
                if not p_name or p_name.lower() == "nan": continue

                def get_val(key, default):
                    if key in mapping:
                        v = row.iloc[mapping[key]]
                        return str(v).strip() if str(v).strip() != "" else default
                    return default

                def get_num(key, default, is_float=False):
                    if key in mapping:
                        raw_v = row.iloc[mapping[key]]
                        val = pd.to_numeric(raw_v, errors='coerce')
                        if pd.isna(val): return default
                        return float(val) if is_float else int(val)
                    return default

                item = {
                    "商品編號": get_val("商品編號", ""),
                    "分類Tag": get_val("分類Tag", "未分類"),
                    "商品名稱": p_name,
                    "預設成本": get_num("預設成本", 0.0, True),
                    "目前庫存": get_num("目前庫存", 0),
                    "最後更新時間": now_str,
                    "初始上架時間": get_val("初始上架時間", now_str),
                    "最後進貨時間": get_val("最後進貨時間", ""),
                    "安全庫存": get_num("安全庫存", 0),
                    "商品連結": get_val("商品連結", "無"),
                    "商品備註": get_val("商品備註", "無"),
                    "單位權重": get_num("單位權重", 1.0, True)
                }
                new_list.append(item)
            except: continue

        if not new_list:
            messagebox.showwarning("警告", "掃描後無有效商品可匯入。")
            return

        # 4. 最終發射
        if messagebox.askyesno("匯入確認", f"已完成資料校準，準備匯入 {len(new_list)} 筆商品。\n確定執行嗎？"):
            if self.save_callback(new_list):
                messagebox.showinfo("成功", "商品資料庫已完成增量更新。")
                self.destroy()
                
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from datetime import datetime
import sys
import os

# 嘗試匯入專業表格套件
try:
    from tksheet import Sheet
except ImportError:
    Sheet = None


def resource_path(relative_path):
    """ 獲取資源的絕對路徑 (打包用) """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


class ImportWizard(tk.Toplevel):
    def __init__(self, parent, save_callback):
        super().__init__(parent)
        self.title("商品資料批次匯入精靈(excel表格版)")
        self.geometry("1200x850")
        self.save_callback = save_callback 
        self.import_raw_df = pd.DataFrame()

        try:
            self.iconbitmap(resource_path("main.ico"))
        except:
            pass

        
        
        # ERP 核心必填欄位
        self.REQUIRED_FIELDS = ["商品名稱", "目前庫存", "預設成本"]
        
        self.grab_set()
        self.setup_ui()

    def setup_ui(self):
        # 頂部：檔案選取區
        header = ttk.Frame(self, padding=20)
        header.pack(fill="x")
        ttk.Label(header, text="Step 1: 開啟舊有的商品 Excel", font=("微軟正黑體", 12, "bold")).pack(side="left")
        ttk.Button(header, text="📁 選擇檔案", command=self.load_file).pack(side="left", padx=10)
        self.lbl_path = ttk.Label(header, text="尚未選取檔案", foreground="gray")
        self.lbl_path.pack(side="left")

        # 中間：雙欄佈局 (左表格預覽，右映射設定)
        paned = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        paned.pack(fill="both", expand=True, padx=20)

        # --- 左側：回歸本來的表格 UI ---
        left_f = ttk.LabelFrame(paned, text="Step 2: 原始資料預覽 (tksheet)", padding=5)
        paned.add(left_f, weight=3)
        
        if Sheet:
            self.sheet = Sheet(left_f, data=[[]], show_row_index=True)
            self.sheet.pack(fill="both", expand=True)
            self.sheet.enable_bindings()
        else:
            # 備援方案：若使用者未安裝 tksheet，才顯示文字
            self.sheet = tk.Text(left_f, wrap="none")
            self.sheet.pack(fill="both", expand=True)
            ttk.Label(left_f, text="建議安裝 tksheet 以獲得最佳表格體驗", foreground="red").pack()

        # --- 右側：欄位映射區 ---
        right_f = ttk.LabelFrame(paned, text="Step 3: ERP 欄位匹配設定", padding=10)
        paned.add(right_f, weight=1)

        self.field_keys = [
            "商品名稱", "商品編號", "分類Tag", "單位權重", 
            "目前庫存", "預設成本", "安全庫存", 
            "初始上架時間", "最後進貨時間", "商品連結", "商品備註"
        ]
        self.vars = {k: tk.StringVar(value="(不匯入 / 留空)") for k in self.field_keys}

        # 映射清單加入滾輪，防止欄位過多塞不下
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
            
            # 每一個對應欄位
            cb = ttk.Combobox(f, textvariable=self.vars[label], state="readonly")
            cb.pack(side="left", fill="x", expand=True)
            setattr(self, f"cb_{label}", cb)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        ttk.Label(right_f, text="\n* 每個欄位均可選擇「不匯入」", foreground="#d9534f", font=("", 9)).pack(anchor="w")

        # 底部：按鈕區
        footer = ttk.Frame(self, padding=20)
        footer.pack(fill="x")
        ttk.Button(footer, text="✅ 開始執行資料核對與匯入", command=self.execute_import, width=35).pack(side="right")
        ttk.Button(footer, text="❌ 取消", command=self.destroy).pack(side="right", padx=10)

    def load_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel 活頁簿", "*.xlsx"), ("舊版 Excel", "*.xls")])
        if not path: return
        try:
            self.lbl_path.config(text=f"已載入: {path.split('/')[-1]}", foreground="green")
            self.import_raw_df = pd.read_excel(path).fillna("")
            headers = self.import_raw_df.columns.tolist()
            
            # 更新 Step 2 的表格資料
            if Sheet and isinstance(self.sheet, Sheet):
                self.sheet.set_sheet_data(self.import_raw_df.values.tolist())
                self.sheet.headers(headers)
            else:
                self.sheet.delete("1.0", tk.END)
                self.sheet.insert(tk.END, self.import_raw_df.to_string())

            # 更新 Step 3 的選單選項
            options = ["(不匯入 / 留空)"] + [f"列 {i}: {h}" for i, h in enumerate(headers)]
            
            for label in self.field_keys:
                cb = getattr(self, f"cb_{label}")
                cb['values'] = options
                cb.set("(不匯入 / 留空)") # 預設重設，防止舊緩存

                # --- 智慧自動匹配邏輯 ---
                for opt in options:
                    h_low = opt.lower()
                    if label in opt: cb.set(opt); break
                    if label == "商品編號" and ("編號" in h_low or "sku" in h_low or "位置" in h_low): cb.set(opt); break
                    if label == "單位權重" and ("g" in h_low or "重量" in h_low or "weight" in h_low): cb.set(opt); break
                    if label == "分類Tag" and ("分類" in h_low or "標籤" in h_low or "tag" in h_low): cb.set(opt); break

        except Exception as e:
            messagebox.showerror("讀取失敗", f"Excel 解析錯誤: {e}")

    def execute_import(self):
        if self.import_raw_df.empty: return

        # 1. 整理匹配對應表
        mapping = {}
        for label, var in self.vars.items():
            val = var.get()
            if val != "(不匯入 / 留空)":
                mapping[label] = int(val.split(":")[0].replace("列 ", ""))

        # 2. 核心欄位檢查
        missing = [f for f in self.REQUIRED_FIELDS if f not in mapping]
        if missing:
            messagebox.showerror("映射不全", f"您漏掉了 ERP 核心必填欄位：\n{', '.join(missing)}")
            return

        # 3. 逐行資料清洗與轉換
        new_list = []
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

        for _, row in self.import_raw_df.iterrows():
            try:
                p_name = str(row.iloc[mapping["商品名稱"]]).strip()
                if not p_name or p_name.lower() == "nan": continue

                def get_val(key, default):
                    if key in mapping:
                        v = row.iloc[mapping[key]]
                        return str(v).strip() if str(v).strip() != "" else default
                    return default

                def get_num(key, default, is_float=False):
                    if key in mapping:
                        raw_v = row.iloc[mapping[key]]
                        val = pd.to_numeric(raw_v, errors='coerce')
                        if pd.isna(val): return default
                        return float(val) if is_float else int(val)
                    return default

                item = {
                    "商品編號": get_val("商品編號", ""),
                    "分類Tag": get_val("分類Tag", "未分類"),
                    "商品名稱": p_name,
                    "預設成本": get_num("預設成本", 0.0, True),
                    "目前庫存": get_num("目前庫存", 0),
                    "最後更新時間": now_str,
                    "初始上架時間": get_val("初始上架時間", now_str),
                    "最後進貨時間": get_val("最後進貨時間", ""),
                    "安全庫存": get_num("安全庫存", 0),
                    "商品連結": get_val("商品連結", "無"),
                    "商品備註": get_val("商品備註", "無"),
                    "單位權重": get_num("單位權重", 1.0, True)
                }
                new_list.append(item)
            except: continue

        if not new_list:
            messagebox.showwarning("警告", "掃描後無有效商品可匯入。")
            return

        # 4. 最終發射
        if messagebox.askyesno("匯入確認", f"已完成資料校準，準備匯入 {len(new_list)} 筆商品。\n確定執行嗎？"):
            if self.save_callback(new_list):
                messagebox.showinfo("成功", "商品資料庫已完成增量更新。")
                self.destroy()
