#excel快速匯入插件

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from datetime import datetime

try:
    from tksheet import Sheet
except ImportError:
    Sheet = None

class ImportWizard(tk.Toplevel):
    def __init__(self, parent, save_callback):
        super().__init__(parent)
        self.title("🚀 商品資料批次匯入精靈 (安全強化版)")
        self.geometry("1100x750")
        self.save_callback = save_callback 
        self.import_raw_df = pd.DataFrame()
        
        # 定義必填欄位
        self.REQUIRED_FIELDS = ["商品名稱", "目前庫存", "預設成本"]
        
        self.grab_set()
        self.setup_ui()

    def setup_ui(self):
        # 頂部說明
        header = ttk.Frame(self, padding=20)
        header.pack(fill="x")
        ttk.Label(header, text="Step 1: 開啟 Excel 檔案", font=("", 12, "bold")).pack(side="left")
        ttk.Button(header, text="📁 選擇檔案", command=self.load_file).pack(side="left", padx=10)
        self.lbl_path = ttk.Label(header, text="尚未選取檔案", foreground="gray")
        self.lbl_path.pack(side="left")

        # 中間區域
        paned = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        paned.pack(fill="both", expand=True, padx=20)

        # 左：預覽
        left_f = ttk.LabelFrame(paned, text="Step 2: 原始資料預覽", padding=5)
        paned.add(left_f, weight=3)
        if Sheet:
            self.sheet = Sheet(left_f, data=[[]], show_row_index=True)
            self.sheet.pack(fill="both", expand=True)
            self.sheet.enable_bindings()
        else:
            ttk.Label(left_f, text="請安裝 tksheet 以獲得最佳預覽體驗").pack()

        # 右：欄位匹配
        right_f = ttk.LabelFrame(paned, text="Step 3: 欄位匹配", padding=10)
        paned.add(right_f, weight=1)

        self.fields = {
            "商品編號": tk.StringVar(value="(未匹配)"),
            "分類Tag": tk.StringVar(value="(未匹配)"),
            "商品名稱": tk.StringVar(value="(未匹配)"), # 必填
            "目前庫存": tk.StringVar(value="(未匹配)"), # 必填
            "預設成本": tk.StringVar(value="(未匹配)"), # 必填
            "安全庫存": tk.StringVar(value="(未匹配)"),
            "商品連結": tk.StringVar(value="(未匹配)"),
            "商品備註": tk.StringVar(value="(未匹配)")
        }

        for label in self.fields.keys():
            f = ttk.Frame(right_f)
            f.pack(fill="x", pady=2)
            
            # 如果是必填，顯示紅色星號
            prefix = "⭐ " if label in self.REQUIRED_FIELDS else "  "
            lbl_color = "red" if label in self.REQUIRED_FIELDS else "black"
            
            lbl = ttk.Label(f, text=f"{prefix}{label}:", width=12)
            lbl.pack(side="left")
            
            cb = ttk.Combobox(f, textvariable=self.fields[label], state="readonly")
            cb.pack(side="left", fill="x", expand=True)
            setattr(self, f"cb_{label}", cb)

        ttk.Label(right_f, text="\n⭐ 為必填項目，否則無法匯入", foreground="red", font=("", 9)).pack(anchor="w")

        # 底部
        footer = ttk.Frame(self, padding=20)
        footer.pack(fill="x")
        ttk.Button(footer, text="✅ 執行安全匯入", command=self.execute_import, width=25, style="Accent.TButton").pack(side="right")
        ttk.Button(footer, text="❌ 取消", command=self.destroy).pack(side="right", padx=10)

    def load_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel 活頁簿", "*.xlsx"), ("舊版 Excel", "*.xls")])
        if not path: return
        try:
            # 讀取時將所有資料轉為字串處理，避免讀取時就出錯
            self.import_raw_df = pd.read_excel(path).fillna("")
            headers = self.import_raw_df.columns.tolist()
            
            if Sheet:
                self.sheet.set_sheet_data(self.import_raw_df.values.tolist())
                self.sheet.headers(headers)

            options = ["(未匹配)"] + [f"列 {i}: {h}" for i, h in enumerate(headers)]
            for label in self.fields.keys():
                cb = getattr(self, f"cb_{label}")
                cb['values'] = options
                # 智慧自動匹配
                for opt in options:
                    if label in opt or (label == "商品編號" and "位置" in opt):
                        cb.set(opt); break
        except Exception as e:
            messagebox.showerror("錯誤", f"讀取失敗: {e}")

    def execute_import(self):
        if self.import_raw_df.empty: return

        # 第一道防線：檢查必填項目的「對應關係」是否有選
        mapping = {}
        missing_mapping = []
        for label, var in self.fields.items():
            val = var.get()
            if val != "(未匹配)":
                mapping[label] = int(val.split(":")[0].replace("列 ", ""))
            elif label in self.REQUIRED_FIELDS:
                missing_mapping.append(label)

        if missing_mapping:
            messagebox.showerror("欄位缺失", f"請先對應以下必填欄位：\n{', '.join(missing_mapping)}")
            return

        # 第二道防線：資料轉換與清洗
        new_list = []
        skip_count = 0
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

        for idx, row in self.import_raw_df.iterrows():
            try:
                # 1. 檢查商品名稱 (絕對不能空白)
                p_name = str(row.iloc[mapping["商品名稱"]]).strip()
                if not p_name or p_name.lower() == "nan":
                    skip_count += 1
                    continue

                # 2. 庫存清洗 (轉數字，失敗則補 0)
                raw_stock = row.iloc[mapping["目前庫存"]]
                stock = int(pd.to_numeric(raw_stock, errors='coerce')) if pd.notna(pd.to_numeric(raw_stock, errors='coerce')) else 0

                # 3. 成本清洗 (轉數字，失敗則補 0.0)
                raw_cost = row.iloc[mapping["預設成本"]]
                cost = float(pd.to_numeric(raw_cost, errors='coerce')) if pd.notna(pd.to_numeric(raw_cost, errors='coerce')) else 0.0

                item = {
                    "商品編號": str(row.iloc[mapping["商品編號"]]).strip() if "商品編號" in mapping else "",
                    "分類Tag": row.iloc[mapping["分類Tag"]] if "分類Tag" in mapping else "未分類",
                    "商品名稱": p_name,
                    "預設成本": cost,
                    "目前庫存": stock,
                    "最後更新時間": now_str,
                    "初始上架時間": now_str,
                    "最後進貨時間": "",
                    "安全庫存": int(pd.to_numeric(row.iloc[mapping["安全庫存"]], errors='coerce')) if "安全庫存" in mapping else 0,
                    "商品連結": row.iloc[mapping["商品連結"]] if "商品連結" in mapping else "無",
                    "商品備註": row.iloc[mapping["商品備註"]] if "商品備註" in mapping else "無"
                }
                new_list.append(item)
            except Exception:
                skip_count += 1
                continue

        if not new_list:
            messagebox.showwarning("警告", "沒有找到任何有效的商品資料可供匯入！")
            return

        # 第三道防線：匯入確認
        msg = f"準備匯入 {len(new_list)} 筆商品。"
        if skip_count > 0:
            msg += f"\n(注意：已自動跳過 {skip_count} 筆名稱空白或格式錯誤的資料)"
        
        if messagebox.askyesno("匯入確認", msg):
            if self.save_callback(new_list):
                messagebox.showinfo("成功", "資料匯入完成！")

                self.destroy()
