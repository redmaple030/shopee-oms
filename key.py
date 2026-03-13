
import tkinter as tk
from tkinter import ttk, messagebox
import hashlib

# ==========================================
# ⚠️ 重要設定：這裡的 SALT 必須跟您 ERP 主程式內的一模一樣
# ==========================================
try:
    from secrets_config import SECRET_SALT, RESCUE_SALT
except ImportError:
    # 如果找不到設定檔，跳出嚴重錯誤並關閉
    root_temp = tk.Tk()
    root_temp.withdraw()
    messagebox.showerror("環境錯誤", "找不到 secrets_config.py 設定檔！\n為了安全，產鑰工具已停止運行。")
    exit()

class KeyGenApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ERP 營運管理工具 - 授權與維修中心")
        self.root.geometry("450x520") # 稍微加大高度
        self.root.resizable(False, False)

        style = ttk.Style()
        style.configure("Title.TLabel", font=("微軟正黑體", 14, "bold"))
        style.configure("Result.TEntry", font=("Consolas", 14, "bold"), foreground="blue")
        style.configure("Rescue.TEntry", font=("Consolas", 14, "bold"), foreground="red")
        
        self.var_user = tk.StringVar()
        self.var_vip_result = tk.StringVar()
        self.var_rescue_result = tk.StringVar()

        self.setup_ui()

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding=20)
        main_frame.pack(fill="both", expand=True)

        ttk.Label(main_frame, text="🛡️ ERP 系統權限管理中心", style="Title.TLabel").pack(pady=(0, 20))

        # --- 第一部分：VIP 客戶授權 ---
        vip_frame = ttk.LabelFrame(main_frame, text="1. VIP 客戶授權生成 (針對 Email/帳號)", padding=15)
        vip_frame.pack(fill="x", pady=5)

        ttk.Label(vip_frame, text="輸入客戶帳號:").pack(anchor="w")
        self.entry_user = ttk.Entry(vip_frame, textvariable=self.var_user, font=("Arial", 11))
        self.entry_user.pack(fill="x", pady=5)
        
        btn_vip = ttk.Button(vip_frame, text="✨ 計算 VIP 啟用碼", command=self.generate_vip_code)
        btn_vip.pack(fill="x", pady=5)

        self.entry_vip = ttk.Entry(vip_frame, textvariable=self.var_vip_result, 
                                   state="readonly", justify="center", style="Result.TEntry")
        self.entry_vip.pack(fill="x", pady=5)
        
        ttk.Button(vip_frame, text="📋 複製 VIP 啟用碼", command=lambda: self.copy_to_clip(self.var_vip_result.get())).pack(fill="x")

        ttk.Separator(main_frame, orient="horizontal").pack(fill="x", pady=15)

        # --- 第二部分：系統緊急救援 ---
        rescue_frame = ttk.LabelFrame(main_frame, text="2. 系統維修救援 (開發者專用)", padding=15)
        rescue_frame.pack(fill="x", pady=5)

        ttk.Label(rescue_frame, text="適用帳號: RESCUE_ADMIN", foreground="gray").pack(anchor="w")
        
        btn_rescue = ttk.Button(rescue_frame, text="🔑 生成當前版本救援密鑰", command=self.generate_rescue_code)
        btn_rescue.pack(fill="x", pady=5)

        self.entry_rescue = ttk.Entry(rescue_frame, textvariable=self.var_rescue_result, 
                                      state="readonly", justify="center", style="Rescue.TEntry")
        self.entry_rescue.pack(fill="x", pady=5)
        
        ttk.Button(rescue_frame, text="📋 複製救援密鑰", command=lambda: self.copy_to_clip(self.var_rescue_result.get())).pack(fill="x")

    def generate_vip_code(self):
        """ 只針對 Email 生成啟用碼，不限機器，由本地 license.json 去數次數 """
        user_id = self.var_user.get().strip()
        if not user_id: 
            return
        
        # 算出的 code 給使用者
        raw_string = user_id + SECRET_SALT
        hashed = hashlib.sha256(raw_string.encode()).hexdigest()
        self.var_vip_result.set(hashed[:10].upper())

    def generate_rescue_code(self):
        """ 生成當前月份有效的 10 碼救援密鑰 """
        import datetime
        # 必須與主程式的格式完全一致
        dynamic_factor = datetime.datetime.now().strftime("%Y%m")
        
        raw_string = SECRET_SALT + RESCUE_SALT + dynamic_factor
        hashed = hashlib.sha256(raw_string.encode()).hexdigest()
        
        license_key = hashed[:10].upper()
        self.var_rescue_result.set(license_key)

    def copy_to_clip(self, content):
        if content:
            self.root.clipboard_clear()
            self.root.clipboard_append(content)
            messagebox.showinfo("成功", f"代碼 [{content}] 已複製！")
        else:
            messagebox.showwarning("提示", "內容為空無法複製")

if __name__ == "__main__":
    root = tk.Tk()
    app = KeyGenApp(root)
    root.mainloop()