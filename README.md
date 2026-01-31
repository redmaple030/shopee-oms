
# 🛒 蝦皮/網拍專用進銷存管理系統 (Shopee OMS Local-First)

> **針對 2026 蝦皮手續費新制開發，堅持「資料落地」的輕量級 ERP 解決方案。**

版本：3.5
發布日期：2026-01-31
適用對象：個人賣家、團購主、網拍經營者


更新內容:

版本:3.5 2026-01-31
修正"平台"及"買家名稱"顯示問題
加入資料清洗功能預防買家姓名錯誤顯示的狀況
優化google備份的流暢度


版本:3.4 2026-01-29
google雲端備份及還原功能 以測試完成(付費解鎖)
新增簡易統計表單
預計新增訂單追蹤
貨物採購物流試算等


版本:3.0 2026-01-28
目前正在測試個人google雲端備份及還原功能(測試中)
字體大小現在可以做調整了(遠視或老花友善)




版本:2.4 2026-01-26
修正部分UI介面及提示 (銷售輸入頁面)
增加庫存數量功能及缺貨提醒
增加交易平台功能(增加複數除蝦皮外平台供選擇)
優化商品排序美觀



版本：2.1 2026-01-23
修正部分UI介面及提示
增加蝦皮手續費顯示選擇


版本：2.0 2026-01-22
初版程式公開


## 📖 專案背景 (Project Background)

隨著 2026 年各大電商平台手續費結構調整，加上雲端服務的不確定性（帳號停權、資安疑慮），許多微型賣家面臨數位轉型的兩難。

本專案旨在開發一套 **「本地優先 (Local-First)」** 的進銷存系統，結合 **Python 強大的數據處理能力** 與 **Google Drive API 雲端備份**，讓賣家既能擁有傳統軟體的資料掌控權，又能享有現代化的雲端備份便利性。

## ✨ 核心功能 (Key Features)

### 1. 🛡️ 資料主權與安全性
- **本地運算**：所有銷售數據皆儲存於本地 Excel/Database，不依賴第三方伺服器，斷網亦可操作。
- **混合雲備份**：整合 Google Drive API，透過 OAuth2 驗證，一鍵將加密資料備份至個人雲端硬碟。

### 2. ⚡ 針對電商優化的 UX 設計
- **自動計算**：內建 2026 年最新手續費費率（含促銷檔期/商城費率），自動計算毛利。
- **庫存防呆**：即時庫存扣除邏輯，防止超賣。
- **長輩友善**：支援介面字體動態縮放 (10pt - 20pt)，自動調整表格行高。

### 3. 🔐 VIP 授權機制 (商業化模組)
- **Hash 驗證**：採用 MD5 演算法結合 Salt 進行軟體授權驗證。
- **雙重控管**：結合「軟體啟用碼」與「Google API 白名單」機制，確保付費會員權益。
- **多執行緒優化**：將登入與備份動作移至背景執行緒 (Threading)，確保 UI 操作流暢不卡頓。

## 🛠️ 技術棧 (Tech Stack)

- **Language**: Python 3.x
- **GUI Framework**: Tkinter / Ttk (Native Look & Feel)
- **Data Manipulation**: Pandas (高效能數據處理)
- **Cloud Integration**: Google Drive API v3, OAuthLib
- **Concurrency**: Python `threading` module
- **Security**: Hashlib (License Key Generation)
- **Packaging**: PyInstaller

## 🚀 安裝與執行 (Installation)

### 開發者模式 (For Developers)
如果您希望研究原始碼或進行二次開發：

1. **Clone 專案**
   ```bash
   git clone https://github.com/您的帳號/Shopee-OMS-System.git
   cd Shopee-OMS-System
安裝依賴
code
Bash
pip install pandas openpyxl google-api-python-client google-auth-oauthlib
設定 Google API
請自行至 Google Cloud Console 申請 credentials.json 並放入專案根目錄。
新增 secrets_config.py 並設定您的 SECRET_SALT。
執行程式
code
Bash
python SalesApp.py
一般使用者 (For Users)
本專案提供打包好的執行檔 (EXE)，無需安裝 Python 環境即可使用。
[下載連結] (此處請放入您的試用版下載連結)
💼 商業模式 (Business Model)
本軟體採用 Freemium (免費增值) 模式：
社群版：開源免費，支援基本進銷存功能。
VIP 專業版：採買斷制，解鎖「一鍵雲端備份」與「無限筆數」功能。
⚠️ 免責聲明 (Disclaimer)
本軟體為個人開發作品，僅供輔助使用。開發者不對因軟體錯誤或資料遺失造成的商業損失負責，請務必定期手動備份重要資料。
