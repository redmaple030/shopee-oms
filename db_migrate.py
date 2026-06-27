"""
db_migrate.py  ─  Excel → SQLite 一次性資料遷移
=================================================
執行方式（在 main.py 同一資料夾下執行）：
    python db_migrate.py

功能：
  1. 讀取 sales_data.xlsx 所有分頁
  2. 清理資料（NaN、格式、欄位名稱）
  3. 全部匯入 erp_local.db
  4. 印出每張表的筆數確認

執行後不會刪除 Excel 檔案，程式仍可正常運作。
"""

import os
import sys
import pandas as pd
from datetime import datetime

# ── 確保能找到 local_db.py ────────────────────────────────────────────────────
script_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, script_dir)

from local_db import DB, SHEET_TO_TABLE, COL_RENAME_TO_DB

# ── Excel 路徑（與 main.py 同一資料夾）────────────────────────────────────────
EXCEL_PATH = os.path.join(script_dir, "sales_data.xlsx")


def clean_string(val) -> str:
    """把 NaN / None / 'nan' 統一轉成空字串"""
    if val is None:
        return ""
    s = str(val).strip()
    if s.lower() in ("nan", "none", "nat", ""):
        return ""
    # 移除 Excel 防科學記號用的前置單引號
    if s.startswith("'"):
        s = s[1:]
    return s


def clean_number(val, as_int=False):
    """NaN → 0，並轉成正確型別"""
    try:
        if pd.isna(val):
            return 0 if as_int else 0.0
    except Exception:
        pass
    try:
        return int(float(val)) if as_int else float(val)
    except Exception:
        return 0 if as_int else 0.0


# ── 各表專用清理函式 ───────────────────────────────────────────────────────────

def migrate_products(df: pd.DataFrame) -> pd.DataFrame:
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    rows = []
    for _, r in df.iterrows():
        rows.append({
            "商品名稱":     clean_string(r.get("商品名稱", "")),
            "分類Tag":      clean_string(r.get("分類Tag", "")),
            "預設成本":     clean_number(r.get("預設成本", 0)),
            "預設售價":     clean_number(r.get("預設售價", 0)),
            "售價":         clean_number(r.get("售價", 0)),
            "目前庫存":     clean_number(r.get("目前庫存", 0), as_int=True),
            "安全庫存":     clean_number(r.get("安全庫存", 0)),
            "單位權重":     clean_number(r.get("單位權重", 1)) or 1.0,
            "商品編號":     clean_string(r.get("商品編號", "")),
            "商品連結":     clean_string(r.get("商品連結", "")),
            "商品備註":     clean_string(r.get("商品備註", "")),
            "最後更新時間": clean_string(r.get("最後更新時間", "")),
            "初始上架時間": clean_string(r.get("初始上架時間", "")),
            "最後進貨時間": clean_string(r.get("最後進貨時間", "")),
            "last_modified": now,
        })
    result = pd.DataFrame(rows)
    # 移除商品名稱為空的行
    return result[result["商品名稱"] != ""].reset_index(drop=True)


def migrate_sales_like(df: pd.DataFrame) -> pd.DataFrame:
    """銷售紀錄、訂單追蹤、退貨紀錄共用相同欄位結構"""
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    rows = []
    for _, r in df.iterrows():
        rows.append({
            "訂單編號":   clean_string(r.get("訂單編號", "")),
            "日期":       clean_string(r.get("日期", "")),
            "買家名稱":   clean_string(r.get("買家名稱", "")),
            "交易平台":   clean_string(r.get("交易平台", "")),
            "寄送方式":   clean_string(r.get("寄送方式", "")),
            "取貨地點":   clean_string(r.get("取貨地點", "")),
            "商品名稱":   clean_string(r.get("商品名稱", "")),
            "商品編號":   clean_string(r.get("商品編號", "")),
            "數量":       clean_number(r.get("數量", 0), as_int=True),
            "單價_售":    clean_number(r.get("單價(售)", 0)),
            "單價_進":    clean_number(r.get("單價(進)", 0)),
            "總銷售額":   clean_number(r.get("總銷售額", 0)),
            "總成本":     clean_number(r.get("總成本", 0)),
            "分攤手續費": clean_number(r.get("分攤手續費", 0)),
            "扣費項目":   clean_string(r.get("扣費項目", "")),
            "總淨利":     clean_number(r.get("總淨利", 0)),
            "毛利率":     clean_number(r.get("毛利率", 0)),
            "稅額":       clean_number(r.get("稅額", 0)),
            "備註":       clean_string(r.get("備註", "")),
            "last_modified": now,
        })
    return pd.DataFrame(rows)


def migrate_purchases_like(df: pd.DataFrame) -> pd.DataFrame:
    """進貨紀錄、進貨追蹤共用"""
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    rows = []
    for _, r in df.iterrows():
        rows.append({
            "進貨單號":           clean_string(r.get("進貨單號", "")),
            "採購日期":           clean_string(r.get("採購日期", "")),
            "入庫日期":           clean_string(r.get("入庫日期", "")),
            "供應商":             clean_string(r.get("供應商", "")),
            "物流狀態":           clean_string(r.get("物流狀態", "")),
            "商品名稱":           clean_string(r.get("商品名稱", "")),
            "數量":               clean_number(r.get("數量", 0), as_int=True),
            "原始預計數量":       clean_number(r.get("原始預計數量", 0), as_int=True),
            "瑕疵數量":           clean_number(r.get("瑕疵數量", 0), as_int=True),
            "進貨單價":           clean_number(r.get("進貨單價", 0)),
            "進貨總額":           clean_number(r.get("進貨總額", 0)),
            "進項稅額":           clean_number(r.get("進項稅額", 0)),
            "分攤運費":           clean_number(r.get("分攤運費", 0)),
            "海關稅金":           clean_number(r.get("海關稅金", 0)),
            "賣家交付日期":       clean_string(r.get("賣家交付日期", "")),
            "備註":               clean_string(r.get("備註", "")),
            "物流追蹤":           clean_string(r.get("物流追蹤", "")),
            "時間_廠商出貨":      clean_string(r.get("時間_廠商出貨", "")),
            "時間_抵達集運倉":    clean_string(r.get("時間_抵達集運倉", "")),
            "時間_集運倉出貨":    clean_string(r.get("時間_集運倉出貨", "")),
            "時間_抵達台灣海關":  clean_string(r.get("時間_抵達台灣海關", "")),
            "時間_國內配送中":    clean_string(r.get("時間_國內配送中", "")),
            "last_modified":      now,
        })
    return pd.DataFrame(rows)


def migrate_vendors(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, r in df.iterrows():
        name = clean_string(r.get("廠商名稱", ""))
        if not name:
            continue
        rows.append({
            "廠商名稱":     name,
            "通路":         clean_string(r.get("通路", "")),
            "統編":         clean_string(r.get("統編", "")),
            "聯絡人":       clean_string(r.get("聯絡人", "")),
            "電話":         clean_string(r.get("電話", "")),
            "地址":         clean_string(r.get("地址", "")),
            "備註":         clean_string(r.get("備註", "")),
            "平均前置天數": clean_number(r.get("平均前置天數", 0)),
            "總到貨率":     clean_string(r.get("總到貨率", "")),
            "總合格率":     clean_string(r.get("總合格率", "")),
            "綜合評等分數": clean_number(r.get("綜合評等分數", 0)),
            "星等":         clean_number(r.get("星等", 5), as_int=True),
            "最後更新":     clean_string(r.get("最後更新", "")),
        })
    return pd.DataFrame(rows)


def migrate_after_sales(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, r in df.iterrows():
        rows.append({
            "訂單編號": clean_string(r.get("訂單編號", "")),
            "商品名稱": clean_string(r.get("商品名稱", "")),
            "發生日期": clean_string(r.get("發生日期", "")),
            "處理類型": clean_string(r.get("處理類型", "")),
            "支出金額": clean_number(r.get("支出金額", 0)),
            "詳細說明": clean_string(r.get("詳細說明", "")),
            "原始買家": clean_string(r.get("原始買家", "")),
        })
    return pd.DataFrame(rows)


def migrate_fee_settings(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, r in df.iterrows():
        name = clean_string(r.get("設定名稱", ""))
        if not name:
            continue
        rows.append({
            "設定名稱":   name,
            "費率百分比": clean_number(r.get("費率百分比", 0)),
            "固定金額":   clean_number(r.get("固定金額", 0)),
        })
    return pd.DataFrame(rows)


def migrate_sys_settings(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, r in df.iterrows():
        name = clean_string(r.get("設定名稱", ""))
        if not name:
            continue
        rows.append({
            "設定名稱": name,
            "參數值":   clean_string(r.get("參數值", "")),
        })
    return pd.DataFrame(rows)


# ── 主要遷移流程 ───────────────────────────────────────────────────────────────
def run_migration():
    if not os.path.exists(EXCEL_PATH):
        print(f"[錯誤] 找不到 Excel 檔案：{EXCEL_PATH}")
        sys.exit(1)

    print(f"[開始] 讀取 {EXCEL_PATH} ...")
    print(f"[目標] SQLite 資料庫：{DB.db_path}\n")

    xls = pd.ExcelFile(EXCEL_PATH)

    # sheet_name → (pandas df, 清理函式, db表名)
    tasks = [
        ("商品資料",     migrate_products,       "products"),
        ("銷售紀錄",     migrate_sales_like,     "sales_history"),
        ("訂單追蹤",     migrate_sales_like,     "sales_tracking"),
        ("進貨紀錄",     migrate_purchases_like, "purchases"),
        ("進貨追蹤",     migrate_purchases_like, "purchase_tracking"),
        ("退貨紀錄",     migrate_sales_like,     "returns"),
        ("手續費設定",   migrate_fee_settings,   "fee_settings"),
        ("系統設定",     migrate_sys_settings,   "sys_settings"),
        ("進貨廠商管理", migrate_vendors,         "vendors"),
        ("售後明細",     migrate_after_sales,    "after_sales"),
    ]

    success = 0
    fail = 0

    for sheet_name, clean_fn, table_name in tasks:
        try:
            if sheet_name not in xls.sheet_names:
                print(f"  [略過] {sheet_name} (分頁不存在)")
                continue

            df_raw = pd.read_excel(xls, sheet_name=sheet_name)
            df_clean = clean_fn(df_raw)

            # sales_history / sales_tracking 沒有備註欄，移除多餘欄位
            if table_name in ("sales_history", "sales_tracking"):
                df_clean = df_clean.drop(columns=["備註"], errors="ignore")

            # 使用直接 SQL 寫入（繞過 write() 的 empty 保護，允許空表遷移）
            import sqlite3
            conn = sqlite3.connect(DB.db_path)
            conn.execute(f"DELETE FROM {table_name}")
            if not df_clean.empty:
                df_clean.to_sql(table_name, conn, if_exists="append", index=False)
            conn.commit()
            conn.close()

            print(f"  [OK] {sheet_name:12s} → {table_name:20s} ({len(df_clean)} 筆)")
            success += 1

        except Exception as e:
            import traceback
            print(f"  [FAIL] {sheet_name}: {e}")
            traceback.print_exc()
            fail += 1

    print(f"\n{'='*50}")
    print(f"遷移完成：{success} 張表成功，{fail} 張失敗")
    print(f"{'='*50}\n")

    # 驗證
    print("驗證 SQLite 資料筆數：")
    info = DB.info()
    for table, count in info.items():
        sheet = {v: k for k, v in SHEET_TO_TABLE.items()}.get(table, table)
        print(f"  {sheet:12s} ({table:20s}): {count} 筆")

    print("\n[完成] 現在可以在 main.py 中改用 from local_db import DB")
    print("[注意] Excel 檔案仍保留，不會被刪除\n")


if __name__ == "__main__":
    # 防止重複遷移
    db_info = DB.info()
    total_rows = sum(db_info.values())

    if total_rows > 0:
        ans = input(
            f"\n[警告] SQLite 已有 {total_rows} 筆資料，重新遷移會覆蓋所有資料。\n"
            "確定要繼續嗎？(y/N): "
        ).strip().lower()
        if ans != "y":
            print("已取消。")
            sys.exit(0)

    run_migration()
