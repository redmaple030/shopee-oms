"""
local_db.py  ─  ERP 本地 SQLite 資料層
=========================================
取代 sales_data.xlsx，所有分頁對應一張 SQLite 資料表。

使用方式：
    from local_db import DB

    # 讀取（回傳 pd.DataFrame，與原本 pd.read_excel 相容）
    df = DB.read("products")

    # 寫入（完全取代該表，與原本 _universal_save 行為相同）
    DB.write("products", df)

    # 只更新部分欄位（單行 upsert）
    DB.upsert("products", {"商品名稱": "xxx", "目前庫存": 5}, pk="商品名稱")

資料表對應：
    Excel 分頁名稱        SQLite 表名
    ──────────────────────────────────
    商品資料              products
    銷售紀錄              sales_history
    訂單追蹤              sales_tracking
    進貨紀錄              purchases
    進貨追蹤              purchase_tracking
    退貨紀錄              returns
    手續費設定            fee_settings
    系統設定              sys_settings
    進貨廠商管理          vendors
    售後明細              after_sales
"""

import sqlite3
import threading
import os
import pandas as pd
from datetime import datetime
from typing import Optional

# ── 設定檔案路徑 ──────────────────────────────────────────────────────────────
DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "erp_local.db")

# Excel 分頁名稱 ↔ SQLite 表名 對照
SHEET_TO_TABLE = {
    "商品資料":     "products",
    "銷售紀錄":     "sales_history",
    "訂單追蹤":     "sales_tracking",
    "進貨紀錄":     "purchases",
    "進貨追蹤":     "purchase_tracking",
    "退貨紀錄":     "returns",
    "手續費設定":   "fee_settings",
    "系統設定":     "sys_settings",
    "進貨廠商管理": "vendors",
    "售後明細":     "after_sales",
}
TABLE_TO_SHEET = {v: k for k, v in SHEET_TO_TABLE.items()}

# ── 每張表的欄位定義（型別對應 SQLite）─────────────────────────────────────────
SCHEMAS = {
    "products": """
        CREATE TABLE IF NOT EXISTS products (
            商品名稱        TEXT PRIMARY KEY,
            分類Tag         TEXT,
            預設成本        REAL DEFAULT 0,
            預設售價        REAL DEFAULT 0,
            售價            REAL DEFAULT 0,
            目前庫存        INTEGER DEFAULT 0,
            安全庫存        REAL DEFAULT 0,
            單位權重        REAL DEFAULT 1,
            商品編號        TEXT,
            商品連結        TEXT,
            商品備註        TEXT,
            最後更新時間    TEXT,
            初始上架時間    TEXT,
            最後進貨時間    TEXT,
            last_modified   TEXT
        )""",

    "sales_history": """
        CREATE TABLE IF NOT EXISTS sales_history (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            訂單編號        TEXT,
            日期            TEXT,
            買家名稱        TEXT,
            交易平台        TEXT,
            寄送方式        TEXT,
            取貨地點        TEXT,
            商品名稱        TEXT,
            商品編號        TEXT,
            數量            INTEGER DEFAULT 0,
            單價_售         REAL DEFAULT 0,
            單價_進         REAL DEFAULT 0,
            總銷售額        REAL DEFAULT 0,
            總成本          REAL DEFAULT 0,
            分攤手續費      REAL DEFAULT 0,
            扣費項目        TEXT,
            總淨利          REAL DEFAULT 0,
            毛利率          REAL DEFAULT 0,
            稅額            REAL DEFAULT 0,
            last_modified   TEXT
        )""",

    "sales_tracking": """
        CREATE TABLE IF NOT EXISTS sales_tracking (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            訂單編號        TEXT,
            日期            TEXT,
            買家名稱        TEXT,
            交易平台        TEXT,
            寄送方式        TEXT,
            取貨地點        TEXT,
            商品名稱        TEXT,
            商品編號        TEXT,
            數量            INTEGER DEFAULT 0,
            單價_售         REAL DEFAULT 0,
            單價_進         REAL DEFAULT 0,
            總銷售額        REAL DEFAULT 0,
            總成本          REAL DEFAULT 0,
            分攤手續費      REAL DEFAULT 0,
            扣費項目        TEXT,
            總淨利          REAL DEFAULT 0,
            毛利率          REAL DEFAULT 0,
            稅額            REAL DEFAULT 0,
            last_modified   TEXT
        )""",

    "purchases": """
        CREATE TABLE IF NOT EXISTS purchases (
            id                  INTEGER PRIMARY KEY AUTOINCREMENT,
            進貨單號            TEXT,
            採購日期            TEXT,
            入庫日期            TEXT,
            供應商              TEXT,
            物流狀態            TEXT,
            商品名稱            TEXT,
            數量                INTEGER DEFAULT 0,
            原始預計數量        INTEGER DEFAULT 0,
            瑕疵數量            INTEGER DEFAULT 0,
            進貨單價            REAL DEFAULT 0,
            進貨總額            REAL DEFAULT 0,
            進項稅額            REAL DEFAULT 0,
            分攤運費            REAL DEFAULT 0,
            海關稅金            REAL DEFAULT 0,
            賣家交付日期        TEXT,
            備註                TEXT,
            物流追蹤            TEXT,
            時間_廠商出貨       TEXT,
            時間_抵達集運倉     TEXT,
            時間_集運倉出貨     TEXT,
            時間_抵達台灣海關   TEXT,
            時間_國內配送中     TEXT,
            last_modified       TEXT
        )""",

    "purchase_tracking": """
        CREATE TABLE IF NOT EXISTS purchase_tracking (
            id                  INTEGER PRIMARY KEY AUTOINCREMENT,
            進貨單號            TEXT,
            採購日期            TEXT,
            入庫日期            TEXT,
            供應商              TEXT,
            物流狀態            TEXT,
            商品名稱            TEXT,
            數量                INTEGER DEFAULT 0,
            原始預計數量        INTEGER DEFAULT 0,
            瑕疵數量            INTEGER DEFAULT 0,
            進貨單價            REAL DEFAULT 0,
            進貨總額            REAL DEFAULT 0,
            進項稅額            REAL DEFAULT 0,
            分攤運費            REAL DEFAULT 0,
            海關稅金            REAL DEFAULT 0,
            賣家交付日期        TEXT,
            備註                TEXT,
            物流追蹤            TEXT,
            時間_廠商出貨       TEXT,
            時間_抵達集運倉     TEXT,
            時間_集運倉出貨     TEXT,
            時間_抵達台灣海關   TEXT,
            時間_國內配送中     TEXT,
            last_modified       TEXT
        )""",

    "returns": """
        CREATE TABLE IF NOT EXISTS returns (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            訂單編號        TEXT,
            日期            TEXT,
            買家名稱        TEXT,
            交易平台        TEXT,
            寄送方式        TEXT,
            取貨地點        TEXT,
            商品名稱        TEXT,
            商品編號        TEXT,
            數量            INTEGER DEFAULT 0,
            單價_售         REAL DEFAULT 0,
            單價_進         REAL DEFAULT 0,
            總銷售額        REAL DEFAULT 0,
            總成本          REAL DEFAULT 0,
            分攤手續費      REAL DEFAULT 0,
            扣費項目        TEXT,
            總淨利          REAL DEFAULT 0,
            毛利率          REAL DEFAULT 0,
            稅額            REAL DEFAULT 0,
            備註            TEXT,
            last_modified   TEXT
        )""",

    "fee_settings": """
        CREATE TABLE IF NOT EXISTS fee_settings (
            設定名稱    TEXT PRIMARY KEY,
            費率百分比  REAL DEFAULT 0,
            固定金額    REAL DEFAULT 0
        )""",

    "sys_settings": """
        CREATE TABLE IF NOT EXISTS sys_settings (
            設定名稱    TEXT PRIMARY KEY,
            參數值      TEXT
        )""",

    "vendors": """
        CREATE TABLE IF NOT EXISTS vendors (
            廠商名稱        TEXT PRIMARY KEY,
            通路            TEXT,
            統編            TEXT,
            聯絡人          TEXT,
            電話            TEXT,
            地址            TEXT,
            備註            TEXT,
            平均前置天數    REAL DEFAULT 0,
            總到貨率        TEXT,
            總合格率        TEXT,
            綜合評等分數    REAL DEFAULT 0,
            星等            INTEGER DEFAULT 5,
            最後更新        TEXT
        )""",

    "after_sales": """
        CREATE TABLE IF NOT EXISTS after_sales (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            訂單編號    TEXT,
            商品名稱    TEXT,
            發生日期    TEXT,
            處理類型    TEXT,
            支出金額    REAL DEFAULT 0,
            詳細說明    TEXT,
            原始買家    TEXT
        )""",
}

# Excel 欄位名稱中有括號，SQLite 不允許，建立對照
# 讀出時自動把 _ 換回括號，寫入時自動把括號換成 _
COL_RENAME_TO_DB = {
    "單價(售)": "單價_售",
    "單價(進)": "單價_進",
}
COL_RENAME_FROM_DB = {v: k for k, v in COL_RENAME_TO_DB.items()}


# ── 核心資料庫類別 ─────────────────────────────────────────────────────────────
class LocalDatabase:
    """
    SQLite 存取層。所有方法都是執行緒安全的（使用 RLock）。
    對外介面刻意設計成與原本 Excel 讀寫行為相容，
    方便 main.py 逐步替換。
    """

    def __init__(self, db_path: str = DB_PATH):
        self.db_path = db_path
        self._lock = threading.RLock()
        self._init_db()

    # ── 初始化 ──────────────────────────────────────────────────────────────────
    def _init_db(self):
        with self._lock:
            conn = self._connect()
            cur = conn.cursor()
            for ddl in SCHEMAS.values():
                cur.executescript(ddl)
            conn.commit()
            conn.close()

    def _connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.db_path, timeout=10)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL")   # 允許讀寫並發
        conn.execute("PRAGMA foreign_keys=ON")
        return conn

    # ── 取得真實表名 ────────────────────────────────────────────────────────────
    def _resolve_table(self, name: str) -> str:
        """接受 Excel 分頁名或 SQLite 表名，統一回傳表名"""
        if name in SHEET_TO_TABLE:
            return SHEET_TO_TABLE[name]
        if name in SCHEMAS:
            return name
        raise ValueError(f"未知的表格名稱: {name}")

    # ── 欄位名稱轉換 ────────────────────────────────────────────────────────────
    def _cols_to_db(self, df: pd.DataFrame) -> pd.DataFrame:
        return df.rename(columns=COL_RENAME_TO_DB)

    def _cols_from_db(self, df: pd.DataFrame) -> pd.DataFrame:
        return df.rename(columns=COL_RENAME_FROM_DB)

    # ── 讀取（回傳 DataFrame，與 pd.read_excel 相容）────────────────────────────
    def read(self, table_or_sheet: str) -> pd.DataFrame:
        """
        等同於 pd.read_excel(FILE_NAME, sheet_name=SHEET_XXX)
        """
        table = self._resolve_table(table_or_sheet)
        with self._lock:
            conn = self._connect()
            try:
                df = pd.read_sql(f"SELECT * FROM {table}", conn)
                # 移除內部欄位
                drop_cols = ["id", "last_modified"]
                df = df.drop(columns=[c for c in drop_cols if c in df.columns])
                df = self._cols_from_db(df)
                return df
            except Exception as e:
                print(f"[DB] read error ({table}): {e}")
                return pd.DataFrame()
            finally:
                conn.close()

    # ── 完整覆蓋寫入（等同於 _universal_save 中對單表的操作）───────────────────
    def write(self, table_or_sheet: str, df: pd.DataFrame) -> bool:
        """
        完全取代該表的資料，行為與原本 _universal_save 相同。
        空 DataFrame 會被拒絕（與原本保護邏輯一致）。
        """
        if df is None or df.empty:
            print(f"[DB] write blocked: empty DataFrame for {table_or_sheet}")
            return False

        table = self._resolve_table(table_or_sheet)
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        with self._lock:
            conn = self._connect()
            try:
                df_clean = df.copy()
                df_clean = self._cols_to_db(df_clean)

                # 清理 NaN
                df_clean = df_clean.where(pd.notna(df_clean), None)

                # 加上 last_modified（若表有這欄）
                if "last_modified" in SCHEMAS.get(table, ""):
                    df_clean["last_modified"] = now

                # 取得該表實際存在的欄位
                cur = conn.cursor()
                cur.execute(f"PRAGMA table_info({table})")
                db_cols = {row["name"] for row in cur.fetchall()}

                # 只保留資料庫有的欄位（多餘欄位略過）
                valid_cols = [c for c in df_clean.columns if c in db_cols]
                df_clean = df_clean[valid_cols]

                # 原子性寫入：先清空再插入
                conn.execute(f"DELETE FROM {table}")
                df_clean.to_sql(table, conn, if_exists="append", index=False)
                conn.commit()
                return True
            except Exception as e:
                conn.rollback()
                print(f"[DB] write error ({table}): {e}")
                return False
            finally:
                conn.close()

    # ── 批次覆蓋（等同於 _universal_save 傳入字典）──────────────────────────────
    def write_many(self, updates: dict) -> bool:
        """
        等同於 self._universal_save({SHEET_A: df_a, SHEET_B: df_b})
        傳入 {表名或分頁名: DataFrame} 字典，全部成功才算成功。
        """
        results = []
        for table_or_sheet, df in updates.items():
            results.append(self.write(table_or_sheet, df))
        return all(results)

    # ── 單行 Upsert（更新單筆商品庫存、廠商資料等用）──────────────────────────
    def upsert(self, table_or_sheet: str, row: dict, pk: str) -> bool:
        """
        根據主鍵插入或更新單筆記錄。
        常用於：更新商品庫存、修改廠商評等等不需要整張表重寫的場合。

        範例：
            DB.upsert("products", {"商品名稱": "xxx", "目前庫存": 5}, pk="商品名稱")
        """
        table = self._resolve_table(table_or_sheet)
        with self._lock:
            conn = self._connect()
            try:
                row_clean = {}
                for k, v in row.items():
                    db_key = COL_RENAME_TO_DB.get(k, k)
                    row_clean[db_key] = None if pd.isna(v) else v

                cols = ", ".join(row_clean.keys())
                placeholders = ", ".join(["?" for _ in row_clean])
                updates_clause = ", ".join(
                    [f"{k}=excluded.{k}" for k in row_clean if k != pk]
                )
                db_pk = COL_RENAME_TO_DB.get(pk, pk)
                sql = (
                    f"INSERT INTO {table} ({cols}) VALUES ({placeholders}) "
                    f"ON CONFLICT({db_pk}) DO UPDATE SET {updates_clause}"
                )
                conn.execute(sql, list(row_clean.values()))
                conn.commit()
                return True
            except Exception as e:
                conn.rollback()
                print(f"[DB] upsert error ({table}): {e}")
                return False
            finally:
                conn.close()

    # ── 刪除單筆或條件刪除 ──────────────────────────────────────────────────────
    def delete(self, table_or_sheet: str, where: dict) -> bool:
        """
        範例：DB.delete("sales_tracking", {"訂單編號": "'20260101"})
        """
        table = self._resolve_table(table_or_sheet)
        with self._lock:
            conn = self._connect()
            try:
                conditions = " AND ".join(
                    [f"{COL_RENAME_TO_DB.get(k, k)}=?" for k in where]
                )
                conn.execute(
                    f"DELETE FROM {table} WHERE {conditions}",
                    list(where.values())
                )
                conn.commit()
                return True
            except Exception as e:
                conn.rollback()
                print(f"[DB] delete error ({table}): {e}")
                return False
            finally:
                conn.close()

    # ── 查詢（回傳 DataFrame）───────────────────────────────────────────────────
    def query(self, table_or_sheet: str, where: Optional[dict] = None) -> pd.DataFrame:
        """
        簡易條件查詢。複雜查詢請直接用 raw_sql()。
        範例：DB.query("products", {"分類Tag": "塔散"})
        """
        table = self._resolve_table(table_or_sheet)
        with self._lock:
            conn = self._connect()
            try:
                if where:
                    conditions = " AND ".join(
                        [f"{COL_RENAME_TO_DB.get(k, k)}=?" for k in where]
                    )
                    sql = f"SELECT * FROM {table} WHERE {conditions}"
                    df = pd.read_sql(sql, conn, params=list(where.values()))
                else:
                    df = pd.read_sql(f"SELECT * FROM {table}", conn)

                drop_cols = ["id", "last_modified"]
                df = df.drop(columns=[c for c in drop_cols if c in df.columns])
                return self._cols_from_db(df)
            except Exception as e:
                print(f"[DB] query error ({table}): {e}")
                return pd.DataFrame()
            finally:
                conn.close()

    # ── 原始 SQL（進階用）──────────────────────────────────────────────────────
    def raw_sql(self, sql: str, params=()) -> pd.DataFrame:
        with self._lock:
            conn = self._connect()
            try:
                df = pd.read_sql(sql, conn, params=params)
                return self._cols_from_db(df)
            except Exception as e:
                print(f"[DB] raw_sql error: {e}")
                return pd.DataFrame()
            finally:
                conn.close()

    # ── 取得資料庫路徑（除錯用）────────────────────────────────────────────────
    def info(self) -> dict:
        with self._lock:
            conn = self._connect()
            try:
                result = {}
                for table in SCHEMAS:
                    cur = conn.execute(f"SELECT COUNT(*) FROM {table}")
                    result[table] = cur.fetchone()[0]
                return result
            finally:
                conn.close()


# ── 全域單例 ──────────────────────────────────────────────────────────────────
DB = LocalDatabase()
