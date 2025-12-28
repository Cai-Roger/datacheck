import streamlit as st
import pandas as pd
import time
from io import BytesIO
from datetime import datetime
from zoneinfo import ZoneInfo
from pathlib import Path

from config import APP_NAME, APP_VERSION, APP_FOOTER
from compare_core import (
    clean_header_name,
    build_key_map,
    count_duplicate_keys,
    diff_directional,
    build_column_diff,
)

# =========================================================
# Page config（一定要第一個）
# =========================================================
st.set_page_config(
    page_title=APP_NAME,
    layout="wide",
)

# =========================================================
# 常數設定
# =========================================================
SESSION_TIMEOUT_SECONDS = 30 * 60
WARNING_SECONDS = 5 * 60

DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)

USAGE_XLSX = DATA_DIR / "usage.xlsx"

# =========================================================
# 工具
# =========================================================
def now_tw():
    return datetime.now(ZoneInfo("Asia/Taipei"))

def gen_download_filename(base):
    ts = now_tw().strftime("%Y%m%d_%H%M%S")
    return f"{base}_{ts}.xlsx"

# =========================================================
# 🔥 全欄位文字清洗（重點）
# =========================================================
def normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    1. 移除 ␣（U+2423）
    2. 移除 CR/LF/TAB
    3. NaN → 真正空白（Excel 不顯示 <NaN>）
    """
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = (
                df[col]
                .astype(object)
                .str.replace("\u2423", " ", regex=False)  # ␣
                .str.replace("\r", " ", regex=False)
                .str.replace("\n", " ", regex=False)
                .str.replace("\t", " ", regex=False)
                .str.strip()
            )
    return df.where(pd.notna(df), None)

# =========================================================
# 系統累積比對次數（持久化）
# =========================================================
def get_total_compare():
    if not USAGE_XLSX.exists():
        return 0
    try:
        df = pd.read_excel(USAGE_XLSX)
        return int(df.loc[0, "total"])
    except Exception:
        return 0

def bump_total_compare():
    n = get_total_compare() + 1
    df = pd.DataFrame([{
        "total": n,
        "updated": now_tw().strftime("%Y-%m-%d %H:%M:%S"),
        "version": APP_VERSION,
    }])
    df.to_excel(USAGE_XLSX, index=False)
    return n

# =========================================================
# 登入檢查
# =========================================================
def check_login():
    now = time.time()
    st.session_state.setdefault("auth", False)
    st.session_state.setdefault("last_active", now)
    st.session_state.setdefault("session_count", 0)

    if st.session_state.auth:
        if now - st.session_state.last_active > SESSION_TIMEOUT_SECONDS:
            st.session_state.auth = False
            return False
        return True

    st.title("🔐 系統登入")
    pwd = st.text_input("請輸入密碼", type="password")

    if st.button("登入"):
        if pwd == st.secrets["auth"]["password"]:
            st.session_state.auth = True
            st.session_state.last_active = now
            st.session_state.session_count = 0
            st.stop()
        else:
            st.error("密碼錯誤")

    return False

if not check_login():
    st.stop()

# =========================================================
# Sidebar
# =========================================================
with st.sidebar:
    st.markdown("### 🟢 登入狀態")
    st.caption(f"版本：{APP_VERSION}")
    st.caption(f"📊 系統累積比對次數：{get_total_compare()}")
    st.caption(f"🔁 本次登入比對次數：{st.session_state.session_count}")

    if st.button("🔁 延長登入"):
        st.session_state.last_active = time.time()

    if st.button("🔓 登出"):
        st.session_state.auth = False
        st.stop()

# =========================================================
# 主畫面
# =========================================================
st.title(f"Excel 比對程式（Web {APP_VERSION}）")

col1, col2 = st.columns(2)
with col1:
    file_a = st.file_uploader("📤 上傳 Excel A", type=["xlsx"])
with col2:
    file_b = st.file_uploader("📤 上傳 Excel B", type=["xlsx"])

if not file_a or not file_b:
    st.info("請上傳兩份 Excel")
    st.stop()

df_a = pd.read_excel(file_a)
df_b = pd.read_excel(file_b)

st.success(f"Excel A：{len(df_a)} 筆 ｜ Excel B：{len(df_b)} 筆")

# Key
st.subheader("🔑 Key 欄位")
cols = list(df_a.columns)
default_keys = [c for c in cols if clean_header_name(c) in {"PLNNR", "VORNR"}] or cols[:2]

keys = st.multiselect("選擇 Key（可多選）", cols, default=default_keys)
if not keys:
    st.stop()

if st.button("🟢 開始差異比對 🟢", type="primary"):
    st.session_state.session_count += 1
    total_count = bump_total_compare()
    st.session_state.last_active = time.time()

    with st.spinner("比對中..."):
        t0 = time.time()

        ka = [df_a.columns.get_loc(k) for k in keys]
        kb = [df_b.columns.get_loc(k) for k in keys]

        map_a = build_key_map(df_a, ka)
        map_b = build_key_map(df_b, kb)

        dup_a = count_duplicate_keys(df_a, ka)
        dup_b = count_duplicate_keys(df_b, kb)

        df_col = build_column_diff(df_a, df_b)

        a_rows, *_ = diff_directional(df_a, df_b, map_a, map_b, ka, "A", "B")
        b_rows, *_ = diff_directional(df_b, df_a, map_b, map_a, kb, "B", "A")

        headers = [f"KEY_{i+1}" for i in range(len(keys))] + ["差異欄位", "A值", "B值", "來源"]

        df_a2b = normalize_dataframe(pd.DataFrame(a_rows, columns=headers))
        df_b2a = normalize_dataframe(pd.DataFrame(b_rows, columns=headers))

        summary = pd.DataFrame([
            ["Key", ", ".join(keys), "", "", ""],
            ["A 重複", dup_a, "", "", ""],
            ["B 重複", dup_b, "", "", ""],
            ["A→B 差異", len(df_a2b), "", "", ""],
            ["B→A 差異", len(df_b2a), "", "", ""],
            ["系統累積比對", total_count, "", "", ""],
            ["本次登入比對", st.session_state.session_count, "", "", ""],
        ], columns=["項目", "值1", "值2", "值3", "值4"])

        out = BytesIO()
        with pd.ExcelWriter(out, engine="xlsxwriter") as w:
            summary.to_excel(w, "Summary", index=False)
            normalize_dataframe(df_col).to_excel(w, "ColumnDiff", index=False)
            df_a2b.to_excel(w, "A_to_B", index=False)
            df_b2a.to_excel(w, "B_to_A", index=False)

        cost = round(time.time() - t0, 2)

    st.success(f"比對完成（耗時 {cost} 秒）")

    st.download_button(
        "📥 下載差異比對結果",
        out.getvalue(),
        file_name=gen_download_filename("Excel差異比對結果"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# =========================================================
# Footer
# =========================================================
st.markdown(
    f"<div style='text-align:center;color:#666;border-top:1px solid #eee;padding:10px'>{APP_FOOTER} {APP_VERSION}</div>",
    unsafe_allow_html=True,
)
