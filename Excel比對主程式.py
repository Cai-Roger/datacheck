import streamlit as st
import pandas as pd
import time
from io import BytesIO
from datetime import datetime
from zoneinfo import ZoneInfo
from email.message import EmailMessage
import smtplib
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
FEEDBACK_XLSX = DATA_DIR / "feedback.xlsx"

# =========================================================
# 工具
# =========================================================
def now_tw():
    return datetime.now(ZoneInfo("Asia/Taipei"))

def gen_download_filename(base):
    ts = now_tw().strftime("%Y%m%d_%H%M%S")
    return f"{base}_{ts}.xlsx"

def empty_str_to_none(df: pd.DataFrame) -> pd.DataFrame:
    """避免 Excel 空白格出現底線"""
    return df.replace("", None)

# =========================================================
# 累積比對次數（持久化）
# =========================================================
def get_total_compare():
    if not USAGE_XLSX.exists():
        return 0
    try:
        df = pd.read_excel(USAGE_XLSX)
        return int(df.loc[0, "total_compare"])
    except Exception:
        return 0

def bump_total_compare():
    n = get_total_compare() + 1
    pd.DataFrame([{
        "total_compare": n,
        "updated_time": now_tw().strftime("%Y-%m-%d %H:%M:%S"),
        "app_version": APP_VERSION,
    }]).to_excel(USAGE_XLSX, index=False, engine="openpyxl")
    return n

# =========================================================
# 登入檢查
# =========================================================
def check_login():
    now = time.time()

    st.session_state.setdefault("authenticated", False)
    st.session_state.setdefault("last_active_ts", now)
    st.session_state.setdefault("warned", False)
    st.session_state.setdefault("compare_count_session", 0)

    if st.session_state.authenticated:
        if now - st.session_state.last_active_ts > SESSION_TIMEOUT_SECONDS:
            st.session_state.authenticated = False
            return False
        return True

    st.title("🔐 系統登入")
    pwd = st.text_input("請輸入系統密碼", type="password")

    if st.button("登入"):
        if pwd == st.secrets["auth"]["password"]:
            st.session_state.authenticated = True
            st.session_state.last_active_ts = now
            st.session_state.warned = False
            st.session_state.compare_count_session = 0
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
    st.caption(f"🔁 本次登入比對次數：{st.session_state.compare_count_session}")

    remaining = SESSION_TIMEOUT_SECONDS - (time.time() - st.session_state.last_active_ts)
    if remaining <= WARNING_SECONDS and remaining > 0 and not st.session_state.warned:
        st.warning("⚠️ 登入即將逾時，請延長登入")
        st.session_state.warned = True

    if remaining <= 0:
        st.session_state.authenticated = False
        st.stop()

    if st.button("🔁 延長登入"):
        st.session_state.last_active_ts = time.time()
        st.session_state.warned = False

    if st.button("🔓 登出"):
        st.session_state.authenticated = False
        st.stop()

# =========================================================
# 主畫面
# =========================================================
st.title(f"Excel 比對程式（Web {APP_VERSION}）")

st.markdown("""
### 使用說明
1. 上傳 Excel A、Excel B  
2. 選擇 Key 欄位  
3. 點擊「開始差異比對」
""")

# =========================================================
# 上傳檔案
# =========================================================
col1, col2 = st.columns(2)
with col1:
    file_a = st.file_uploader("📤 上傳 Excel A", type="xlsx")
with col2:
    file_b = st.file_uploader("📤 上傳 Excel B", type="xlsx")

if not file_a or not file_b:
    st.info("請先上傳兩份 Excel")
    st.stop()

df_a = pd.read_excel(file_a)
df_b = pd.read_excel(file_b)

st.success(f"Excel A：{len(df_a)} 筆 ｜ Excel B：{len(df_b)} 筆")

# =========================================================
# Key 設定
# =========================================================
st.subheader("🔑 Key 欄位設定")

cols = list(df_a.columns)
default_keys = [c for c in cols if clean_header_name(c) in {"PLNNR", "VORNR"}]
if not default_keys:
    default_keys = cols[:2]

selected_keys = st.multiselect(
    "選擇 Key 欄位（可多選）",
    cols,
    default=default_keys,
)

if not selected_keys:
    st.stop()

st.markdown("---")
start_compare = st.button("🟢 開始差異比對 🟢", type="primary")

if not start_compare:
    st.stop()

# =========================================================
# ✅ 計次（只在這裡 +1）
# =========================================================
st.session_state.compare_count_session += 1
total_now = bump_total_compare()
st.session_state.last_active_ts = time.time()
st.session_state.warned = False

# =========================================================
# 比對執行
# =========================================================
with st.spinner("資料比對中..."):
    t0 = time.time()

    key_cols_a = [df_a.columns.get_loc(k) for k in selected_keys]
    key_cols_b = [df_b.columns.get_loc(k) for k in selected_keys]

    map_a = build_key_map(df_a, key_cols_a)
    map_b = build_key_map(df_b, key_cols_b)

    dup_a = count_duplicate_keys(df_a, key_cols_a)
    dup_b = count_duplicate_keys(df_b, key_cols_b)

    df_col_diff = build_column_diff(df_a, df_b)

    a_rows, *_ = diff_directional(df_a, df_b, map_a, map_b, key_cols_a, "A", "B")
    b_rows, *_ = diff_directional(df_b, df_a, map_b, map_a, key_cols_b, "B", "A")

    key_headers = [f"KEY_{i+1}" for i in range(len(selected_keys))]
    headers = key_headers + ["差異欄位", "A值", "B值", "差異來源"]

    df_a_to_b = pd.DataFrame(a_rows, columns=headers)
    df_b_to_a = pd.DataFrame(b_rows, columns=headers) if b_rows else pd.DataFrame(columns=headers)

    df_summary = pd.DataFrame([
        ["Key 欄位", ", ".join(selected_keys), "", "", ""],
        ["A 筆數", len(df_a), "", "", ""],
        ["B 筆數", len(df_b), "", "", ""],
        ["A 重複 Key", dup_a, "", "", ""],
        ["B 重複 Key", dup_b, "", "", ""],
        ["系統累積比對次數", total_now, "", "", ""],
        ["本次登入比對次數", st.session_state.compare_count_session, "", "", ""],
        ["比對耗時(秒)", round(time.time() - t0, 2), "", "", ""],
    ], columns=["項目", "值1", "值2", "值3", "值4"])

    # 🔑 移除空白底線
    df_summary = empty_str_to_none(df_summary)
    df_col_diff = empty_str_to_none(df_col_diff)
    df_a_to_b = empty_str_to_none(df_a_to_b)
    df_b_to_a = empty_str_to_none(df_b_to_a)

    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_summary.to_excel(writer, "Summary", index=False)
        df_col_diff.to_excel(writer, "ColumnDiff", index=False)
        df_a_to_b.to_excel(writer, "A_to_B", index=False)
        df_b_to_a.to_excel(writer, "B_to_A", index=False)

st.success("✅ 比對完成")

st.download_button(
    "📥 下載差異比對結果 Excel",
    data=output.getvalue(),
    file_name=gen_download_filename("Excel差異比對結果"),
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

# =========================================================
# Footer
# =========================================================
st.markdown(
    f"""
    <div style="margin-top:40px;text-align:center;font-size:13px;color:#666;border-top:1px solid #e0e0e0;">
        {APP_FOOTER} {APP_VERSION}
    </div>
    """,
    unsafe_allow_html=True,
)
